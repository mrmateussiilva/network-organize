import io
import re
from typing import Dict, List, Optional, Tuple

import matplotlib.pyplot as plt
import networkx as nx
import pandas as pd
import streamlit as st


st.set_page_config(page_title="Network Mapping", layout="wide")


def clean_text(value) -> str:
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_colname(col: str) -> str:
    col = clean_text(col).lower()
    col = re.sub(r"\s+", "_", col)
    col = col.replace("/", "_")
    col = col.replace("-", "_")
    return col


def normalize_df_columns(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [normalize_colname(c) for c in df.columns]
    return df


def parse_sheet_name(sheet_name: str) -> Tuple[str, str]:
    parts = [p.strip() for p in str(sheet_name).split("-")]
    if len(parts) >= 2:
        return parts[0], parts[1]
    return "", str(sheet_name)


def find_col(columns: List[str], aliases: List[str]) -> Optional[str]:
    for alias in aliases:
        if alias in columns:
            return alias
    for alias in aliases:
        for col in columns:
            if alias in col:
                return col
    return None


def is_active(value: str) -> bool:
    value = clean_text(value).lower()
    return value in {"yes", "y", "true", "1", "active", "up", "ok", "used"}


@st.cache_data(show_spinner=False)
def read_excel_file(file_bytes: bytes) -> Dict[str, pd.DataFrame]:
    excel = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets = {}
    for sheet_name in excel.sheet_names:
        try:
            raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name)
            sheets[sheet_name] = raw
        except Exception:
            continue
    return sheets


def normalize_sheet(raw_df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    df = normalize_df_columns(raw_df)
    deck_from_sheet, room_from_sheet = parse_sheet_name(sheet_name)

    deck_col = find_col(df.columns.tolist(), ["deck"])
    room_col = find_col(df.columns.tolist(), ["room", "location", "sala"])
    rack_source_col = find_col(df.columns.tolist(), ["rack", "rack1", "source_rack"])
    patch_panel_col = find_col(df.columns.tolist(), ["optic_patch_panel", "patch_panel", "panel", "odf"])
    patch_port_col = find_col(df.columns.tolist(), ["port", "patch_port", "panel_port"])
    rack_target_col = find_col(df.columns.tolist(), ["rack2", "target_rack", "destination_rack"])
    switch_col = find_col(df.columns.tolist(), ["switch", "switch_name"])
    switch_port_col = find_col(df.columns.tolist(), ["switch_port", "switchport", "port_switch"])
    active_col = find_col(df.columns.tolist(), ["active", "status"])
    observation_col = find_col(df.columns.tolist(), ["observation", "obs", "type", "description"])

    out = pd.DataFrame()
    out["sheet"] = [sheet_name] * len(df)
    out["deck"] = df[deck_col].map(clean_text) if deck_col else deck_from_sheet
    out["room"] = df[room_col].map(clean_text) if room_col else room_from_sheet
    out["rack_source"] = df[rack_source_col].map(clean_text) if rack_source_col else ""
    out["patch_panel"] = df[patch_panel_col].map(clean_text) if patch_panel_col else ""
    out["patch_port"] = df[patch_port_col].map(clean_text) if patch_port_col else ""
    out["rack_target"] = df[rack_target_col].map(clean_text) if rack_target_col else ""
    out["switch_name"] = df[switch_col].map(clean_text) if switch_col else ""
    out["switch_port"] = df[switch_port_col].map(clean_text) if switch_port_col else ""
    out["active"] = df[active_col].map(clean_text) if active_col else ""
    out["observation"] = df[observation_col].map(clean_text) if observation_col else ""

    mask = (
        out["patch_panel"].ne("")
        | out["patch_port"].ne("")
        | out["switch_name"].ne("")
        | out["switch_port"].ne("")
    )
    out = out.loc[mask].copy()
    out = out[~out["sheet"].str.contains("graf", case=False, na=False)].copy()
    return out.reset_index(drop=True)


@st.cache_data(show_spinner=False)
def build_connection_table(file_bytes: bytes) -> pd.DataFrame:
    sheets = read_excel_file(file_bytes)
    normalized = []
    for sheet_name, raw_df in sheets.items():
        parsed = normalize_sheet(raw_df, sheet_name)
        if not parsed.empty:
            normalized.append(parsed)

    if not normalized:
        return pd.DataFrame(
            columns=[
                "sheet",
                "deck",
                "room",
                "rack_source",
                "patch_panel",
                "patch_port",
                "rack_target",
                "switch_name",
                "switch_port",
                "active",
                "observation",
            ]
        )
    return pd.concat(normalized, ignore_index=True)


def patch_node(panel: str, port: str) -> str:
    return f"{panel} / Porta {port}"


def switch_node(sw: str, port: str) -> str:
    return f"{sw} / Porta {port}"


def build_graph(df: pd.DataFrame, include_inactive: bool) -> Tuple[nx.Graph, Dict[str, str]]:
    G = nx.Graph()
    color_map = {}

    work = df.copy()
    if not include_inactive:
        work = work[work["active"].map(is_active)]

    for _, row in work.iterrows():
        panel = row["patch_panel"] or "PatchPanel?"
        patch_port = row["patch_port"] or "?"
        switch_name = row["switch_name"] or "Switch?"
        switch_port = row["switch_port"] or "?"

        patch_id = patch_node(panel, patch_port)
        switch_id = switch_node(switch_name, switch_port)

        color_map[patch_id] = "#60a5fa"
        color_map[switch_id] = "#34d399"

        edge_color = "#22c55e" if is_active(row["active"]) else "#94a3b8"
        width = 3 if "SFP" in row["observation"].upper() else 1.5

        G.add_node(patch_id)
        G.add_node(switch_id)
        G.add_edge(
            patch_id,
            switch_id,
            color=edge_color,
            width=width,
            title=(
                f"Deck: {row['deck']} | Sala: {row['room']} | "
                f"{panel} porta {patch_port} -> {switch_name} porta {switch_port} | "
                f"Ativo: {row['active']} | Obs: {row['observation']}"
            ),
        )

    return G, color_map


def draw_graph(G: nx.Graph, color_map: Dict[str, str]):
    fig = plt.figure(figsize=(16, 10))
    pos = nx.spring_layout(G, seed=42, k=0.9)

    node_colors = [color_map.get(node, "#cbd5e1") for node in G.nodes()]
    edge_colors = [attrs.get("color", "#94a3b8") for _, _, attrs in G.edges(data=True)]
    edge_widths = [attrs.get("width", 1.5) for _, _, attrs in G.edges(data=True)]

    nx.draw_networkx_nodes(G, pos, node_color=node_colors, node_size=2200)
    nx.draw_networkx_edges(G, pos, edge_color=edge_colors, width=edge_widths)
    nx.draw_networkx_labels(G, pos, font_size=8)

    plt.axis("off")
    plt.tight_layout()
    return fig


st.title("Network Mapping")
st.caption("Patch panel / porta  ->  switch / porta")

uploaded_file = st.file_uploader("Envie a planilha Excel", type=["xlsx", "xls"])

if uploaded_file is None:
    st.info("Faça o upload da planilha para gerar o grafo.")
    st.stop()

file_bytes = uploaded_file.getvalue()
df = build_connection_table(file_bytes)

if df.empty:
    st.error("Não consegui extrair conexões da planilha.")
    st.stop()

st.sidebar.header("Filtros")

decks = sorted([x for x in df["deck"].dropna().unique().tolist() if x])
rooms = sorted([x for x in df["room"].dropna().unique().tolist() if x])
switches = sorted([x for x in df["switch_name"].dropna().unique().tolist() if x])
patch_panels = sorted([x for x in df["patch_panel"].dropna().unique().tolist() if x])

selected_deck = st.sidebar.selectbox("Deck", ["Todos"] + decks)
selected_room = st.sidebar.selectbox("Sala", ["Todas"] + rooms)
selected_switch = st.sidebar.selectbox("Switch", ["Todos"] + switches)
selected_patch = st.sidebar.selectbox("Patch Panel", ["Todos"] + patch_panels)
include_inactive = st.sidebar.checkbox("Incluir inativas", value=True)
only_sfp = st.sidebar.checkbox("Somente SFP", value=False)

filtered = df.copy()

if selected_deck != "Todos":
    filtered = filtered[filtered["deck"] == selected_deck]

if selected_room != "Todas":
    filtered = filtered[filtered["room"] == selected_room]

if selected_switch != "Todos":
    filtered = filtered[filtered["switch_name"] == selected_switch]

if selected_patch != "Todos":
    filtered = filtered[filtered["patch_panel"] == selected_patch]

if only_sfp:
    filtered = filtered[filtered["observation"].str.upper().str.contains("SFP", na=False)]

if not include_inactive:
    filtered = filtered[filtered["active"].map(is_active)]

if filtered.empty:
    st.warning("Nenhum resultado com os filtros atuais.")
    st.stop()

c1, c2, c3, c4 = st.columns(4)
c1.metric("Conexões", len(filtered))
c2.metric("Patch Panels", filtered["patch_panel"].replace("", pd.NA).dropna().nunique())
c3.metric("Switches", filtered["switch_name"].replace("", pd.NA).dropna().nunique())
c4.metric("Ativas", int(filtered["active"].map(is_active).sum()))

tab1, tab2, tab3 = st.tabs(["Grafo", "Tabela", "Leitura humana"])

with tab1:
    st.subheader("Grafo")
    G, color_map = build_graph(filtered, include_inactive=include_inactive)
    fig = draw_graph(G, color_map)
    st.pyplot(fig, use_container_width=True)

with tab2:
    st.subheader("Tabela normalizada")
    st.dataframe(filtered, use_container_width=True, hide_index=True)
    csv_data = filtered.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Baixar CSV normalizado",
        data=csv_data,
        file_name="network_mapping_normalized.csv",
        mime="text/csv",
    )

with tab3:
    st.subheader("Conexões")
    for _, row in filtered.head(200).iterrows():
        status = "ATIVA" if is_active(row["active"]) else "INATIVA"
        st.markdown(
            f"- **{row['patch_panel']} / porta {row['patch_port']}** -> "
            f"**{row['switch_name']} / porta {row['switch_port']}** "
            f"| Deck: `{row['deck']}` | Sala: `{row['room']}` | "
            f"Status: `{status}` | Obs: `{row['observation']}`"
        )
