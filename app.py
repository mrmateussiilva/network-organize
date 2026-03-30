import io
import re
import tempfile
from typing import Dict, List, Optional, Tuple

import networkx as nx
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
from pyvis.network import Network


st.set_page_config(page_title="Network Mapping Graph", layout="wide")


# ============================================================
# Helpers
# ============================================================
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
    # Example: Deck M - L519
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


# ============================================================
# Excel parsing
# ============================================================
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
    """
    Converts a sheet to a common schema.
    The app tries to adapt to imperfect column names.
    """
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

    # Remove almost empty rows
    mask = (
        out["patch_panel"].ne("")
        | out["patch_port"].ne("")
        | out["switch_name"].ne("")
        | out["switch_port"].ne("")
    )
    out = out.loc[mask].copy()

    # Remove chart/helper sheets
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


# ============================================================
# Graph model
# ============================================================
def patch_node_id(panel: str, port: str) -> str:
    return f"PATCH::{panel}::PORT::{port}"


def switch_node_id(sw: str, port: str) -> str:
    return f"SWITCH::{sw}::PORT::{port}"


def patch_label(panel: str, port: str) -> str:
    return f"Patch Panel\n{panel}\nPorta {port}"


def switch_label(sw: str, port: str) -> str:
    return f"Switch\n{sw}\nPorta {port}"


def build_graph(df: pd.DataFrame, include_inactive: bool) -> nx.Graph:
    G = nx.Graph()

    work = df.copy()
    if not include_inactive:
        work = work[work["active"].map(is_active)]

    for _, row in work.iterrows():
        panel = row["patch_panel"] or "PatchPanel?"
        patch_port = row["patch_port"] or "?"
        switch_name = row["switch_name"] or "Switch?"
        switch_port = row["switch_port"] or "?"

        patch_id = patch_node_id(panel, patch_port)
        switch_id = switch_node_id(switch_name, switch_port)

        G.add_node(
            patch_id,
            label=patch_label(panel, patch_port),
            title=(
                f"Deck: {row['deck']}\n"
                f"Sala: {row['room']}\n"
                f"Rack origem: {row['rack_source']}\n"
                f"Patch panel: {panel}\n"
                f"Porta: {patch_port}"
            ),
            kind="patch",
            deck=row["deck"],
            room=row["room"],
            rack=row["rack_source"],
        )

        G.add_node(
            switch_id,
            label=switch_label(switch_name, switch_port),
            title=(
                f"Deck: {row['deck']}\n"
                f"Sala: {row['room']}\n"
                f"Rack destino: {row['rack_target']}\n"
                f"Switch: {switch_name}\n"
                f"Porta: {switch_port}\n"
                f"Ativo: {row['active']}\n"
                f"Obs: {row['observation']}"
            ),
            kind="switch",
            deck=row["deck"],
            room=row["room"],
            rack=row["rack_target"],
        )

        is_sfp = "SFP" in row["observation"].upper()
        active_value = row["active"]

        G.add_edge(
            patch_id,
            switch_id,
            title=(
                f"{panel} porta {patch_port}  →  {switch_name} porta {switch_port}\n"
                f"Ativo: {active_value}\n"
                f"Obs: {row['observation']}"
            ),
            is_active=is_active(active_value),
            is_sfp=is_sfp,
            room=row["room"],
            deck=row["deck"],
        )

    return G


def render_pyvis_graph(G: nx.Graph, height: str = "760px") -> str:
    net = Network(height=height, width="100%", bgcolor="#0b1020", font_color="#e5e7eb")
    net.barnes_hut(gravity=-30000, central_gravity=0.25, spring_length=140, spring_strength=0.02)

    for node_id, attrs in G.nodes(data=True):
        kind = attrs.get("kind", "generic")
        if kind == "patch":
            color = "#3b82f6"
            shape = "box"
        elif kind == "switch":
            color = "#10b981"
            shape = "ellipse"
        else:
            color = "#94a3b8"
            shape = "dot"

        net.add_node(
            node_id,
            label=attrs.get("label", node_id),
            title=attrs.get("title", ""),
            color=color,
            shape=shape,
        )

    for source, target, attrs in G.edges(data=True):
        if attrs.get("is_active"):
            color = "#22c55e"
        else:
            color = "#64748b"

        if attrs.get("is_sfp"):
            width = 4
            dashes = False
        else:
            width = 2
            dashes = True

        net.add_edge(
            source,
            target,
            title=attrs.get("title", ""),
            color=color,
            width=width,
            dashes=dashes,
        )

    net.set_options(
        """
        {
          "interaction": {
            "hover": true,
            "navigationButtons": true,
            "keyboard": true
          },
          "physics": {
            "enabled": true,
            "stabilization": {
              "enabled": true,
              "iterations": 500
            }
          }
        }
        """
    )

    with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as tmp_file:
        net.save_graph(tmp_file.name)
        with open(tmp_file.name, "r", encoding="utf-8") as f:
            html = f.read()
    return html


# ============================================================
# UI
# ============================================================
st.title("Network Mapping Graph")
st.caption("Mapeamento de patch panel ↔ switch em formato de grafo")

uploaded_file = st.file_uploader("Envie a planilha Excel", type=["xlsx", "xls"])

with st.expander("Como esse app funciona"):
    st.markdown(
        """
        - Cada **porta de patch panel** vira um **nó**.
        - Cada **porta de switch** vira um **nó**.
        - A **conexão entre elas** vira uma **aresta**.
        - Linhas **ativas** ficam em verde.
        - Conexões com **SFP** ficam mais grossas.
        """
    )

if uploaded_file is None:
    st.info("Envie uma planilha para gerar o mapeamento em grafo.")
    st.stop()

file_bytes = uploaded_file.getvalue()
df = build_connection_table(file_bytes)

if df.empty:
    st.error("Não consegui extrair conexões válidas da planilha.")
    st.stop()

# Sidebar filters
st.sidebar.header("Filtros")

decks = sorted([x for x in df["deck"].dropna().unique().tolist() if x])
rooms = sorted([x for x in df["room"].dropna().unique().tolist() if x])
racks = sorted(
    list(
        set(
            [x for x in df["rack_source"].dropna().unique().tolist() if x]
            + [x for x in df["rack_target"].dropna().unique().tolist() if x]
        )
    )
)
switches = sorted([x for x in df["switch_name"].dropna().unique().tolist() if x])
patch_panels = sorted([x for x in df["patch_panel"].dropna().unique().tolist() if x])

selected_deck = st.sidebar.selectbox("Deck", ["Todos"] + decks)
selected_room = st.sidebar.selectbox("Sala", ["Todas"] + rooms)
selected_rack = st.sidebar.selectbox("Rack", ["Todos"] + racks)
selected_switch = st.sidebar.selectbox("Switch", ["Todos"] + switches)
selected_patch = st.sidebar.selectbox("Patch Panel", ["Todos"] + patch_panels)
include_inactive = st.sidebar.checkbox("Incluir conexões inativas", value=True)
only_sfp = st.sidebar.checkbox("Somente SFP", value=False)

filtered = df.copy()

if selected_deck != "Todos":
    filtered = filtered[filtered["deck"] == selected_deck]

if selected_room != "Todas":
    filtered = filtered[filtered["room"] == selected_room]

if selected_rack != "Todos":
    filtered = filtered[(filtered["rack_source"] == selected_rack) | (filtered["rack_target"] == selected_rack)]

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

# Metrics
col1, col2, col3, col4 = st.columns(4)
col1.metric("Conexões", len(filtered))
col2.metric("Patch Panels", filtered["patch_panel"].replace("", pd.NA).dropna().nunique())
col3.metric("Switches", filtered["switch_name"].replace("", pd.NA).dropna().nunique())
col4.metric("Ativas", int(filtered["active"].map(is_active).sum()))

# Tabs
tab1, tab2, tab3 = st.tabs(["Grafo", "Tabela normalizada", "Resumo"])

with tab1:
    st.subheader("Grafo de conexões")
    graph = build_graph(filtered, include_inactive=include_inactive)
    html = render_pyvis_graph(graph)
    components.html(html, height=780, scrolling=True)

with tab2:
    st.subheader("Tabela de conexões")
    st.dataframe(filtered, use_container_width=True, hide_index=True)
    csv_data = filtered.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Baixar CSV normalizado",
        data=csv_data,
        file_name="network_mapping_normalized.csv",
        mime="text/csv",
    )

with tab3:
    st.subheader("Leitura humana")
    for _, row in filtered.head(100).iterrows():
        status = "ATIVA" if is_active(row["active"]) else "INATIVA"
        st.markdown(
            f"- **{row['patch_panel']} / porta {row['patch_port']}** → "
            f"**{row['switch_name']} / porta {row['switch_port']}** "
            f"| Sala: `{row['room']}` | Deck: `{row['deck']}` | Status: `{status}` | Obs: `{row['observation']}`"
        )

st.divider()
st.markdown(
    "**Dependências para rodar:** `streamlit pandas openpyxl networkx pyvis`  \n"
    "**Comando:** `streamlit run app.py`"
)
