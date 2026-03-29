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


def detect_header_row(raw_df: pd.DataFrame) -> Optional[int]:
    keywords = ["deck", "rack", "switch", "port", "active", "observation"]
    max_rows = min(len(raw_df), 20)

    for idx in range(max_rows):
        row_values = [clean_text(v).lower() for v in raw_df.iloc[idx].tolist()]
        score = 0
        for keyword in keywords:
            if any(keyword in cell for cell in row_values):
                score += 1
        if score >= 4:
            return idx

    return None


@st.cache_data(show_spinner=False)
def read_excel_file(file_bytes: bytes) -> Dict[str, pd.DataFrame]:
    excel = pd.ExcelFile(io.BytesIO(file_bytes))
    sheets = {}

    for sheet_name in excel.sheet_names:
        try:
            raw = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name, header=None)
            sheets[sheet_name] = raw
        except Exception:
            continue

    return sheets


def normalize_sheet(raw_df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    deck_from_sheet, room_from_sheet = parse_sheet_name(sheet_name)

    header_row = detect_header_row(raw_df)
    if header_row is None:
        return pd.DataFrame()

    header = [clean_text(v) for v in raw_df.iloc[header_row].tolist()]
    data = raw_df.iloc[header_row + 1:].copy()
    data.columns = header
    data = normalize_df_columns(data)

    deck_col = find_col(data.columns.tolist(), ["deck"])
    rack_source_col = find_col(data.columns.tolist(), ["rack"])
    patch_panel_col = find_col(data.columns.tolist(), ["patch_panel", "optic_patch_painel", "panel"])
    patch_port_col = find_col(data.columns.tolist(), ["port"])
    rack_target_col = find_col(data.columns.tolist(), ["rack2"])
    switch_col = find_col(data.columns.tolist(), ["switch"])
    switch_port_col = find_col(data.columns.tolist(), ["switch_port"])
    active_col = find_col(data.columns.tolist(), ["active"])
    observation_col = find_col(data.columns.tolist(), ["observation"])

    out = pd.DataFrame()
    out["sheet"] = [sheet_name] * len(data)
    out["deck"] = data[deck_col].map(clean_text) if deck_col else deck_from_sheet
    out["room"] = room_from_sheet
    out["rack_source"] = data[rack_source_col].map(clean_text) if rack_source_col else ""
    out["patch_panel"] = data[patch_panel_col].map(clean_text) if patch_panel_col else ""
    out["patch_port"] = data[patch_port_col].map(clean_text) if patch_port_col else ""
    out["rack_target"] = data[rack_target_col].map(clean_text) if rack_target_col else ""
    out["switch_name"] = data[switch_col].map(clean_text) if switch_col else ""
    out["switch_port"] = data[switch_port_col].map(clean_text) if switch_port_col else ""
    out["active"] = data[active_col].map(clean_text) if active_col else ""
    out["observation"] = data[observation_col].map(clean_text) if observation_col else ""

    mask = (
        out["patch_panel"].ne("")
        | out["patch_port"].ne("")
        | out["switch_name"].ne("")
        | out["switch_port"].ne("")
    )

    out = out.loc[mask].copy()
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
        return pd.DataFrame()

    return pd.concat(normalized, ignore_index=True)


def patch_node(panel: str, port: str) -> str:
    return f"{panel} / Porta {port}"


def switch_node(sw: str, port: str) -> str:
    return f"{sw} / Porta {port}"


def build_graph(df: pd.DataFrame):
    G = nx.Graph()
    color_map = {}

    for _, row in df.iterrows():
        panel = row["patch_panel"]
        patch_port = row["patch_port"]
        switch_name = row["switch_name"]
        switch_port = row["switch_port"]

        if not panel or not patch_port or not switch_name or not switch_port:
            continue

        patch_id = patch_node(panel, patch_port)
        switch_id = switch_node(switch_name, switch_port)

        color_map[patch_id] = "#60a5fa"
        color_map[switch_id] = "#34d399"

        G.add_edge(patch_id, switch_id)

    return G, color_map


def draw_graph(G: nx.Graph, color_map: Dict[str, str]):
    fig = plt.figure(figsize=(16, 10))

    if len(G.nodes()) == 0:
        plt.text(0.5, 0.5, "Nenhuma conexão encontrada", ha="center")
        plt.axis("off")
        return fig

    pos = nx.spring_layout(G, seed=42)

    node_colors = [color_map.get(node, "#cbd5e1") for node in G.nodes()]

    nx.draw_networkx_nodes(G, pos, node_color=node_colors, node_size=2000)
    nx.draw_networkx_edges(G, pos)
    nx.draw_networkx_labels(G, pos, font_size=8)

    plt.axis("off")
    return fig


st.title("Network Mapping")

uploaded_file = st.file_uploader("Envie a planilha Excel", type=["xlsx"])

if uploaded_file is None:
    st.stop()

file_bytes = uploaded_file.getvalue()
df = build_connection_table(file_bytes)

if df.empty:
    st.error("Não consegui extrair conexões da planilha.")
    st.stop()

G, color_map = build_graph(df)

st.subheader("Grafo de conexões")

fig = draw_graph(G, color_map)
st.pyplot(fig, use_container_width=True)

st.subheader("Tabela")
st.dataframe(df)
