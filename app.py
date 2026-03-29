import io
import pandas as pd
import streamlit as st
import networkx as nx
import matplotlib.pyplot as plt

st.set_page_config(page_title="Network Topology", layout="wide")

st.title("Network Topology Viewer")

uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])

if uploaded_file is None:
    st.stop()

excel = pd.ExcelFile(uploaded_file)

sheet = st.selectbox("Select room", excel.sheet_names)

df = pd.read_excel(uploaded_file, sheet_name=sheet)

df.columns = [str(c).lower().strip().replace(" ", "_") for c in df.columns]

possible_switch = [c for c in df.columns if "switch" in c and "port" not in c]
possible_switch_port = [c for c in df.columns if "switch_port" in c or "switchport" in c]
possible_patch = [c for c in df.columns if "patch" in c]
possible_port = [c for c in df.columns if "port" in c and "switch" not in c]

if not possible_switch or not possible_patch:
    st.error("Não consegui identificar as colunas de switch ou patch panel.")
    st.write("Colunas detectadas:", df.columns)
    st.stop()

switch_col = possible_switch[0]
patch_col = possible_patch[0]

connections = df[[patch_col, switch_col]].dropna()

connections.columns = ["patch", "switch"]

G = nx.Graph()

for _, row in connections.iterrows():
    patch = f"PATCH: {row['patch']}"
    switch = f"SWITCH: {row['switch']}"
    G.add_node(patch, type="patch")
    G.add_node(switch, type="switch")
    G.add_edge(patch, switch)

fig = plt.figure(figsize=(14,8))

pos = nx.kamada_kawai_layout(G)

node_colors = []
for node in G.nodes():
    if "PATCH" in node:
        node_colors.append("#3b82f6")
    else:
        node_colors.append("#22c55e")

nx.draw(
    G,
    pos,
    with_labels=True,
    node_size=3000,
    node_color=node_colors,
    font_size=8
)

plt.title("Patch Panels ↔ Switches")

st.pyplot(fig)

st.subheader("Tabela de conexões")

st.dataframe(connections)
