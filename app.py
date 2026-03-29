
import streamlit as st
import pandas as pd
import networkx as nx
import matplotlib.pyplot as plt

st.set_page_config(page_title="Network Mapping", layout="wide")

st.title("Network Topology Viewer")

uploaded_file = st.file_uploader("Upload Excel", type=["xlsx"])

if uploaded_file is None:
    st.stop()

# --- FIX: header=4 because real header is on line 5 in your Excel ---
excel = pd.ExcelFile(uploaded_file)

sheet = st.selectbox("Select room", excel.sheet_names)

df = pd.read_excel(uploaded_file, sheet_name=sheet, header=4)

# normalize column names
df.columns = [str(c).lower().strip().replace(" ", "_") for c in df.columns]

# try to detect columns
patch_candidates = [c for c in df.columns if "patch" in c]
switch_candidates = [c for c in df.columns if "switch" in c and "port" not in c]

if not patch_candidates or not switch_candidates:
    st.error("Could not detect Patch Panel or Switch columns.")
    st.write("Detected columns:", df.columns.tolist())
    st.stop()

patch_col = patch_candidates[0]
switch_col = switch_candidates[0]

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

st.subheader("Connections table")
st.dataframe(connections)
