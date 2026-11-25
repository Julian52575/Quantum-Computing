# Figures for the WorkshopOrders deck. Usage: python make_assets.py (from this folder, in the project venv)
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import networkx as nx
from qiskit.circuit.library import QAOAAnsatz
from qiskit_optimization import QuadraticProgram

PROFITS = [5, 3, 4, 6, 2, 5, 3, 4]
PAIRS = [(1, 2), (1, 3), (2, 4), (3, 4), (3, 5), (4, 6), (5, 6), (5, 7), (6, 8), (7, 8)]
OPT = {1, 4, 5, 8}
# Ladder layout: the conflict graph is two rails (odd / even orders) joined by rungs.
POS = {1: (0, 1), 3: (1, 1), 5: (2, 1), 7: (3, 1), 2: (0, 0), 4: (1, 0), 6: (2, 0), 8: (3, 0)}

def graph(name, highlight):
    g = nx.Graph(PAIRS)
    fig, ax = plt.subplots(figsize=(6.4, 2.9))
    nx.draw_networkx_edges(g, POS, ax=ax, width=2.2, edge_color="#8D99AE")
    colors = ["#E76F51" if highlight and n in OPT else "#2A9D8F" for n in g.nodes]
    nx.draw_networkx_nodes(g, POS, ax=ax, node_size=1500, node_color=colors, edgecolors="white", linewidths=2)
    nx.draw_networkx_labels(g, POS, ax=ax, labels={n: f"{n}\np={PROFITS[n-1]}" for n in g.nodes},
                            font_color="white", font_weight="bold", font_size=10)
    ax.axis("off"); ax.margins(0.12, 0.25)
    fig.savefig(name, dpi=250, bbox_inches="tight", transparent=True); plt.close(fig)

graph("graph.png", False)
graph("graph_opt.png", True)

qp = QuadraticProgram()
for i in range(1, 9):
    qp.binary_var(f"x{i}")
qp.minimize(linear={f"x{i}": -p for i, p in enumerate(PROFITS, 1)}, quadratic={(f"x{i}", f"x{j}"): 7 for i, j in PAIRS})
op, _ = qp.to_ising()
fig = QAOAAnsatz(op, reps=1).decompose(reps=2).draw("mpl", fold=-1, scale=0.55)
fig.savefig("circuit.png", dpi=220, bbox_inches="tight", facecolor="white"); plt.close(fig)
