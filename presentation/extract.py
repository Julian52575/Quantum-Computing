import json, sys, numpy as np, matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
nb = json.load(open(sys.argv[1])); out = sys.argv[2]
g = {}
for c in nb["cells"]:
    if c["cell_type"] != "code": continue
    src = "".join(c["source"])
    if "%pip" in src: continue
    exec(src, g)
    plt.close("all")
d = {}
dist = g["qaoa_run"]["distribution"]
top = sorted(dist.items(), key=lambda t: -t[1])[:15]
d["top"] = [{"orders": [i+1 for i,b in enumerate(x) if b], "p": p, "valid": g["is_valid"](x, g["CONFLICTS"]),
             "opt": g["is_valid"](x, g["CONFLICTS"]) and g["profit"](x, g["PROFITS"]) == g["BEST_PROFIT"]} for x,p in top]
fv = [r["f"] for r in g["ranking"]]
bins = list(range(min(fv), max(fv)+1))
d["hist"] = {"bins": bins, "valid": [sum(1 for r in g["ranking"] if r["valid"] and r["f"]==b) for b in bins],
             "invalid": [sum(1 for r in g["ranking"] if not r["valid"] and r["f"]==b) for b in bins]}
clean = lambda rows: [{k: (float(v) if isinstance(v,(np.floating,float)) else v) for k,v in r.items()} for r in rows]
d["layers"] = clean(g["layer_rows"]); d["penalty"] = clean(g["penalty_rows"]); d["size"] = clean(g["size_rows"])
d["random"] = g["RANDOM_BASELINE"]; d["qaoa"] = g["qaoa_metrics"]; d["noisy"] = g["noisy_metrics"]
d["params"] = list(map(float, g["qaoa_run"]["params"]))
d["transpile"] = {"logical_depth": g["logical"].depth(), "logical_2q": g["two_qubit_count"](g["logical"]),
                  "isa_depth": g["isa_circuit"].depth(), "isa_2q": g["two_qubit_count"](g["isa_circuit"])}
json.dump(d, open(f"{out}/data.json","w"), indent=1, default=float)
# Figures
for name, sel in [("graph", None), ("graph_opt", [0,3,4,7])]:
    fig = plt.figure(figsize=(6.4, 2.6)); g["draw_instance"](g["PROFITS"], g["CONFLICTS"], selected=sel, title=" "); plt.gca().margins(0.1, 0.3)
    fig.savefig(f"{out}/{name}.png", dpi=250, bbox_inches="tight", transparent=True); plt.close(fig)
circ = g["qaoa_run"]["circuit"]
from qiskit_optimization.minimum_eigensolvers import QAOA
from qiskit.circuit.library import QAOAAnsatz
op,_ = g["build_qubo"](g["PROFITS"], g["CONFLICTS"], g["A"], scale=g["A"]).to_ising()
one = QAOAAnsatz(op, reps=1).decompose(reps=2)
fig = one.draw("mpl", fold=-1, scale=0.55); fig.savefig(f"{out}/circuit.png", dpi=220, bbox_inches="tight", facecolor="white"); plt.close(fig)
print(json.dumps({k: d[k] for k in ["random","qaoa","noisy","transpile"]}, default=float))
