# Quantum-Computing

Order selection with QAOA: [`order_selection_qaoa.ipynb`](order_selection_qaoa.ipynb)
models an order-selection problem as a QUBO and solves it with Qiskit (QAOA on ideal and
noisy emulators, optionally IBM hardware), compared against classical solvers.

The notebook **is** the project. Everything in this repo exists to run it and track it in git.

See [CONTRIBUTOR.md](CONTRIBUTOR.md) for the contribution workflow.

## Quick start (Nix)

Install [Nix](https://determinate.systems/install/) with flakes enabled, then:

```bash
nix develop                                  # first run builds the env (uv sync + venv)
jupyter lab order_selection_qaoa.ipynb       # open the notebook
```

Entering the shell:

- provides pinned `python3.12` and `uv` from [nixpkgs 26.05](https://github.com/NixOS/nixpkgs/releases/tag/26.05)
- installs the locked Python dependencies into `.venv/` (`uv sync --frozen`) and activates it
- enables [nbdime](https://nbdime.readthedocs.io/) as git's diff/merge tool for notebooks

Editors (VS Code, PyCharm) can point at `.venv/bin/python` as the kernel.

## Without Nix

Requires Python 3.12 and [uv](https://docs.astral.sh/uv/):

```bash
uv sync
uv run jupyter lab order_selection_qaoa.ipynb
```

## Common commands

```bash
jupyter nbconvert --to notebook --execute --inplace order_selection_qaoa.ipynb  # run all cells headless
nbdiff-web                                   # visual diff of uncommitted notebook changes
ruff check order_selection_qaoa.ipynb        # lint the notebook's code cells
uv add <pkg>                                 # add a dependency (updates uv.lock)
```

## IBM Quantum hardware (optional)

The hardware section is off by default (`RUN_ON_HARDWARE = False`). To use it, save an
IBM Quantum account once with `QiskitRuntimeService.save_account(...)`, flip the flag, and
run the cell. Results are cached to `results/hardware_counts.json`; commit that file so
others can reproduce the comparison without spending hardware time.

## Stack

- **Python 3.12**, **uv** (dependencies + lockfile), **Nix flake** (dev shell)
- **Qiskit**, **qiskit-optimization**, **qiskit-aer**, **qiskit-ibm-runtime**
- **JupyterLab**, **nbdime**, **ruff**
