# Quantum-Computing

See [CONTRIBUTOR.md](CONTRIBUTOR.md) for the contribution workflow.

## Stack

- **Python 3.12**
- **[uv](https://docs.astral.sh/uv/)** — dependency management, virtual envs, and lockfile
- **[Qiskit](https://www.ibm.com/quantum/qiskit)** — quantum computing library
- **[pytest](https://docs.pytest.org/)** — testing
- **[ruff](https://docs.astral.sh/ruff/)** — linting and formatting
- **[nix](https://determinate.systems/install/)** - _optional_ - exposes all the previous dependencies into a temporary shell

## Structure

```
src/quantum_computing/   # package source
tests/                   # tests
```

## Usage

```bash
uv sync                      # install dependencies
uv run quantum-computing                    # run the main entrypoint
uv run quantum-computing-benchmark          # run the app N times and report perf metrics
uv run quantum-computing-benchmark -n 500   # override the iteration count (default: 100)
uv run pytest                               # run tests
uv run ruff check .                         # lint
```

### Benchmark metrics

`uv run quantum-computing-benchmark` runs `main()` repeatedly and reports:

- **Startup time** — time to import `quantum_computing.main` (relevant once Qiskit imports land, since they can be heavy)
- **Wall-clock time** — total elapsed time across all iterations (`time.perf_counter()`)
- **CPU time** — user vs. system time (`resource.getrusage`)
- **Peak memory (RSS)** — maximum resident set size (`ru_maxrss`)
- **Throughput** — iterations per second
- **Latency distribution** — mean/median/p95/p99 across iterations, not just a single sample

## Nix dev shell

A `flake.nix` is provided for a reproducible shell that exposes `uv` and `python3.12`
preinstalled, pinned to [nixpkgs 26.05](https://github.com/NixOS/nixpkgs/releases/tag/26.05):

- `uv` 0.11.16
- `python3` 3.12.13

```bash
julian@RulianPC:~/Quantum-Computing$ uv --version    # not on path
Command 'uv' not found, but can be installed with:
sudo snap install astral-uv
julian@RulianPC:~/Quantum-Computing$ nix develop     # drop into the shell
julian@RulianPC:~/Quantum-Computing$ uv --version    # uv is available under the nix shell
uv 0.11.16 (x86_64-unknown-linux-gnu)
julian@RulianPC:~/Quantum-Computing$ which uv
/nix/store/.../uv-0.11.16/bin/uv
```
