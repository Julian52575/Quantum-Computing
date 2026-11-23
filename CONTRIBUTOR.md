# Contributing

## Project goal

This project's goal is to optimize a Python algorithm using quantum programming
(via Qiskit), comparing its performance against the classical implementation.
All of that work lives in `order_selection_qaoa.ipynb`.

## Working on the notebook

- Use the Nix shell (`nix develop`) so everyone runs the same Python and dependencies.
- Before committing, **restart the kernel and run all cells** so the committed outputs
  match the committed code (or run it headless, see the README).
- Review your change with `git diff` / `nbdiff-web` (nbdime), which shows cell-level diffs
  instead of raw JSON.
- Add dependencies with `uv add <pkg>` and commit the updated `uv.lock` with it.

## Commit messages

All commits must follow [Conventional Commits](https://www.conventionalcommits.org/):

```
<type>[optional scope]: <description>

[optional body]

[optional footer(s)]
```

Common types: `feat`, `fix`, `docs`, `style`, `refactor`, `perf`, `test`, `build`, `ci`, `chore`.

Example:

```
feat(benchmark): add latency percentiles to json output
```

## Branches and PRs

- Never commit directly to `main`.
- All work happens on a branch, merged into `main` through a pull request.
- Keep branches focused on a single change; prefer a small PR over a large one.
