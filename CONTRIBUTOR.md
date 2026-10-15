# Contributing

## Project goal

This project's goal is to optimize a Python algorithm using quantum programming
(via Qiskit), comparing its performance against the classical implementation.

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
