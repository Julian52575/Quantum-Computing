"""Build the benchmark job summary: a comparison table against a baseline run,
or a plain dump of the current run when no baseline is available.
"""
import argparse
import json

METRICS: list[tuple[str, str, str]] = [
    ("startup_time_s", "Startup time", "s"),
    ("wall_clock_time_s", "Wall-clock time", "s"),
    ("cpu_user_time_s", "CPU time (user)", "s"),
    ("cpu_sys_time_s", "CPU time (sys)", "s"),
    ("peak_memory_mb", "Peak memory (RSS)", "MB"),
    ("throughput_ops_s", "Throughput", "ops/s"),
    ("latency_mean_ms", "Latency (mean)", "ms"),
    ("latency_median_ms", "Latency (median)", "ms"),
    ("latency_p95_ms", "Latency (p95)", "ms"),
    ("latency_p99_ms", "Latency (p99)", "ms"),
]


def commit_link(repo_url: str, sha: str) -> str:
    return f"[`{sha[:7]}`]({repo_url}/commit/{sha})"


def fmt(value: float | None, unit: str) -> str:
    return f"{value:.4f} {unit}" if value is not None else "n/a"


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--current", required=True)
    parser.add_argument("--baseline")
    parser.add_argument("--repo-url", required=True)
    args = parser.parse_args()

    with open(args.current) as f:
        current = json.load(f)

    lines = ["## Benchmark results", ""]
    lines.append(f"Current run: {commit_link(args.repo_url, current['commit'])} (`{current['ref']}`)")

    baseline = None
    if args.baseline:
        try:
            with open(args.baseline) as f:
                baseline = json.load(f)
        except FileNotFoundError:
            baseline = None

    if baseline is None:
        lines.append("")
        lines.append("_No cached baseline found to compare against — showing current run only._")
        lines.append("")
        lines.append("| Metric | Value |")
        lines.append("|---|---|")
        for key, label, unit in METRICS:
            lines.append(f"| {label} | {fmt(current[key], unit)} |")
        lines.append(f"| Iterations | {current['iterations']} |")
    else:
        lines.append(
            f"Baseline run: {commit_link(args.repo_url, baseline['commit'])} (`{baseline['ref']}`)"
        )
        lines.append("")
        lines.append("| Metric | Baseline | Current | Delta |")
        lines.append("|---|---|---|---|")
        for key, label, unit in METRICS:
            base_val = baseline[key]
            cur_val = current[key]
            if base_val is None or cur_val is None:
                lines.append(f"| {label} | {fmt(base_val, unit)} | {fmt(cur_val, unit)} | n/a |")
                continue
            delta = cur_val - base_val
            sign = "+" if delta >= 0 else ""
            pct = f"{sign}{delta / base_val * 100:.1f}%" if base_val else "N/A"
            lines.append(
                f"| {label} | {base_val:.4f} {unit} | {cur_val:.4f} {unit} "
                f"| {sign}{delta:.4f} {unit} ({pct}) |"
            )
        lines.append(f"| Iterations | {baseline['iterations']} | {current['iterations']} | |")

    print("\n".join(lines))


if __name__ == "__main__":
    main()
