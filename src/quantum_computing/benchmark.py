import argparse
import json
import resource
import statistics
import time
from typing import Any


def run(iterations: int = 100) -> dict[str, Any]:
    import_start = time.perf_counter()
    from quantum_computing.main import main
    startup_time = time.perf_counter() - import_start

    latencies: list[float] = []
    exit_code = 0
    for _ in range(iterations):
        start = time.perf_counter()
        exit_code = main()
        latencies.append(time.perf_counter() - start)

    total_time = sum(latencies)
    usage = resource.getrusage(resource.RUSAGE_SELF)
    # ru_maxrss is KB on Linux, bytes on macOS
    peak_mb = usage.ru_maxrss / 1024

    sorted_latencies = sorted(latencies)
    p95 = sorted_latencies[int(len(sorted_latencies) * 0.95) - 1]
    p99 = sorted_latencies[int(len(sorted_latencies) * 0.99) - 1]

    return {
        "exit_code": exit_code,
        "iterations": iterations,
        "startup_time_s": startup_time,
        "wall_clock_time_s": total_time,
        "cpu_user_time_s": usage.ru_utime,
        "cpu_sys_time_s": usage.ru_stime,
        "peak_memory_mb": peak_mb,
        "throughput_ops_s": iterations / total_time if total_time > 0 else None,
        "latency_mean_ms": statistics.mean(latencies) * 1000,
        "latency_median_ms": statistics.median(latencies) * 1000,
        "latency_p95_ms": p95 * 1000,
        "latency_p99_ms": p99 * 1000,
    }


def print_metrics(metrics: dict[str, Any]) -> None:
    print(f"exit code:        {metrics['exit_code']}")
    print(f"iterations:       {metrics['iterations']}")
    print(f"startup time:     {metrics['startup_time_s']:.4f} s")
    print(f"wall-clock time:  {metrics['wall_clock_time_s']:.4f} s")
    print(f"cpu time (user):  {metrics['cpu_user_time_s']:.4f} s")
    print(f"cpu time (sys):   {metrics['cpu_sys_time_s']:.4f} s")
    print(f"peak memory:      {metrics['peak_memory_mb']:.2f} MB")
    throughput = metrics["throughput_ops_s"]
    print(f"throughput:       {throughput:.2f} ops/s" if throughput is not None else "throughput:       n/a (elapsed time too small to measure)")
    print(f"latency mean:     {metrics['latency_mean_ms']:.4f} ms")
    print(f"latency median:   {metrics['latency_median_ms']:.4f} ms")
    print(f"latency p95:      {metrics['latency_p95_ms']:.4f} ms")
    print(f"latency p99:      {metrics['latency_p99_ms']:.4f} ms")


def cli() -> int:
    parser = argparse.ArgumentParser(description="Benchmark the quantum-computing app")
    parser.add_argument(
        "-n", "--iterations", type=int, default=100, help="number of runs to average over"
    )
    parser.add_argument(
        "--json", metavar="PATH", help="also write the metrics as JSON to this path"
    )
    args = parser.parse_args()

    metrics = run(args.iterations)
    print_metrics(metrics)

    if args.json:
        with open(args.json, "w") as f:
            json.dump(metrics, f, indent=2)

    return metrics["exit_code"]


if __name__ == "__main__":
    raise SystemExit(cli())
