from __future__ import annotations

import csv
import itertools
import json
import math
import statistics
from collections.abc import Iterable
from concurrent.futures import ProcessPoolExecutor
from datetime import UTC, datetime
from pathlib import Path
from typing import Any

from simlab.config import ProjectConfig, SimulationConfig
from simlab.kpi import build_metric_catalog
from simlab.rng import derive_seed
from simlab.simulation import run_replication_payload


def set_dotted_value(data: dict[str, Any], path: str, value: Any) -> None:
    parts = path.split(".")
    if parts[0] == "simulation":
        parts = parts[1:]
    if not parts or any(not part for part in parts):
        raise ValueError(f"invalid parameter path: {path!r}")
    try:
        cursor: Any = data
        for part in parts[:-1]:
            cursor = cursor[int(part)] if isinstance(cursor, list) else cursor[part]
        last = parts[-1]
        if isinstance(cursor, list):
            cursor[int(last)] = value
        else:
            if last not in cursor:
                raise KeyError(last)
            cursor[last] = value
    except (KeyError, IndexError, TypeError, ValueError) as error:
        raise ValueError(f"parameter path does not exist: {path!r}") from error


def expand_scenarios(config: ProjectConfig) -> list[tuple[str, dict[str, Any], SimulationConfig]]:
    grid = config.experiment.parameter_grid
    if not grid:
        return [("base", {}, config.simulation)]

    keys = list(grid)
    scenarios: list[tuple[str, dict[str, Any], SimulationConfig]] = []
    for index, values in enumerate(itertools.product(*(grid[key] for key in keys)), start=1):
        parameters = dict(zip(keys, values, strict=True))
        simulation_data = config.simulation.model_dump(mode="python")
        for key, value in parameters.items():
            set_dotted_value(simulation_data, key, value)
        simulation = SimulationConfig.model_validate(simulation_data)
        label_values = "__".join(
            f"{key.replace('.', '_')}={str(value).replace(' ', '')}"
            for key, value in parameters.items()
        )
        scenarios.append((f"s{index:03d}__{label_values}", parameters, simulation))
    return scenarios


def flatten_numeric(data: dict[str, Any], prefix: str = "") -> dict[str, float]:
    flattened: dict[str, float] = {}
    for key, value in data.items():
        name = f"{prefix}.{key}" if prefix else key
        if isinstance(value, dict):
            flattened.update(flatten_numeric(value, name))
        elif isinstance(value, (int, float)) and not isinstance(value, bool):
            flattened[name] = float(value)
    return flattened


def aggregate(
    records: list[dict[str, Any]],
    confidence_level: float,
    metric_catalog: list[dict[str, str]] | None = None,
) -> list[dict[str, Any]]:
    grouped: dict[tuple[str, str], list[float]] = {}
    parameters: dict[str, dict[str, Any]] = {}
    scenario_counts: dict[str, int] = {}
    catalog_by_metric = {item["metric"]: item for item in metric_catalog or []}
    allowed_metrics = set(catalog_by_metric) if metric_catalog is not None else None
    for record in records:
        scenario = record["scenario"]
        parameters[scenario] = record["parameters"]
        scenario_counts[scenario] = scenario_counts.get(scenario, 0) + 1
        for metric, value in flatten_numeric(record["metrics"]).items():
            if allowed_metrics is not None and metric not in allowed_metrics:
                continue
            grouped.setdefault((scenario, metric), []).append(value)

    z_value = statistics.NormalDist().inv_cdf((1 + confidence_level) / 2)
    pairs = set(grouped)
    if metric_catalog is not None:
        pairs = {
            (scenario, item["metric"]) for scenario in scenario_counts for item in metric_catalog
        }
    rows: list[dict[str, Any]] = []
    for scenario, metric in sorted(pairs):
        values = grouped.get((scenario, metric), [])
        n_valid = len(values)
        n_total = scenario_counts[scenario]
        mean = statistics.fmean(values) if values else None
        std = statistics.stdev(values) if n_valid > 1 else None
        standard_error = std / math.sqrt(n_valid) if std is not None else None
        margin = z_value * standard_error if standard_error is not None else None
        metadata = catalog_by_metric.get(metric, {})
        rows.append(
            {
                "scenario": scenario,
                "parameters": parameters[scenario],
                "metric": metric,
                "role": metadata.get("role"),
                "direction": metadata.get("direction"),
                "unit": metadata.get("unit"),
                "n": n_valid,
                "n_total": n_total,
                "n_missing": n_total - n_valid,
                "mean": mean,
                "std": std,
                "standard_error": standard_error,
                "confidence_level": confidence_level,
                "ci_method": "normal_approximation",
                "ci_low": mean - margin if mean is not None and margin is not None else None,
                "ci_high": mean + margin if mean is not None and margin is not None else None,
                "min": min(values) if values else None,
                "max": max(values) if values else None,
            }
        )
    return rows


def _write_csv(path: Path, rows: Iterable[dict[str, Any]]) -> None:
    rows = list(rows)
    if not rows:
        return
    fieldnames: list[str] = []
    for row in rows:
        for key in row:
            if key not in fieldnames:
                fieldnames.append(key)
    with path.open("w", newline="", encoding="utf-8") as stream:
        writer = csv.DictWriter(stream, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(
                {
                    key: json.dumps(value, ensure_ascii=False, sort_keys=True)
                    if isinstance(value, (dict, list))
                    else value
                    for key, value in row.items()
                }
            )


class ExperimentRunner:
    def __init__(self, config: ProjectConfig):
        self.config = config

    def _tasks(self) -> list[dict[str, Any]]:
        tasks: list[dict[str, Any]] = []
        for name, parameters, simulation in expand_scenarios(self.config):
            for replication in range(self.config.experiment.replications):
                seed_namespace = f"replication:{replication}"
                if not self.config.experiment.common_random_numbers:
                    scenario_fingerprint = json.dumps(
                        parameters,
                        ensure_ascii=False,
                        sort_keys=True,
                        separators=(",", ":"),
                    )
                    seed_namespace += f"|scenario:{scenario_fingerprint}"
                tasks.append(
                    {
                        "simulation": simulation.model_dump(mode="python"),
                        "seed": derive_seed(
                            self.config.experiment.base_seed,
                            seed_namespace,
                        ),
                        "replication": replication,
                        "scenario": name,
                        "parameters": parameters,
                    }
                )
        return tasks

    def run(self, workers: int = 1) -> dict[str, Any]:
        tasks = self._tasks()
        if workers > 1:
            with ProcessPoolExecutor(max_workers=workers) as executor:
                records = list(executor.map(run_replication_payload, tasks))
        else:
            records = [run_replication_payload(task) for task in tasks]

        station_names = sorted(
            {station for record in records for station in record["metrics"].get("station", {})}
        )
        metric_catalog = build_metric_catalog(station_names)
        summary = aggregate(
            records,
            self.config.experiment.confidence_level,
            metric_catalog=metric_catalog,
        )
        return {
            "schema_version": "1.1",
            "generated_at": datetime.now(UTC).isoformat(),
            "project_name": self.config.project_name,
            "config": self.config.model_dump(mode="json"),
            "random_streams": {
                "method": "blake2b_namespaced_v1",
                "common_random_numbers": self.config.experiment.common_random_numbers,
            },
            "metric_catalog": metric_catalog,
            "replications": records,
            "summary": summary,
        }

    def save(self, result: dict[str, Any], output_dir: str | Path | None = None) -> Path:
        root = Path(output_dir or self.config.experiment.output_dir)
        root.mkdir(parents=True, exist_ok=True)

        with (root / "results.json").open("w", encoding="utf-8") as stream:
            json.dump(result, stream, ensure_ascii=False, indent=2)

        replication_rows = []
        for record in result["replications"]:
            replication_rows.append(
                {
                    "scenario": record["scenario"],
                    "parameters": record["parameters"],
                    "replication": record["replication"],
                    "seed": record["seed"],
                    **flatten_numeric(record["metrics"]),
                }
            )
        _write_csv(root / "replications.csv", replication_rows)
        _write_csv(root / "summary.csv", result["summary"])
        return root
