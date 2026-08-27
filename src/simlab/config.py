from __future__ import annotations

import os
import random
from pathlib import Path
from typing import Any, Literal

import yaml
from pydantic import BaseModel, ConfigDict, Field, model_validator


class DistributionConfig(BaseModel):
    """Supported duration/interarrival distributions, expressed in model time units."""

    model_config = ConfigDict(extra="forbid")

    kind: Literal["exponential", "deterministic", "uniform", "triangular"]
    mean: float | None = None
    value: float | None = None
    low: float | None = None
    high: float | None = None
    mode: float | None = None

    @model_validator(mode="after")
    def validate_parameters(self) -> DistributionConfig:
        if self.kind == "exponential" and (self.mean is None or self.mean <= 0):
            raise ValueError("exponential distribution requires mean > 0")
        if self.kind == "deterministic" and (self.value is None or self.value < 0):
            raise ValueError("deterministic distribution requires value >= 0")
        if self.kind == "uniform":
            if self.low is None or self.high is None or self.low > self.high:
                raise ValueError("uniform distribution requires low <= high")
            if self.low < 0:
                raise ValueError("uniform distribution requires low >= 0")
        if self.kind == "triangular":
            if self.low is None or self.high is None or self.mode is None:
                raise ValueError("triangular distribution requires low, mode and high")
            if not self.low <= self.mode <= self.high:
                raise ValueError("triangular distribution requires low <= mode <= high")
            if self.low < 0:
                raise ValueError("triangular distribution requires low >= 0")
        return self

    def sample(self, rng: random.Random) -> float:
        if self.kind == "exponential":
            return rng.expovariate(1.0 / float(self.mean))
        if self.kind == "deterministic":
            return float(self.value)
        if self.kind == "uniform":
            return rng.uniform(float(self.low), float(self.high))
        return rng.triangular(float(self.low), float(self.high), float(self.mode))


class StationConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    name: str = Field(min_length=1)
    capacity: int = Field(default=1, ge=1)
    service_time: DistributionConfig


class SimulationConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    name: str = "service_system"
    until: float = Field(gt=0)
    warmup: float = Field(default=0, ge=0)
    first_arrival_at_zero: bool = True
    max_arrivals: int | None = Field(default=None, ge=1)
    arrival_interarrival: DistributionConfig
    stations: list[StationConfig] = Field(min_length=1)
    cycle_time_target: float | None = Field(default=None, gt=0)

    @model_validator(mode="after")
    def validate_simulation(self) -> SimulationConfig:
        if self.warmup >= self.until:
            raise ValueError("warmup must be smaller than until")
        names = [station.name for station in self.stations]
        if len(names) != len(set(names)):
            raise ValueError("station names must be unique")
        arrival = self.arrival_interarrival
        if arrival.kind == "deterministic" and arrival.value == 0:
            raise ValueError("deterministic interarrival value must be > 0")
        if arrival.kind in {"uniform", "triangular"} and arrival.low == 0:
            raise ValueError("interarrival distribution must have strictly positive support")
        return self


class ExperimentConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    replications: int = Field(default=10, ge=1)
    base_seed: int = Field(default=20260825, ge=0)
    common_random_numbers: bool = True
    confidence_level: float = Field(default=0.95, gt=0, lt=1)
    parameter_grid: dict[str, list[Any]] = Field(default_factory=dict)
    output_dir: str = "outputs"

    @model_validator(mode="after")
    def validate_grid(self) -> ExperimentConfig:
        empty = [key for key, values in self.parameter_grid.items() if not values]
        if empty:
            raise ValueError(f"parameter_grid entries cannot be empty: {empty}")
        return self


class OpenAIConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    model: str = "gpt-5.6"
    max_output_tokens: int = Field(default=2500, ge=256)
    timeout_seconds: float = Field(default=60.0, gt=0)
    max_retries: int = Field(default=2, ge=0, le=10)
    store: bool = False


class ProjectConfig(BaseModel):
    model_config = ConfigDict(extra="forbid")

    project_name: str = "simpy-kpi-lab"
    simulation: SimulationConfig
    experiment: ExperimentConfig = Field(default_factory=ExperimentConfig)
    openai: OpenAIConfig = Field(default_factory=OpenAIConfig)

    @classmethod
    def load(cls, path: str | Path) -> ProjectConfig:
        config_path = Path(path)
        with config_path.open(encoding="utf-8") as stream:
            data = yaml.safe_load(stream)
        if not isinstance(data, dict):
            raise ValueError(f"configuration root must be a mapping: {config_path}")
        config = cls.model_validate(data)
        env_model = os.getenv("SIMLAB_OPENAI_MODEL")
        if env_model:
            config.openai.model = env_model
        return config
