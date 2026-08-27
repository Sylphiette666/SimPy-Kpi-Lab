from __future__ import annotations

import random
from typing import Any

import simpy

from simlab.config import SimulationConfig
from simlab.kpi import KPICollector
from simlab.rng import derive_seed


def run_replication(
    simulation: SimulationConfig,
    seed: int,
    replication: int = 0,
    scenario: str = "base",
    parameters: dict[str, Any] | None = None,
) -> dict[str, Any]:
    """Run one deterministic replication for a given seed."""

    arrival_rng = random.Random(derive_seed(seed, "arrivals"))
    station_rngs = {
        station.name: random.Random(derive_seed(seed, f"service:{station.name}"))
        for station in simulation.stations
    }
    env = simpy.Environment()
    resources = {
        station.name: simpy.Resource(env, capacity=station.capacity)
        for station in simulation.stations
    }
    collector = KPICollector(
        warmup=simulation.warmup,
        until=simulation.until,
        station_capacities={station.name: station.capacity for station in simulation.stations},
        cycle_time_target=simulation.cycle_time_target,
    )

    def customer(customer_id: int):
        collector.arrival(customer_id, env.now)
        for station in simulation.stations:
            collector.queue_enter(customer_id, station.name, env.now)
            with resources[station.name].request() as request:
                yield request
                collector.service_start(customer_id, station.name, env.now)
                duration = station.service_time.sample(station_rngs[station.name])
                collector.service_interval(station.name, env.now, env.now + duration)
                yield env.timeout(duration)
        collector.completion(customer_id, env.now)

    def arrivals():
        customer_id = 0
        if not simulation.first_arrival_at_zero:
            delay = simulation.arrival_interarrival.sample(arrival_rng)
            if delay >= simulation.until:
                return
            yield env.timeout(delay)

        while env.now < simulation.until:
            if simulation.max_arrivals is not None and customer_id >= simulation.max_arrivals:
                return
            env.process(customer(customer_id))
            customer_id += 1
            delay = simulation.arrival_interarrival.sample(arrival_rng)
            if env.now + delay >= simulation.until:
                return
            yield env.timeout(delay)

    env.process(arrivals())
    env.run(until=simulation.until)

    return {
        "scenario": scenario,
        "parameters": parameters or {},
        "replication": replication,
        "seed": seed,
        "metrics": collector.finalize(),
    }


def run_replication_payload(payload: dict[str, Any]) -> dict[str, Any]:
    """Pickle-friendly adapter used by ProcessPoolExecutor."""

    simulation = SimulationConfig.model_validate(payload["simulation"])
    return run_replication(
        simulation=simulation,
        seed=payload["seed"],
        replication=payload["replication"],
        scenario=payload["scenario"],
        parameters=payload["parameters"],
    )
