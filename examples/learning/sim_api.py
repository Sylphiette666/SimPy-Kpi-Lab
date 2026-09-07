from fastapi import FastAPI
from pydantic import BaseModel

from config import SimConfig
from simulation import run_simulation


app = FastAPI(
    title="SimPy Simulation API"
)


class SimulationRequest(BaseModel):

    seed: int = 42

    policy: str = (
        "SHORTEST_QUEUE"
    )

    sim_time: float = 480


@app.post("/simulate")
def simulate(
    request: SimulationRequest
):

    config = SimConfig(
        sim_time=request.sim_time
    )

    result, decisions = (
        run_simulation(
            config=config,
            seed=request.seed,
            policy=request.policy,
        )
    )

    return {
        "result":
            result,

        "decisions":
            decisions
    }