from fastapi import FastAPI
from pydantic import BaseModel


app = FastAPI()


class SimulationRequest(BaseModel):

    machines: int

    sim_time: float


@app.post("/simulate")
def simulate(
    request: SimulationRequest
):

    return {
        "machines":
            request.machines,

        "sim_time":
            request.sim_time
    }