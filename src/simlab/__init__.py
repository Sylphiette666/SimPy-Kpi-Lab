"""SimPy KPI Lab public package."""

from importlib.metadata import PackageNotFoundError, version

from simlab.config import ProjectConfig
from simlab.experiment import ExperimentRunner
from simlab.service import SimulationControlService
from simlab.simulation import run_replication
from simlab.workflow import ApprovalWorkflow

__all__ = [
    "ApprovalWorkflow",
    "ExperimentRunner",
    "ProjectConfig",
    "SimulationControlService",
    "run_replication",
]
try:
    __version__ = version("simpy-kpi-lab")
except PackageNotFoundError:  # pragma: no cover - only when run outside an installation.
    __version__ = "0+unknown"
