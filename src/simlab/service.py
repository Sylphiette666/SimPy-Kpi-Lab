from __future__ import annotations

import hashlib
import json
import re
from collections.abc import Callable, Mapping
from datetime import UTC, datetime
from pathlib import Path
from threading import RLock
from typing import Any, Literal
from uuid import uuid4

from pydantic import BaseModel, ConfigDict, Field, JsonValue

from simlab.agent import AllowedAction, OpenAIProposalAgent
from simlab.config import ProjectConfig, StationConfig
from simlab.experiment import ExperimentRunner
from simlab.policy import (
    ActionPolicyError,
    PolicyLimits,
    allowed_actions,
    apply_action_to_config,
    canonical_config_hash,
    proposed_action_to_workflow,
    resource_catalog,
    validate_workload,
)
from simlab.workflow import (
    Action,
    ApprovalWorkflow,
    EnableResourceAction,
    IdempotencyConflict,
    Proposal,
    ProposalNotFound,
    ProposalStatus,
    SetParameterAction,
    VersionConflict,
    WorkflowSession,
    validate_action,
)

_OPERATION_ID_PATTERN = re.compile(r"^[A-Za-z0-9][A-Za-z0-9_.:@/-]{0,89}$")


class ControlServiceError(RuntimeError):
    """Base class for control-plane orchestration errors."""


class SessionNotFound(ControlServiceError):
    pass


class RunNotFound(ControlServiceError):
    pass


class StaleProposal(ControlServiceError):
    pass


class RunExecutionError(ControlServiceError):
    pass


class OperationInProgress(ControlServiceError):
    pass


class _ViewModel(BaseModel):
    model_config = ConfigDict(extra="forbid", frozen=True)


class ServiceEvent(_ViewModel):
    sequence: int = Field(ge=1)
    timestamp: datetime
    event: str
    session_id: str
    payload: dict[str, JsonValue] = Field(default_factory=dict)


class RunRecord(_ViewModel):
    run_id: str
    operation_id: str
    session_id: str
    status: Literal["running", "succeeded", "failed"]
    config_hash: str
    workflow_version: int = Field(ge=0)
    created_at: datetime
    completed_at: datetime | None = None
    output_dir: str | None = None
    error: str | None = None
    result: dict[str, Any] | None = None


class AIPlanRecord(_ViewModel):
    plan_id: str
    operation_id: str
    session_id: str
    base_config_hash: str
    objective: str | None
    summary: str
    caveats: tuple[str, ...]
    proposals: tuple[Proposal, ...]


class ControlSessionView(_ViewModel):
    session_id: str
    workflow_version: int = Field(ge=0)
    config_hash: str
    config: ProjectConfig
    workflow: WorkflowSession
    run_ids: tuple[str, ...]
    plan_ids: tuple[str, ...]


class _IdempotentResult:
    __slots__ = ("fingerprint", "result")

    def __init__(self, fingerprint: str, result: Any):
        self.fingerprint = fingerprint
        self.result = result


class _SessionRuntime:
    def __init__(self, config: ProjectConfig, session_id: str):
        self.lock = RLock()
        self.config = config.model_copy(deep=True)
        self.resources: dict[str, StationConfig] = resource_catalog(config)
        self.workflow = ApprovalWorkflow(
            config.experiment,
            session_id=session_id,
            initial_resources=self.resources,
        )
        self.proposal_hashes: dict[str, str] = {}
        self.plans: dict[str, AIPlanRecord] = {}
        self.runs: dict[str, RunRecord] = {}
        self.events: list[ServiceEvent] = []
        self.operations: dict[str, _IdempotentResult] = {}


class SimulationControlService:
    """Safe orchestration for propose -> human approve/reject -> next immutable run.

    The service never mutates an in-flight SimPy environment. Every run receives
    a deep-copied ProjectConfig snapshot and writes to a server-assigned directory.
    """

    def __init__(
        self,
        *,
        output_root: str | Path = "outputs/api_sessions",
        limits: PolicyLimits | None = None,
    ) -> None:
        self.output_root = Path(output_root)
        self.limits = limits or PolicyLimits()
        self._sessions: dict[str, _SessionRuntime] = {}
        self._lock = RLock()

    def create_session(
        self,
        config: ProjectConfig | Mapping[str, Any],
        *,
        session_id: str | None = None,
    ) -> ControlSessionView:
        config_value = (
            config.model_copy(deep=True)
            if isinstance(config, ProjectConfig)
            else ProjectConfig.model_validate(dict(config))
        )
        validate_workload(config_value, self.limits)
        session_key = session_id or f"session-{uuid4().hex}"
        if re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9_-]{0,63}", session_key) is None:
            raise ValueError(
                "session_id must be 1-64 ASCII letters, digits, underscores or hyphens"
            )
        runtime = _SessionRuntime(config_value, session_key)
        with self._lock:
            if session_key in self._sessions:
                raise ControlServiceError(f"session already exists: {session_key}")
            self._sessions[session_key] = runtime
        with runtime.lock:
            self._append_event(runtime, "session_created", {"config_hash": self._hash(runtime)})
            return self._view(runtime)

    def get_session(self, session_id: str) -> ControlSessionView:
        runtime = self._runtime(session_id)
        with runtime.lock:
            return self._view(runtime)

    def list_events(self, session_id: str, after_sequence: int = 0) -> tuple[ServiceEvent, ...]:
        runtime = self._runtime(session_id)
        with runtime.lock:
            return tuple(
                event.model_copy(deep=True)
                for event in runtime.events
                if event.sequence > after_sequence
            )

    def get_allowed_actions(self, session_id: str) -> list[AllowedAction]:
        runtime = self._runtime(session_id)
        with runtime.lock:
            catalog = allowed_actions(runtime.config)
            known_pairs = {(item.action_type, item.target) for item in catalog}
            for resource_name in runtime.resources:
                if ("enable_resource", resource_name) not in known_pairs:
                    catalog.append(
                        AllowedAction(
                            action_type="enable_resource",
                            target=resource_name,
                            description="重新启用会话创建时登记的工位。",
                        )
                    )
            return catalog

    def submit_action(
        self,
        session_id: str,
        action: Action | Mapping[str, Any],
        *,
        operation_id: str,
        expected_version: int,
        actor: str = "operator",
    ) -> Proposal:
        self._validate_operation_id(operation_id)
        runtime = self._runtime(session_id)
        action_value = validate_action(action)
        with runtime.lock:
            fingerprint = self._fingerprint(
                "submit_action",
                {
                    "action": action_value.model_dump(mode="json"),
                    "actor": actor,
                    "expected_version": expected_version,
                },
            )
            replay = self._replay(runtime, operation_id, fingerprint)
            if replay is not None:
                return replay
            self._require_version(runtime, expected_version)
            apply_action_to_config(
                runtime.config,
                action_value,
                resources=runtime.resources,
                limits=self.limits,
            )
            proposal = runtime.workflow.propose(
                action_value,
                operation_id=operation_id,
                expected_version=expected_version,
                actor=actor,
            )
            runtime.proposal_hashes[proposal.proposal_id] = self._hash(runtime)
            self._remember(runtime, operation_id, fingerprint, proposal)
            self._append_event(
                runtime,
                "proposal_created",
                {
                    "proposal_id": proposal.proposal_id,
                    "status": proposal.status.value,
                    "workflow_version": runtime.workflow.version,
                },
            )
            return proposal

    def submit_ai_plan(
        self,
        session_id: str,
        agent: OpenAIProposalAgent,
        *,
        operation_id: str,
        expected_version: int,
        objective: str | None = None,
        kpi_summary: Any = None,
        actor: str = "openai-agent",
    ) -> AIPlanRecord:
        self._validate_operation_id(operation_id)
        runtime = self._runtime(session_id)
        with runtime.lock:
            base_config = runtime.config.model_copy(deep=True)
            base_hash = self._hash(runtime)
            fingerprint = self._fingerprint(
                "ai_plan",
                {
                    "session_id": session_id,
                    "objective": objective,
                    "kpi_summary": kpi_summary,
                    "actor": actor,
                    "expected_version": expected_version,
                },
            )
            replay = self._replay(runtime, operation_id, fingerprint)
            if replay is not None:
                return replay
            self._require_version(runtime, expected_version)
            action_catalog = self.get_allowed_actions(session_id)
            resource_snapshot = {
                name: station.model_copy(deep=True) for name, station in runtime.resources.items()
            }
            self._reserve(runtime, operation_id, fingerprint)

        try:
            generated = agent.propose(
                base_config.model_dump(mode="json"),
                kpi_summary,
                action_catalog,
                objective=objective,
            )
            workflow_actions = [proposed_action_to_workflow(action) for action in generated.actions]
            preview = base_config
            step_base_hashes: list[str] = []
            topology_action_seen = False
            for action in workflow_actions:
                if (
                    topology_action_seen
                    and isinstance(action, SetParameterAction)
                    and re.fullmatch(r"simulation\.stations\.\d+\..+", action.path)
                ):
                    raise ActionPolicyError(
                        "an AI plan cannot use an indexed station parameter after an "
                        "enable_resource action; move resource actions to the end"
                    )
                step_base_hashes.append(canonical_config_hash(preview))
                preview = apply_action_to_config(
                    preview,
                    action,
                    resources=resource_snapshot,
                    limits=self.limits,
                )
                for station in preview.simulation.stations:
                    resource_snapshot[station.name] = station.model_copy(deep=True)
                if isinstance(action, EnableResourceAction):
                    topology_action_seen = True

            with runtime.lock:
                self._require_version(runtime, expected_version)
                if self._hash(runtime) != base_hash:
                    raise StaleProposal(
                        "session config changed while the AI proposal was generated"
                    )
                proposals: list[Proposal] = []
                for index, (action, action_base_hash) in enumerate(
                    zip(workflow_actions, step_base_hashes, strict=True),
                    start=1,
                ):
                    proposal = runtime.workflow.propose(
                        action,
                        operation_id=self._internal_operation_id(
                            operation_id,
                            f"proposal:{index}",
                        ),
                        expected_version=runtime.workflow.version,
                        actor=actor,
                    )
                    runtime.proposal_hashes[proposal.proposal_id] = action_base_hash
                    proposals.append(proposal)
                plan = AIPlanRecord(
                    plan_id=f"plan-{uuid4().hex}",
                    operation_id=operation_id,
                    session_id=session_id,
                    base_config_hash=base_hash,
                    objective=objective,
                    summary=generated.summary,
                    caveats=tuple(generated.caveats),
                    proposals=tuple(proposals),
                )
                runtime.plans[plan.plan_id] = plan
                self._remember(runtime, operation_id, fingerprint, plan)
                self._append_event(
                    runtime,
                    "ai_plan_created",
                    {
                        "plan_id": plan.plan_id,
                        "proposal_count": len(proposals),
                        "workflow_version": runtime.workflow.version,
                    },
                )
                return plan
        except Exception:
            with runtime.lock:
                self._clear_pending(runtime, operation_id, fingerprint)
            raise

    def approve_and_apply(
        self,
        session_id: str,
        proposal_id: str,
        *,
        operation_id: str,
        expected_version: int,
        actor: str,
        reason: str | None = None,
    ) -> Proposal:
        self._validate_operation_id(operation_id)
        runtime = self._runtime(session_id)
        with runtime.lock:
            fingerprint = self._fingerprint(
                "approve_and_apply",
                {
                    "proposal_id": proposal_id,
                    "actor": actor,
                    "reason": reason,
                    "expected_version": expected_version,
                },
            )
            replay = self._replay(runtime, operation_id, fingerprint)
            if replay is not None:
                return replay
            self._require_version(runtime, expected_version)
            proposal = runtime.workflow.get_proposal(proposal_id)
            self._require_fresh_proposal(runtime, proposal)
            preview = apply_action_to_config(
                runtime.config,
                proposal.action,
                resources=runtime.resources,
                limits=self.limits,
            )
            runtime.workflow.approve(
                proposal_id,
                operation_id=self._internal_operation_id(operation_id, "approve"),
                expected_version=expected_version,
                actor=actor,
                reason=reason,
            )
            applied = runtime.workflow.apply(
                proposal_id,
                operation_id=self._internal_operation_id(operation_id, "apply"),
                expected_version=runtime.workflow.version,
                actor="control-service",
            )
            if applied.status is not ProposalStatus.APPLIED:
                raise ControlServiceError(applied.error or "approved proposal could not be applied")
            runtime.config = preview
            for station in preview.simulation.stations:
                runtime.resources[station.name] = station.model_copy(deep=True)
            self._remember(runtime, operation_id, fingerprint, applied)
            self._append_event(
                runtime,
                "proposal_applied",
                {
                    "proposal_id": proposal_id,
                    "status": applied.status.value,
                    "config_hash": self._hash(runtime),
                    "workflow_version": runtime.workflow.version,
                },
            )
            return applied

    def reject(
        self,
        session_id: str,
        proposal_id: str,
        *,
        operation_id: str,
        expected_version: int,
        actor: str,
        reason: str | None = None,
    ) -> Proposal:
        self._validate_operation_id(operation_id)
        runtime = self._runtime(session_id)
        with runtime.lock:
            fingerprint = self._fingerprint(
                "reject",
                {
                    "proposal_id": proposal_id,
                    "actor": actor,
                    "reason": reason,
                    "expected_version": expected_version,
                },
            )
            replay = self._replay(runtime, operation_id, fingerprint)
            if replay is not None:
                return replay
            self._require_version(runtime, expected_version)
            rejected = runtime.workflow.reject(
                proposal_id,
                operation_id=self._internal_operation_id(operation_id, "reject"),
                expected_version=expected_version,
                actor=actor,
                reason=reason,
            )
            self._remember(runtime, operation_id, fingerprint, rejected)
            self._append_event(
                runtime,
                "proposal_rejected",
                {
                    "proposal_id": proposal_id,
                    "status": rejected.status.value,
                    "workflow_version": runtime.workflow.version,
                },
            )
            return rejected

    def run(
        self,
        session_id: str,
        *,
        operation_id: str,
        expected_version: int,
        workers: int = 1,
    ) -> RunRecord:
        self._validate_operation_id(operation_id)
        if not 1 <= workers <= 16:
            raise ValueError("workers must be between 1 and 16")
        runtime = self._runtime(session_id)
        with runtime.lock:
            config_snapshot = runtime.config.model_copy(deep=True)
            config_hash = canonical_config_hash(config_snapshot)
            fingerprint = self._fingerprint(
                "run",
                {
                    "session_id": session_id,
                    "workers": workers,
                    "expected_version": expected_version,
                },
            )
            replay = self._replay(runtime, operation_id, fingerprint)
            if replay is not None:
                if replay.status == "failed":
                    raise RunExecutionError(
                        f"simulation run failed: {replay.error or 'unknown error'}"
                    )
                return replay
            self._require_version(runtime, expected_version)
            run_id = f"run-{uuid4().hex}"
            record = RunRecord(
                run_id=run_id,
                operation_id=operation_id,
                session_id=session_id,
                status="running",
                config_hash=config_hash,
                workflow_version=runtime.workflow.version,
                created_at=datetime.now(UTC),
            )
            runtime.runs[run_id] = record
            self._remember(runtime, operation_id, fingerprint, record)
            self._append_event(
                runtime,
                "run_started",
                {"run_id": run_id, "config_hash": config_hash},
            )

        try:
            runner = ExperimentRunner(config_snapshot)
            result = runner.run(workers=workers)
            output_dir = self.output_root / session_id / run_id
            runner.save(result, output_dir)
        except Exception as error:
            message = str(error)[:2000]
            failed = record.model_copy(
                update={
                    "status": "failed",
                    "completed_at": datetime.now(UTC),
                    "error": message,
                },
                deep=True,
            )
            with runtime.lock:
                runtime.runs[run_id] = failed
                self._remember(runtime, operation_id, fingerprint, failed)
                self._append_event(
                    runtime,
                    "run_failed",
                    {"run_id": run_id, "error": message},
                )
            raise RunExecutionError(f"simulation run failed: {message}") from error

        completed = record.model_copy(
            update={
                "status": "succeeded",
                "completed_at": datetime.now(UTC),
                "output_dir": str(output_dir.resolve()),
                "result": result,
            },
            deep=True,
        )
        with runtime.lock:
            runtime.runs[run_id] = completed
            self._remember(runtime, operation_id, fingerprint, completed)
            self._append_event(
                runtime,
                "run_succeeded",
                {
                    "run_id": run_id,
                    "output_dir": completed.output_dir,
                    "config_hash": config_hash,
                },
            )
            return completed

    def get_run(self, session_id: str, run_id: str) -> RunRecord:
        runtime = self._runtime(session_id)
        with runtime.lock:
            record = runtime.runs.get(run_id)
            if record is None:
                raise RunNotFound(f"run not found: {run_id}")
            return record.model_copy(deep=True)

    def get_workflow_audit(self, session_id: str):
        return self._runtime(session_id).workflow.audit_log()

    def _runtime(self, session_id: str) -> _SessionRuntime:
        with self._lock:
            runtime = self._sessions.get(session_id)
        if runtime is None:
            raise SessionNotFound(f"session not found: {session_id}")
        return runtime

    @staticmethod
    def _validate_operation_id(operation_id: str) -> None:
        if _OPERATION_ID_PATTERN.fullmatch(operation_id) is None or operation_id.startswith(
            "simlab:"
        ):
            raise ValueError(
                "operation_id must be 1-90 supported ASCII characters and start "
                "with a letter or digit; the simlab: prefix is reserved"
            )

    @staticmethod
    def _internal_operation_id(operation_id: str, phase: str) -> str:
        digest = hashlib.sha256(operation_id.encode("utf-8")).hexdigest()[:32]
        return f"simlab:{digest}:{phase}"

    @staticmethod
    def _hash(runtime: _SessionRuntime) -> str:
        return canonical_config_hash(runtime.config)

    @staticmethod
    def _require_version(runtime: _SessionRuntime, expected_version: int) -> None:
        current = runtime.workflow.version
        if expected_version != current:
            raise VersionConflict(expected_version, current)

    def _require_fresh_proposal(self, runtime: _SessionRuntime, proposal: Proposal) -> None:
        base_hash = runtime.proposal_hashes.get(proposal.proposal_id)
        if base_hash is None:
            raise ProposalNotFound(f"proposal not registered: {proposal.proposal_id}")
        if base_hash != self._hash(runtime):
            raise StaleProposal("proposal was generated for an older config; create a new proposal")

    @staticmethod
    def _fingerprint(operation: str, payload: Mapping[str, Any]) -> str:
        return json.dumps(
            {"operation": operation, "payload": payload},
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
            allow_nan=False,
        )

    @staticmethod
    def _replay(runtime: _SessionRuntime, operation_id: str, fingerprint: str):
        record = runtime.operations.get(operation_id)
        if record is None:
            return None
        if record.fingerprint != fingerprint:
            raise IdempotencyConflict(
                f"operation_id {operation_id!r} was already used for a different request"
            )
        result = record.result
        if result is None:
            raise OperationInProgress(f"operation is still in progress: {operation_id}")
        return result.model_copy(deep=True) if isinstance(result, BaseModel) else result

    @staticmethod
    def _reserve(runtime: _SessionRuntime, operation_id: str, fingerprint: str) -> None:
        runtime.operations[operation_id] = _IdempotentResult(fingerprint, None)

    @staticmethod
    def _clear_pending(
        runtime: _SessionRuntime,
        operation_id: str,
        fingerprint: str,
    ) -> None:
        record = runtime.operations.get(operation_id)
        if record is not None and record.fingerprint == fingerprint and record.result is None:
            runtime.operations.pop(operation_id, None)

    @staticmethod
    def _remember(
        runtime: _SessionRuntime,
        operation_id: str,
        fingerprint: str,
        result: Any,
    ) -> None:
        stored = result.model_copy(deep=True) if isinstance(result, BaseModel) else result
        runtime.operations[operation_id] = _IdempotentResult(fingerprint, stored)

    @staticmethod
    def _append_event(
        runtime: _SessionRuntime,
        event: str,
        payload: dict[str, JsonValue],
    ) -> None:
        runtime.events.append(
            ServiceEvent(
                sequence=len(runtime.events) + 1,
                timestamp=datetime.now(UTC),
                event=event,
                session_id=runtime.workflow.snapshot().session_id,
                payload=payload,
            )
        )

    @staticmethod
    def _view(runtime: _SessionRuntime) -> ControlSessionView:
        workflow = runtime.workflow.snapshot()
        return ControlSessionView(
            session_id=workflow.session_id,
            workflow_version=workflow.version,
            config_hash=canonical_config_hash(runtime.config),
            config=runtime.config.model_copy(deep=True),
            workflow=workflow,
            run_ids=tuple(runtime.runs),
            plan_ids=tuple(runtime.plans),
        )


AgentFactory = Callable[[], OpenAIProposalAgent]


__all__ = [
    "AIPlanRecord",
    "AgentFactory",
    "ControlServiceError",
    "OperationInProgress",
    "ControlSessionView",
    "RunExecutionError",
    "RunNotFound",
    "RunRecord",
    "ServiceEvent",
    "SessionNotFound",
    "SimulationControlService",
    "StaleProposal",
]
