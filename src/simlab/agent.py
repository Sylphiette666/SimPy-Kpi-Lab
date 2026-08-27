from __future__ import annotations

import json
import os
from collections.abc import Mapping, Sequence
from typing import Any, Literal, Protocol, TypeAlias

from openai import (
    APIConnectionError,
    APIStatusError,
    APITimeoutError,
    AuthenticationError,
    OpenAI,
    OpenAIError,
    RateLimitError,
)
from pydantic import BaseModel, ConfigDict, Field, ValidationError

ActionType: TypeAlias = Literal["set_parameter", "change_policy", "enable_resource"]
RiskLevel: TypeAlias = Literal["low", "medium", "high"]
ProposedValue: TypeAlias = str | int | float | bool | None


class AllowedAction(BaseModel):
    """One action/target pair that the proposal layer may recommend."""

    model_config = ConfigDict(extra="forbid")

    action_type: ActionType
    target: str
    description: str = ""


class ProposedAction(BaseModel):
    """A recommendation only; this model intentionally exposes no execution hook."""

    model_config = ConfigDict(extra="forbid")

    action_type: ActionType
    target: str
    proposed_value: ProposedValue
    rationale: str
    expected_effect: str
    risk: RiskLevel
    requires_approval: bool = True


class ActionProposal(BaseModel):
    """Structured, human-reviewable actions proposed from simulation evidence."""

    model_config = ConfigDict(extra="forbid")

    summary: str
    actions: list[ProposedAction] = Field(default_factory=list, max_length=10)
    caveats: list[str] = Field(default_factory=list)


class ProposalGenerationError(RuntimeError):
    """A safe, user-facing failure raised by the optional OpenAI proposal adapter."""


class _ResponsesParser(Protocol):
    def parse(self, **kwargs: Any) -> Any: ...


class _ProposalClient(Protocol):
    responses: _ResponsesParser


class OpenAIProposalAgent:
    """Generate allowlisted action proposals without applying any changes."""

    def __init__(
        self,
        model: str = "gpt-5.6",
        max_output_tokens: int = 1500,
        timeout_seconds: float = 60.0,
        max_retries: int = 2,
        store: bool = False,
        client: _ProposalClient | None = None,
    ) -> None:
        if client is None:
            api_key = os.getenv("OPENAI_API_KEY")
            if not api_key:
                raise ProposalGenerationError("OPENAI_API_KEY is not set")
            client = OpenAI(
                api_key=api_key,
                timeout=timeout_seconds,
                max_retries=max_retries,
            )
        self.client = client
        self.model = model
        self.max_output_tokens = max_output_tokens
        self.store = store

    def propose(
        self,
        current_config: Mapping[str, Any],
        kpi_summary: Any,
        allowed_actions: Sequence[AllowedAction | Mapping[str, Any]],
        objective: str | None = None,
    ) -> ActionProposal:
        """Return proposals only; callers must separately review and apply them."""

        normalized_actions = self._normalize_allowed_actions(allowed_actions)
        payload = {
            "current_config": dict(current_config),
            "kpi_summary": kpi_summary,
            "objective": objective,
            "allowed_actions": [action.model_dump(mode="json") for action in normalized_actions],
        }
        try:
            serialized_payload = json.dumps(payload, ensure_ascii=False)
        except (TypeError, ValueError) as error:
            raise ProposalGenerationError(
                "提案输入必须是可序列化的配置、KPI 摘要和允许动作。"
            ) from error

        try:
            response = self.client.responses.parse(
                model=self.model,
                input=[
                    {
                        "role": "system",
                        "content": (
                            "你是离散事件仿真的安全提案助手。你只能生成供人工审核的结构化"
                            "动作建议，绝不能执行、应用或模拟执行任何动作，也不得输出代码、"
                            "命令、脚本或工具调用。每项建议必须与 allowed_actions 中的"
                            " action_type 和 target 完全匹配；动作类型只能是 set_parameter、"
                            "change_policy 或 enable_resource。所有动作的 requires_approval"
                            " 必须为 true。多动作计划中，enable_resource 必须放在所有针对"
                            " simulation.stations.<index> 的 set_parameter 动作之后，以免工位"
                            " 下标漂移。只根据 current_config、kpi_summary 和 objective"
                            " 推断；证据不足时返回空 actions 并在 caveats 中说明。用中文回答。"
                        ),
                    },
                    {"role": "user", "content": serialized_payload},
                ],
                text_format=ActionProposal,
                max_output_tokens=self.max_output_tokens,
                store=self.store,
            )
        except AuthenticationError as error:
            raise ProposalGenerationError("OpenAI API 认证失败，请检查 OPENAI_API_KEY。") from error
        except RateLimitError as error:
            raise ProposalGenerationError("OpenAI API 当前限流或额度不足，请稍后重试。") from error
        except APITimeoutError as error:
            raise ProposalGenerationError("OpenAI API 请求超时。") from error
        except APIConnectionError as error:
            raise ProposalGenerationError("无法连接 OpenAI API，请检查网络。") from error
        except APIStatusError as error:
            request_id = getattr(error, "request_id", None)
            suffix = f"（request_id={request_id}）" if request_id else ""
            raise ProposalGenerationError(
                f"OpenAI API 返回 HTTP {error.status_code}{suffix}。"
            ) from error
        except OpenAIError as error:
            raise ProposalGenerationError("OpenAI API 请求失败。") from error

        status = getattr(response, "status", None)
        if status not in {None, "completed"}:
            details = getattr(response, "incomplete_details", None)
            reason = getattr(details, "reason", None)
            response_error = getattr(response, "error", None)
            error_code = getattr(response_error, "code", None)
            detail = reason or error_code
            suffix = f"：{detail}" if detail else ""
            raise ProposalGenerationError(f"OpenAI 响应状态为 {status}{suffix}")

        parsed = getattr(response, "output_parsed", None)
        if parsed is None:
            raise ProposalGenerationError(
                "OpenAI 响应未包含可解析的动作提案，可能是拒答或输出不完整。"
            )
        try:
            proposal = ActionProposal.model_validate(parsed)
        except ValidationError as error:
            raise ProposalGenerationError("OpenAI 动作提案未通过本地结构校验。") from error

        allowed_pairs = {(action.action_type, action.target) for action in normalized_actions}
        for action in proposal.actions:
            if not action.requires_approval:
                raise ProposalGenerationError("动作提案必须要求人工审批。")
            if (action.action_type, action.target) not in allowed_pairs:
                raise ProposalGenerationError(
                    f"OpenAI 动作提案超出允许范围：{action.action_type}:{action.target}"
                )
        return proposal

    @staticmethod
    def _normalize_allowed_actions(
        allowed_actions: Sequence[AllowedAction | Mapping[str, Any]],
    ) -> list[AllowedAction]:
        if isinstance(allowed_actions, (str, bytes)) or not allowed_actions:
            raise ProposalGenerationError("allowed_actions must contain at least one action")

        normalized: list[AllowedAction] = []
        try:
            for action in allowed_actions:
                normalized.append(
                    action
                    if isinstance(action, AllowedAction)
                    else AllowedAction.model_validate(action)
                )
        except (TypeError, ValidationError) as error:
            raise ProposalGenerationError("allowed_actions contains an invalid action") from error
        return normalized


__all__ = [
    "ActionProposal",
    "ActionType",
    "AllowedAction",
    "OpenAIProposalAgent",
    "ProposalGenerationError",
    "ProposedAction",
    "ProposedValue",
    "RiskLevel",
]
