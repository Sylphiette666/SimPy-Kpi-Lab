from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any, Literal

from openai import (
    APIConnectionError,
    APIStatusError,
    APITimeoutError,
    AuthenticationError,
    OpenAI,
    OpenAIError,
    RateLimitError,
)
from pydantic import BaseModel, Field


class Finding(BaseModel):
    title: str
    evidence: str
    impact: str


class Recommendation(BaseModel):
    priority: Literal["high", "medium", "low"]
    action: str
    rationale: str
    expected_effect: str


class KPIAnalysis(BaseModel):
    executive_summary: str
    best_scenario: str | None
    findings: list[Finding] = Field(default_factory=list)
    recommendations: list[Recommendation] = Field(default_factory=list)
    caveats: list[str] = Field(default_factory=list)

    def to_markdown(self) -> str:
        lines = ["# AI KPI 分析", "", self.executive_summary, ""]
        if self.best_scenario:
            lines.extend([f"**建议场景：** `{self.best_scenario}`", ""])
        lines.extend(["## 关键发现", ""])
        for finding in self.findings:
            lines.extend(
                [f"### {finding.title}", "", finding.evidence, "", f"影响：{finding.impact}", ""]
            )
        lines.extend(["## 建议", ""])
        for item in self.recommendations:
            lines.extend(
                [
                    f"- **[{item.priority}] {item.action}** — "
                    f"{item.rationale}；预期：{item.expected_effect}",
                    "",
                ]
            )
        lines.extend(["## 注意事项", ""])
        lines.extend(f"- {caveat}" for caveat in self.caveats)
        lines.append("")
        return "\n".join(lines)


class AIAnalysisError(RuntimeError):
    """A safe, user-facing failure raised by the optional OpenAI adapter."""


class OpenAIKPIAnalyst:
    def __init__(
        self,
        model: str = "gpt-5.6",
        max_output_tokens: int = 2500,
        timeout_seconds: float = 60.0,
        max_retries: int = 2,
        store: bool = False,
        client: Any | None = None,
    ):
        if client is None:
            api_key = os.getenv("OPENAI_API_KEY")
            if not api_key:
                raise AIAnalysisError("OPENAI_API_KEY is not set")
            client = OpenAI(
                api_key=api_key,
                timeout=timeout_seconds,
                max_retries=max_retries,
            )
        self.client = client
        self.model = model
        self.max_output_tokens = max_output_tokens
        self.store = store

    def analyze(self, results: dict, question: str | None = None) -> KPIAnalysis:
        compact_payload = {
            "project_name": results.get("project_name"),
            "simulation": results.get("config", {}).get("simulation"),
            "experiment": results.get("config", {}).get("experiment"),
            "random_streams": results.get("random_streams"),
            "metric_catalog": results.get("metric_catalog"),
            "summary": results.get("summary"),
        }
        user_question = question or (
            "比较所有实验场景，指出 KPI 权衡、瓶颈、最佳场景和下一步实验建议。"
        )
        try:
            response = self.client.responses.parse(
                model=self.model,
                input=[
                    {
                        "role": "system",
                        "content": (
                            "你是离散事件仿真与运营研究专家。只根据提供的数据下结论；"
                            "明确区分观测、推断和不确定性；不要伪造因果关系。"
                            "若数据未提供成本、目标或决策门槛，不得臆断最佳方案，"
                            "应将 best_scenario 设为 null 并说明缺失信息。用中文回答。"
                        ),
                    },
                    {
                        "role": "user",
                        "content": (
                            f"问题：{user_question}\n\n"
                            f"实验数据：{json.dumps(compact_payload, ensure_ascii=False)}"
                        ),
                    },
                ],
                text_format=KPIAnalysis,
                max_output_tokens=self.max_output_tokens,
                store=self.store,
            )
        except AuthenticationError as error:
            raise AIAnalysisError("OpenAI API 认证失败，请检查 OPENAI_API_KEY。") from error
        except RateLimitError as error:
            raise AIAnalysisError("OpenAI API 当前限流或额度不足，请稍后重试。") from error
        except APITimeoutError as error:
            raise AIAnalysisError("OpenAI API 请求超时。") from error
        except APIConnectionError as error:
            raise AIAnalysisError("无法连接 OpenAI API，请检查网络。") from error
        except APIStatusError as error:
            request_id = getattr(error, "request_id", None)
            suffix = f"（request_id={request_id}）" if request_id else ""
            raise AIAnalysisError(f"OpenAI API 返回 HTTP {error.status_code}{suffix}。") from error
        except OpenAIError as error:
            raise AIAnalysisError("OpenAI API 请求失败。") from error

        status = getattr(response, "status", None)
        if status not in {None, "completed"}:
            details = getattr(response, "incomplete_details", None)
            reason = getattr(details, "reason", None)
            suffix = f"：{reason}" if reason else ""
            raise AIAnalysisError(f"OpenAI 响应状态为 {status}{suffix}")
        if response.output_parsed is None:
            raise AIAnalysisError("OpenAI 响应未包含可解析的 KPI 分析，可能是拒答或输出不完整。")
        return response.output_parsed


def save_analysis(analysis: KPIAnalysis, output_dir: str | Path) -> tuple[Path, Path]:
    root = Path(output_dir)
    root.mkdir(parents=True, exist_ok=True)
    json_path = root / "ai_analysis.json"
    markdown_path = root / "ai_analysis.md"
    with json_path.open("w", encoding="utf-8") as stream:
        json.dump(analysis.model_dump(mode="json"), stream, ensure_ascii=False, indent=2)
    with markdown_path.open("w", encoding="utf-8") as stream:
        stream.write(analysis.to_markdown())
    return json_path, markdown_path
