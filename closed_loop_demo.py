"""闭环演示驱动：建会话 → 跑仿真 → DeepSeek 生成提案 → 人工审批 → 再跑仿真。

用法：
  python closed_loop_demo.py                # 阶段一：建会话、跑基线、生成 AI 提案
  python closed_loop_demo.py --approve PID  # 阶段二：批准某条提案并再跑仿真对比
  python closed_loop_demo.py --reject PID   # 阶段二：拒绝某条提案
  python closed_loop_demo.py --audit        # 查看会话审计日志
"""
from __future__ import annotations

import argparse
import json
import sys
import uuid
from pathlib import Path

import httpx

BASE_URL = "http://127.0.0.1:8000"
TOKEN = "simlab-demo-token"
HEADERS = {"Authorization": f"Bearer {TOKEN}"}
STATE_PATH = Path("outputs/api_sessions/demo_state.json")
CONFIG_PATH = Path("examples/api_service_center.json")
OBJECTIVE = (
    "在专家工位利用率不低于 0.6 的前提下，"
    "降低平均周期时间和等待时间，提高服务水平。"
)
KEY_METRICS = [
    "avg_cycle_time",
    "p95_cycle_time",
    "avg_wait_time",
    "service_level",
    "station.specialist.utilization",
]


def operation_id(tag: str) -> str:
    return f"demo-{tag}-{uuid.uuid4().hex[:12]}"


def api(method: str, path: str, **kwargs):
    with httpx.Client(base_url=BASE_URL, headers=HEADERS, timeout=180.0) as client:
        response = client.request(method, path, **kwargs)
    if response.status_code >= 400:
        print(f"[HTTP {response.status_code}] {response.text}", file=sys.stderr)
        sys.exit(1)
    return response.json()


def load_config() -> dict:
    config = json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
    config["openai"]["model"] = "deepseek-chat"
    return config


def current_version(session_id: str) -> int:
    return api("GET", f"/v1/sessions/{session_id}")["workflow_version"]


def print_kpi_summary(title: str, result: dict) -> None:
    print(f"\n===== {title} =====")
    summary = result.get("summary", [])
    for row in summary:
        if row.get("metric") in KEY_METRICS:
            mean = row.get("mean")
            print(
                f"  {row['metric']:<32} mean={mean:.3f}  "
                f"95%CI=[{row['ci_low']:.3f}, {row['ci_high']:.3f}]"
                if mean is not None
                else f"  {row['metric']:<32} mean=null"
            )


def run_experiment(session_id: str) -> dict:
    version = current_version(session_id)
    record = api(
        "POST",
        f"/v1/sessions/{session_id}/runs",
        json={
            "operation_id": operation_id("run"),
            "expected_version": version,
            "workers": 1,
        },
    )
    print(f"仿真运行完成：{record['run_id']}（状态 {record['status']}）")
    return api(
        "GET",
        f"/v1/sessions/{session_id}/runs/{record['run_id']}",
        params={"include_result": "true"},
    )


def phase_create_and_generate() -> None:
    session = api("POST", "/v1/sessions", json={"config": load_config()})
    session_id = session["session_id"]
    print(f"会话已创建：{session_id}（版本 {session['workflow_version']}）")

    baseline = run_experiment(session_id)
    print_kpi_summary("基线仿真（专家容量=2）", baseline["result"])

    allowed = api("GET", f"/v1/sessions/{session_id}/allowed-actions")
    print("\n当前白名单动作：")
    for action in allowed:
        print(f"  - {action['action_type']}:{action['target']} — {action['description']}")

    print("\n正在调用 DeepSeek 生成优化提案……")
    plan = api(
        "POST",
        f"/v1/sessions/{session_id}/proposals:generate",
        json={
            "operation_id": operation_id("ai"),
            "expected_version": current_version(session_id),
            "objective": OBJECTIVE,
        },
    )
    print(f"\nAI 计划：{plan['plan_id']}")
    print(f"目标：{plan['objective']}")
    print(f"摘要：{plan['summary']}")
    for caveat in plan["caveats"]:
        print(f"  ⚠ {caveat}")
    print("\n提案列表（待人工审批）：")
    for proposal in plan["proposals"]:
        action = proposal["action"]
        print(
            f"  [{proposal['proposal_id']}] {action['action']} "
            f"target={action.get('path') or action.get('policy') or action.get('resource')} "
            f"value={action.get('value')}"
        )
    STATE_PATH.write_text(
        json.dumps(
            {
                "session_id": session_id,
                "plan_id": plan["plan_id"],
                "proposal_ids": [p["proposal_id"] for p in plan["proposals"]],
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print(f"\n状态已保存到 {STATE_PATH}")
    print("请人工审批后运行：python closed_loop_demo.py --approve <提案ID>（或 --reject）")


def phase_decide(proposal_id: str, approve: bool) -> None:
    state = json.loads(STATE_PATH.read_text(encoding="utf-8"))
    session_id = state["session_id"]
    verb = "approve" if approve else "reject"
    result = api(
        "POST",
        f"/v1/sessions/{session_id}/proposals/{proposal_id}:{verb}",
        json={
            "operation_id": operation_id(verb),
            "expected_version": current_version(session_id),
            "reason": "闭环演示人工审批",
        },
    )
    print(f"提案 {proposal_id} 状态：{result['status']}")
    if result["status"] != "applied":
        return
    print("\n提案已应用，用新配置再跑一轮仿真……")
    after = run_experiment(session_id)
    print_kpi_summary("应用提案后的仿真", after["result"])


def phase_audit() -> None:
    state = json.loads(STATE_PATH.read_text(encoding="utf-8"))
    session_id = state["session_id"]
    audit = api("GET", f"/v1/sessions/{session_id}/audit")
    print("\n===== 审计日志（workflow）=====")
    for entry in audit["workflow"]:
        print(
            f"  #{entry['sequence']} v{entry['version']} {entry['event']} "
            f"{entry['proposal_id']} by {entry['actor']}"
        )
    print("===== 服务事件 =====")
    for event in audit["events"]:
        print(f"  #{event['sequence']} {event['event']}")


def main() -> None:
    parser = argparse.ArgumentParser(description="SimPy KPI Lab 闭环演示")
    parser.add_argument("--approve", help="批准提案 ID")
    parser.add_argument("--reject", help="拒绝提案 ID")
    parser.add_argument("--audit", action="store_true", help="查看审计日志")
    args = parser.parse_args()

    if args.audit:
        phase_audit()
    elif args.approve or args.reject:
        phase_decide(args.approve or args.reject, approve=bool(args.approve))
    else:
        phase_create_and_generate()


if __name__ == "__main__":
    main()
