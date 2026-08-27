from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from pydantic import ValidationError

from simlab.ai import OpenAIKPIAnalyst, save_analysis
from simlab.config import ProjectConfig
from simlab.experiment import ExperimentRunner, expand_scenarios


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="simlab",
        description="SimPy 多实验、KPI 统计与 OpenAI 分析框架",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    validate = subparsers.add_parser("validate", help="校验 YAML 配置")
    validate.add_argument("config", type=Path)

    run = subparsers.add_parser("run", help="执行参数场景与多次 replication")
    run.add_argument("config", type=Path)
    run.add_argument("--output", type=Path, help="覆盖配置中的输出目录")
    run.add_argument("--workers", type=int, default=1, help="并行进程数，默认 1")
    run.add_argument("--analyze", action="store_true", help="完成后调用 OpenAI 分析")
    run.add_argument("--question", help="给 AI 分析器的业务问题")
    run.add_argument("--model", help="覆盖 OpenAI 模型")

    analyze = subparsers.add_parser("analyze", help="用 OpenAI 分析已有 results.json")
    analyze.add_argument("results", type=Path)
    analyze.add_argument("--question", help="希望模型回答的问题")
    analyze.add_argument("--model", help="覆盖结果配置中的模型")
    analyze.add_argument("--output", type=Path, help="分析文件输出目录")

    serve = subparsers.add_parser("serve", help="启动人工审批与仿真控制 API")
    serve.add_argument("--host", default="127.0.0.1", help="监听地址，默认仅本机")
    serve.add_argument("--port", type=int, default=8000, help="监听端口，默认 8000")
    serve.add_argument("--reload", action="store_true", help="开发时自动重载")
    return parser


def analyze_results(
    result: dict,
    output_dir: Path,
    model_override: str | None = None,
    question: str | None = None,
) -> None:
    openai_config = result.get("config", {}).get("openai", {})
    model = model_override or openai_config.get("model", "gpt-5.6")
    max_tokens = openai_config.get("max_output_tokens", 2500)
    print(f"正在使用 {model} 分析 KPI 汇总……")
    analysis = OpenAIKPIAnalyst(
        model=model,
        max_output_tokens=max_tokens,
        timeout_seconds=openai_config.get("timeout_seconds", 60.0),
        max_retries=openai_config.get("max_retries", 2),
        store=openai_config.get("store", False),
    ).analyze(result, question=question)
    json_path, markdown_path = save_analysis(analysis, output_dir)
    print(f"AI 分析已写入 {json_path} 和 {markdown_path}")


def main(argv: list[str] | None = None) -> None:
    args = build_parser().parse_args(argv)
    try:
        if args.command == "serve":
            if not 1 <= args.port <= 65535:
                raise ValueError("--port must be between 1 and 65535")
            try:
                from simlab.api import run_api
            except ImportError as error:
                raise RuntimeError(
                    'API dependencies are missing; install with pip install -e ".[api]"'
                ) from error
            run_api(host=args.host, port=args.port, reload=args.reload)
            return

        if args.command == "validate":
            config = ProjectConfig.load(args.config)
            scenarios = expand_scenarios(config)
            total = len(scenarios) * config.experiment.replications
            print(
                f"配置有效：{len(scenarios)} 个场景，"
                f"每场景 {config.experiment.replications} 次，共 {total} 次仿真。"
            )
            return

        if args.command == "run":
            if args.workers < 1:
                raise ValueError("--workers must be at least 1")
            config = ProjectConfig.load(args.config)
            runner = ExperimentRunner(config)
            tasks = len(expand_scenarios(config)) * config.experiment.replications
            print(f"开始执行 {tasks} 次仿真（workers={args.workers}）……")
            result = runner.run(workers=args.workers)
            output_dir = runner.save(result, args.output)
            print(f"完成。结果已写入 {output_dir.resolve()}")
            if args.analyze:
                analyze_results(result, output_dir, args.model, args.question)
            return

        with args.results.open(encoding="utf-8") as stream:
            result = json.load(stream)
        analyze_results(
            result,
            args.output or args.results.parent,
            model_override=args.model,
            question=args.question,
        )
    except (ValidationError, ValueError, OSError, RuntimeError, json.JSONDecodeError) as error:
        print(f"错误：{error}", file=sys.stderr)
        raise SystemExit(2) from error


if __name__ == "__main__":
    main()
