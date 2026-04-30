#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

import legacy_engine
from mix2md_pipeline.models import PipelineConfig, PipelineState
from mix2md_pipeline.polish import ARTIFACT_LEVELS, POLISH_PROFILES
from mix2md_pipeline.runner import run_pipeline


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="批量将 .doc / .docx / .epub / .pdf / .wps / .wpt / .hwp / 常见图片 转为 Markdown，并跳过已存在的同名 .md。"
    )
    parser.add_argument("input", nargs="?", help="单个文件或目录")
    parser.add_argument("-o", "--output-dir", help="输出目录，默认写回输入文件所在目录")
    parser.add_argument(
        "--check",
        action="store_true",
        help="仅检查运行环境与关键依赖，不执行转换",
    )
    parser.add_argument(
        "--polish",
        choices=POLISH_PROFILES,
        default="none",
        help="转换后整理模式：none=关闭，auto=自动识别，legal=法律结构，judgment=裁判/裁决文书，notice=通知类文书，book=书籍结构，generic=基础清洗",
    )
    parser.add_argument(
        "--report",
        action="store_true",
        help="生成转换审核报告；启用 --polish 时会自动生成",
    )
    parser.add_argument(
        "--artifact-level",
        choices=ARTIFACT_LEVELS,
        default="minimal",
        help="过程产物保留级别：minimal|standard|debug",
    )
    return parser


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()

    if args.check:
        legacy_engine.print_environment_check()
        return 0

    if not args.input:
        parser.error("缺少 input 参数（单个文件或目录），或使用 --check 仅做环境检查")

    input_path = Path(args.input).expanduser().resolve()
    state = PipelineState(
        config=PipelineConfig(
            input_path=input_path,
            explicit_output_dir=args.output_dir,
            polish_profile=args.polish,
            write_report=args.report,
            artifact_level=args.artifact_level,
        )
    )
    final_state = run_pipeline(state)
    return final_state.exit_code


if __name__ == "__main__":
    raise SystemExit(main())
