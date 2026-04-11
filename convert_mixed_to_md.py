#!/usr/bin/env python3
from __future__ import annotations

import argparse
from pathlib import Path

import legacy_engine
from mix2md_pipeline.models import PipelineConfig, PipelineState
from mix2md_pipeline.runner import run_pipeline


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="批量将 .doc / .docx / .epub / .pdf / .wps / .wpt / .hwp 转为 Markdown，并跳过已存在的同名 .md。"
    )
    parser.add_argument("input", nargs="?", help="单个文件或目录")
    parser.add_argument("-o", "--output-dir", help="输出目录，默认写回输入文件所在目录")
    parser.add_argument(
        "--check",
        action="store_true",
        help="仅检查运行环境与关键依赖，不执行转换",
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
        )
    )
    final_state = run_pipeline(state)
    return final_state.exit_code


if __name__ == "__main__":
    raise SystemExit(main())

