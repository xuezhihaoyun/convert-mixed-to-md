from __future__ import annotations

import sys

from mix2md_pipeline.models import PipelineState
from mix2md_pipeline.steps.convert import convert_files_step
from mix2md_pipeline.steps.discover import discover_files_step
from mix2md_pipeline.steps.preflight import preflight_checks_step
from mix2md_pipeline.steps.report import report_step


def run_pipeline(state: PipelineState) -> PipelineState:
    state = discover_files_step(state)
    if not state.files:
        print("没有找到可转换的文件（支持 .doc / .docx / .epub / .pdf / .wps / .wpt / .hwp）。", file=sys.stderr)
        state.exit_code = 1
        return state

    steps = [preflight_checks_step, convert_files_step, report_step]
    for step in steps:
        state = step(state)
    return state
