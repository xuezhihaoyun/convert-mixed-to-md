from __future__ import annotations

import sys

from mix2md_pipeline.models import PipelineState


def report_step(state: PipelineState) -> PipelineState:
    if state.failures:
        print(
            f"\n完成，成功 {state.succeeded} 个，跳过 {state.skipped} 个，失败 {len(state.failures)} 个。",
            file=sys.stderr,
        )
        state.exit_code = 2
    else:
        print(f"\n完成，成功 {state.succeeded} 个，跳过 {state.skipped} 个。")
        state.exit_code = 0
    return state

