from __future__ import annotations

from mix2md_pipeline.models import PipelineState

import legacy_engine


def discover_files_step(state: PipelineState) -> PipelineState:
    config = state.config
    state.base_output_dir = legacy_engine.resolve_base_output_dir(
        config.input_path,
        config.explicit_output_dir,
    )
    state.files = legacy_engine.discover_files(config.input_path)
    state.suffixes = {path.suffix.lower() for path in state.files}
    return state

