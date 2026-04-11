from __future__ import annotations

import sys

from mix2md_pipeline.models import FileRecord, PipelineState

import legacy_engine


def convert_files_step(state: PipelineState) -> PipelineState:
    for source_path in state.files:
        record = FileRecord(source_path=source_path)
        try:
            file_output_dir = legacy_engine.output_dir_for_file(
                source_path,
                state.config.input_path,
                state.base_output_dir,
            )
            outputs = legacy_engine.convert_file(source_path, file_output_dir)
            if not outputs:
                record.status = "skip"
                state.skipped += 1
                print(f"[SKIP] {source_path}")
            else:
                record.status = "ok"
                record.output_paths = outputs
                state.succeeded += 1
                print(f"[OK] {source_path}")
                for output in outputs:
                    print(f"     -> {output}")
        except Exception as exc:  # noqa: BLE001
            message = str(exc)
            record.status = "fail"
            record.error = message
            state.failures.append((source_path, message))
            print(f"[FAIL] {source_path}: {message}", file=sys.stderr)
        state.records.append(record)
    return state

