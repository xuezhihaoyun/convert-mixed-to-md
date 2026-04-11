from __future__ import annotations

import os
import sys

from mix2md_pipeline.models import PipelineState

import legacy_engine


def preflight_checks_step(state: PipelineState) -> PipelineState:
    suffixes = state.suffixes
    if suffixes.intersection({".doc", ".docx", ".epub", ".wps", ".wpt"}) and not legacy_engine.which_cached("pandoc"):
        print("[WARN] 未检测到 pandoc，doc/docx/epub/wps/wpt 可能转换失败。", file=sys.stderr)
    if ".pdf" in suffixes and not legacy_engine.which_cached("pdftotext"):
        print("[INFO] 未检测到 pdftotext，将使用 Python 提取器处理 PDF。", file=sys.stderr)
    if os.name == "nt" and suffixes.intersection({".doc", ".wps", ".wpt"}) and not legacy_engine.which_cached("textutil"):
        print("[INFO] Windows 不支持 textutil；旧版文档建议先转为 .docx。", file=sys.stderr)
    if ".hwp" in suffixes and not legacy_engine.get_hwp5txt_runner():
        print("[INFO] 检测到 .hwp：首次转换会自动安装 pyhwp，耗时会稍长。", file=sys.stderr)
    return state

