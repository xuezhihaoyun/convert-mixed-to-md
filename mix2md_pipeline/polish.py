from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
import re
import shutil


POLISH_PROFILES = ("none", "auto", "legal", "judgment", "notice", "book", "generic")
ARTIFACT_LEVELS = ("minimal", "standard", "debug")
_CN_NUM = "零〇一二三四五六七八九十百千万两0-9"
_HEADING_PREFIX_RE = re.compile(r"^(#{1,6})[ \t]+")
_CANONICAL_HEADING_RE = re.compile(r"^(#{1,6}) ")
_ARTICLE_RE = re.compile(rf"^(\s*)(第[{_CN_NUM}]+条)(.*)$")
_ARTICLE_HEAD_RE = re.compile(rf"^\s*第[{_CN_NUM}]+条")
_PART_RE = re.compile(rf"^\s*第[{_CN_NUM}]+(?:编|分编)")
_CHAPTER_RE = re.compile(rf"^\s*第[{_CN_NUM}]+章")
_SECTION_RE = re.compile(rf"^\s*第[{_CN_NUM}]+节")
_BOOK_CHAPTER_RE = re.compile(rf"^\s*(?:第[{_CN_NUM}]+章|Chapter\s+\d+|CHAPTER\s+\d+)\b")
_BOOK_SECTION_RE = re.compile(rf"^\s*(?:第[{_CN_NUM}]+节|\d+(?:\.\d+){{1,3}}\s+)")
_COURT_DOC_RE = re.compile(r"(?:判决书|裁定书|调解书|决定书)")
_ARBITRAL_DOC_RE = re.compile(r"(?:仲裁裁决书|裁决书)")
_NOTICE_DOC_RE = re.compile(r"(?:通知书|告知书)")
_JUDGMENT_ANCHOR_RE = re.compile(
    r"^(?:本院经审理查明[:：]?|本院认为[，,:：]?|本院经审查认为[，,:：]?|"
    r"判决如下[:：]?|裁定如下[:：]?|裁决如下[:：]?|事实和理由[:：]?|相关案情[:：]?)$"
)
_NOTICE_ANCHOR_RE = re.compile(
    r"^(?:查封通知书|冻结通知书|扣押通知书|协助执行通知书|评估报告通知书|"
    r"执行通知书|通知书|告知书)$"
)
_ITEM_RE = re.compile(r"（[一二三四五六七八九十百千万零〇两]+）")
_SUBITEM_RE = re.compile(r"(?:\([0-9]{1,3}\)|[0-9]{1,3}[、\.．])")
_TRAILING_WS_RE = re.compile(r"[ \t\u3000]+$")


@dataclass
class PolishResult:
    output_path: Path
    report_path: Path | None = None
    stage1_path: Path | None = None
    stage2_path: Path | None = None
    profile: str = "none"
    detected_profile: str = "generic"
    status: str = "skipped"
    preserve_passed: bool = True
    changed: bool = False
    stats: dict[str, object] = field(default_factory=dict)

    @property
    def extra_paths(self) -> list[Path]:
        paths: list[Path] = []
        for path in (self.report_path, self.stage1_path, self.stage2_path):
            if path is not None and path.exists():
                paths.append(path)
        return paths


def _strip_heading_prefix(line: str) -> str:
    return _HEADING_PREFIX_RE.sub("", line, count=1)


def _canonical_text(text: str) -> str:
    pieces: list[str] = []
    for line in text.splitlines():
        normalized = _CANONICAL_HEADING_RE.sub("", line, count=1)
        normalized = re.sub(r"[ \t\r\n\u3000]+", "", normalized)
        pieces.append(normalized)
    return "".join(pieces)


def _split_item_and_subitem(line: str) -> tuple[list[str], int]:
    markers = [m.start() for m in _ITEM_RE.finditer(line)]
    markers.extend(m.start() for m in _SUBITEM_RE.finditer(line))
    markers = sorted(set(markers))
    if len(markers) <= 1:
        return [line], 0

    parts: list[str] = []
    last = 0
    for idx in markers[1:]:
        parts.append(line[last:idx])
        last = idx
    parts.append(line[last:])
    return parts, len(markers) - 1


def _cleanup_spaces(lines: list[str]) -> tuple[list[str], int]:
    out: list[str] = []
    changes = 0
    for line in lines:
        new_line = _TRAILING_WS_RE.sub("", line)
        new_line = new_line.lstrip(" \t　")
        if new_line != line:
            changes += 1
        if new_line.startswith("#"):
            match = re.match(r"^(#{1,6})[ \t\u3000]*(.*)$", new_line)
            if match:
                marks, content = match.groups()
                normalized = f"{marks} {content.strip()}" if content.strip() else marks
                if normalized != new_line:
                    changes += 1
                new_line = normalized
        out.append(new_line)
    return out, changes


def _squash_blank_lines(text: str) -> str:
    return re.sub(r"\n{3,}", "\n\n", text.replace("\r\n", "\n").replace("\r", "\n")).strip() + "\n"


def detect_profile(text: str) -> str:
    sample_lines = [line.strip() for line in text.splitlines()[:220] if line.strip()]
    sample = "\n".join(sample_lines)
    article_count = sum(1 for line in sample_lines if _ARTICLE_HEAD_RE.match(line))
    has_chapter = any(_CHAPTER_RE.match(line) for line in sample_lines)
    has_section = any(_SECTION_RE.match(line) for line in sample_lines)
    has_book_chapter = any(_BOOK_CHAPTER_RE.match(line) for line in sample_lines)
    has_court = "人民法院" in sample
    has_arbitral_org = any(x in sample for x in ("仲裁委员会", "仲裁院", "仲裁庭"))
    has_court_doc = bool(_COURT_DOC_RE.search(sample))
    has_arbitral_doc = bool(_ARBITRAL_DOC_RE.search(sample))
    judgment_hits = sum(
        1
        for marker in ("本院认为", "本院经审理查明", "本院经审查认为", "判决如下", "裁定如下", "裁决如下")
        if marker in sample
    )
    notice_hits = sum(
        1
        for marker in ("通知书", "查封", "冻结", "扣押", "协助执行", "评估报告", "执行通知")
        if marker in sample
    )
    if (has_court and has_court_doc) or (has_arbitral_org and has_arbitral_doc) or judgment_hits >= 2:
        return "judgment"
    if notice_hits >= 2 or (_NOTICE_DOC_RE.search(sample) and any(x in sample for x in ("人民法院", "执行", "保全", "评估"))):
        return "notice"
    if article_count >= 2 or (article_count >= 1 and has_chapter) or (has_chapter and has_section):
        return "legal"
    if has_book_chapter or re.search(r"(?m)^\s*(目录|前言|序言|后记)\s*$", sample):
        return "book"
    return "generic"


def normalize_legal(text: str) -> tuple[str, dict[str, object]]:
    lines = text.splitlines()
    had_trailing_newline = text.endswith("\n")
    structure_indexes: list[int] = []
    for idx, raw in enumerate(lines):
        content = _strip_heading_prefix(raw)
        if _ARTICLE_HEAD_RE.match(content) or _SECTION_RE.match(content) or _CHAPTER_RE.match(content) or _PART_RE.match(content):
            structure_indexes.append(idx)

    if not structure_indexes:
        return text, {"reason": "legal-structure-not-detected"}

    title_idx = -1
    first_structure_idx = min(structure_indexes)
    for idx, raw in enumerate(lines):
        if idx >= first_structure_idx:
            break
        if raw.strip():
            title_idx = idx
            break

    out: list[str] = []
    stats = {
        "title_count": 0,
        "part_count": 0,
        "chapter_count": 0,
        "section_count": 0,
        "article_count": 0,
        "item_split_count": 0,
        "space_cleanup_count": 0,
    }
    for idx, raw in enumerate(lines):
        if raw == "":
            out.append(raw)
            continue
        content = _strip_heading_prefix(raw)
        if idx == title_idx:
            out.append(f"# {content}")
            stats["title_count"] += 1
            continue
        if _PART_RE.match(content):
            out.append(f"## {content}")
            stats["part_count"] += 1
            continue
        if _CHAPTER_RE.match(content):
            out.append(f"### {content}")
            stats["chapter_count"] += 1
            continue
        if _SECTION_RE.match(content):
            out.append(f"#### {content}")
            stats["section_count"] += 1
            continue
        article_match = _ARTICLE_RE.match(content)
        if article_match:
            leading, token, rest = article_match.groups()
            if re.match(r"^\s*【[^】]+】", rest):
                out.append(f"##### {content}")
            else:
                out.append(f"##### {leading}{token}")
                if rest:
                    split_lines, split_count = _split_item_and_subitem(rest)
                    out.extend(split_lines)
                    stats["item_split_count"] += split_count
            stats["article_count"] += 1
            continue
        split_lines, split_count = _split_item_and_subitem(raw)
        out.extend(split_lines)
        stats["item_split_count"] += split_count

    out, cleanup_count = _cleanup_spaces(out)
    stats["space_cleanup_count"] = cleanup_count
    new_text = "\n".join(out)
    if had_trailing_newline:
        new_text += "\n"
    stats["reason"] = "applied" if new_text != text else "already-normalized"
    return _squash_blank_lines(new_text), stats


def normalize_judgment(text: str) -> tuple[str, dict[str, object]]:
    lines = text.splitlines()
    out: list[str] = []
    title_count = 0
    anchor_count = 0
    case_no_count = 0
    for raw in lines:
        if not raw.strip():
            out.append("")
            continue
        content = _strip_heading_prefix(raw).strip()
        compact = re.sub(r"\s+", "", content)
        if "人民法院" in compact or "仲裁委员会" in compact or "仲裁院" in compact:
            out.append(f"# {content}")
            title_count += 1
            continue
        if (_COURT_DOC_RE.search(compact) or _ARBITRAL_DOC_RE.search(compact)) and len(compact) <= 18:
            out.append(f"# {content}")
            title_count += 1
            continue
        if re.search(r"[（(]\d{4}[）)].{1,20}号", compact) and len(compact) <= 40:
            out.append(f"## {content}")
            case_no_count += 1
            continue
        if _JUDGMENT_ANCHOR_RE.match(content):
            out.append(f"## {content}")
            anchor_count += 1
            continue
        if re.match(r"^(如不服本(?:判决|裁定)|如果未按本(?:判决|裁定))", content):
            out.append(f"## {content}")
            anchor_count += 1
            continue
        out.append(raw)

    out, cleanup_count = _cleanup_spaces(out)
    new_text = _squash_blank_lines("\n".join(out))
    return new_text, {
        "reason": "applied" if new_text != text else "already-normalized",
        "title_count": title_count,
        "case_no_count": case_no_count,
        "anchor_count": anchor_count,
        "space_cleanup_count": cleanup_count,
    }


def normalize_notice(text: str) -> tuple[str, dict[str, object]]:
    lines = text.splitlines()
    out: list[str] = []
    title_count = 0
    case_no_count = 0
    anchor_count = 0
    for raw in lines:
        if not raw.strip():
            out.append("")
            continue
        content = _strip_heading_prefix(raw).strip()
        compact = re.sub(r"\s+", "", content)
        if "人民法院" in compact and len(compact) <= 30:
            out.append(f"# {content}")
            title_count += 1
            continue
        if _NOTICE_ANCHOR_RE.match(compact) or (_NOTICE_DOC_RE.search(compact) and len(compact) <= 24):
            out.append(f"# {content}")
            title_count += 1
            continue
        if re.search(r"[（(]\d{4}[）)].{1,20}号", compact) and len(compact) <= 40:
            out.append(f"## {content}")
            case_no_count += 1
            continue
        out.append(raw)

    out, cleanup_count = _cleanup_spaces(out)
    new_text = _squash_blank_lines("\n".join(out))
    return new_text, {
        "reason": "applied" if new_text != text else "already-normalized",
        "title_count": title_count,
        "case_no_count": case_no_count,
        "anchor_count": anchor_count,
        "space_cleanup_count": cleanup_count,
    }


def normalize_book(text: str) -> tuple[str, dict[str, object]]:
    lines = text.splitlines()
    out: list[str] = []
    title_count = 0
    chapter_count = 0
    section_count = 0
    for raw in lines:
        if not raw.strip():
            out.append("")
            continue
        content = _strip_heading_prefix(raw).strip()
        if title_count == 0 and not _BOOK_CHAPTER_RE.match(content) and not re.match(r"^(目录|前言|序言|后记)$", content):
            out.append(f"# {content}")
            title_count += 1
            continue
        if re.match(r"^(目录|前言|序言|后记)$", content):
            out.append(f"## {content}")
            section_count += 1
            continue
        if _BOOK_CHAPTER_RE.match(content):
            out.append(f"## {content}")
            chapter_count += 1
            continue
        if _BOOK_SECTION_RE.match(content):
            out.append(f"### {content}")
            section_count += 1
            continue
        out.append(raw)

    out, cleanup_count = _cleanup_spaces(out)
    new_text = _squash_blank_lines("\n".join(out))
    return new_text, {
        "reason": "applied" if new_text != text else "already-normalized",
        "title_count": title_count,
        "chapter_count": chapter_count,
        "section_count": section_count,
        "space_cleanup_count": cleanup_count,
    }


def normalize_generic(text: str) -> tuple[str, dict[str, object]]:
    lines, cleanup_count = _cleanup_spaces(text.splitlines())
    new_text = _squash_blank_lines("\n".join(lines))
    return new_text, {
        "reason": "applied" if new_text != text else "already-normalized",
        "space_cleanup_count": cleanup_count,
    }


def _report_path(output_path: Path) -> Path:
    return output_path.with_name(f"{output_path.stem}+审核报告.md")


def _stage_path(output_path: Path, stage: str) -> Path:
    return output_path.with_name(f"{output_path.stem}.{stage}{output_path.suffix or '.md'}")


def _write_report(result: PolishResult) -> None:
    assert result.report_path is not None
    lines = [
        "# 文档转换审核报告",
        "",
        "## 基本信息",
        "",
        f"- 生成时间：{datetime.now().isoformat(timespec='seconds')}",
        f"- 输出文件：{result.output_path}",
        f"- 请求模式：{result.profile}",
        f"- 实际模式：{result.detected_profile}",
        f"- 审核结论：{result.status}",
        f"- 保真检查：{'通过' if result.preserve_passed else '未通过'}",
        f"- 是否改动：{'是' if result.changed else '否'}",
        "",
        "## 处理统计",
        "",
    ]
    if result.stats:
        for key, value in result.stats.items():
            lines.append(f"- {key}: {value}")
    else:
        lines.append("- 无")
    lines.extend(["", "## 产物", ""])
    for label, path in (
        ("最终 Markdown", result.output_path),
        ("Stage1", result.stage1_path),
        ("Stage2", result.stage2_path),
    ):
        if path is not None and path.exists():
            lines.append(f"- {label}: {path}")
    result.report_path.parent.mkdir(parents=True, exist_ok=True)
    result.report_path.write_text("\n".join(lines).rstrip() + "\n", encoding="utf-8")


def apply_polish(
    output_path: Path,
    *,
    profile: str = "none",
    write_report: bool = False,
    artifact_level: str = "minimal",
) -> PolishResult:
    if profile not in POLISH_PROFILES:
        profile = "none"
    if artifact_level not in ARTIFACT_LEVELS:
        artifact_level = "minimal"

    result = PolishResult(output_path=output_path, profile=profile)
    if profile == "none" and not write_report:
        return result

    original = output_path.read_text(encoding="utf-8")
    if profile == "none":
        result.detected_profile = "none"
        result.status = "approved_no_change"
        result.preserve_passed = True
        result.changed = False
        result.stats = {"reason": "report-only"}
        result.report_path = _report_path(output_path)
        _write_report(result)
        return result

    detected = detect_profile(original) if profile == "auto" else (profile if profile != "none" else "generic")
    result.detected_profile = detected

    if detected == "legal":
        polished, stats = normalize_legal(original)
    elif detected == "judgment":
        polished, stats = normalize_judgment(original)
    elif detected == "notice":
        polished, stats = normalize_notice(original)
    elif detected == "book":
        polished, stats = normalize_book(original)
    else:
        polished, stats = normalize_generic(original)

    preserve_passed = _canonical_text(original) == _canonical_text(polished)
    result.preserve_passed = preserve_passed
    result.changed = polished != original
    result.stats = stats
    result.report_path = _report_path(output_path)

    keep_stage = artifact_level in {"standard", "debug"} or not preserve_passed
    if keep_stage:
        result.stage1_path = _stage_path(output_path, "stage1")
        result.stage2_path = _stage_path(output_path, "stage2")
        shutil.copyfile(output_path, result.stage1_path)
        result.stage2_path.write_text(polished, encoding="utf-8")

    if preserve_passed:
        if result.changed:
            output_path.write_text(polished, encoding="utf-8")
            result.status = "approved"
        else:
            result.status = "approved_no_change"
    else:
        result.status = "rejected_preserve_check_failed"

    if write_report or profile != "none":
        _write_report(result)

    return result
