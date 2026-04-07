#!/usr/bin/env python3
from __future__ import annotations

import argparse
import hashlib
import html
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import requests


SUPPORTED_EXTENSIONS = {".doc", ".docx", ".epub", ".pdf", ".wps", ".wpt", ".hwp"}
HTML_START = b"<html><body>"
HTML_END = b"</html>"
HTML_ANY_START = b"<html"
CHARSET_RE = re.compile(r"charset\s*=\s*['\"]?([A-Za-z0-9._-]+)", re.IGNORECASE)
MINERU_API_BASE = "https://mineru.net/api/v4"
DEFAULT_MINERU_TOKEN = ""
OCR_CHUNK_PAGE_THRESHOLD = 40
OCR_CHUNK_SIZE = 25
DIRECT_OCR_FILESIZE_THRESHOLD = 50 * 1024 * 1024
DIRECT_OCR_PAGE_THRESHOLD = 120


def normalize_markdown(markdown: str, title: str | None = None) -> str:
    markdown = markdown.replace("\r\n", "\n").replace("\r", "\n").replace("\f", "\n\n")
    lines = [line.rstrip() for line in markdown.split("\n")]
    markdown = "\n".join(lines).strip()
    markdown = re.sub(r"\n{3,}", "\n\n", markdown)
    first_nonempty = next((line for line in markdown.split("\n") if line.strip()), "")
    if title and not first_nonempty.startswith("#"):
        markdown = f"# {title}\n\n{markdown}"
    return markdown.rstrip() + "\n"


def polish_mineru_markdown(markdown: str, title: str) -> str:
    text = markdown.replace("\r\n", "\n").replace("\r", "\n")

    text = re.sub(r"(?m)^[ \t]*## Part \d+[ \t]*\n?", "", text)

    start = re.search(r"(?m)^# 序\s*$", text)
    if start:
        text = f"# {title}\n\n{text[start.start():].lstrip()}"
    else:
        text = normalize_markdown(text, title=title)

    text = re.sub(r"(?m)^#\s*前\s*言\s*$", "# 前言", text)

    table_pattern = re.compile(r"<table>.*?</table>", re.S)

    def table_to_md(match: re.Match[str]) -> str:
        raw = match.group(0)
        try:
            root = ET.fromstring(raw)
        except ET.ParseError:
            return raw

        rows: list[list[str]] = []
        for tr in root.findall(".//tr"):
            row: list[str] = []
            for cell in tr:
                cell_text = "".join(cell.itertext())
                cell_text = html.unescape(cell_text)
                cell_text = re.sub(r"\s*\n\s*", " ", cell_text)
                cell_text = re.sub(r"\s{2,}", " ", cell_text).strip()
                cell_text = cell_text.replace("|", "\\|")
                row.append(cell_text)
            if row:
                rows.append(row)

        if not rows:
            return raw

        width = max(len(row) for row in rows)
        rows = [row + [""] * (width - len(row)) for row in rows]
        header = rows[0]
        sep = ["---"] * width
        lines = [
            "| " + " | ".join(header) + " |",
            "| " + " | ".join(sep) + " |",
        ]
        for row in rows[1:]:
            lines.append("| " + " | ".join(row) + " |")
        return "\n" + "\n".join(lines) + "\n"

    text = table_pattern.sub(table_to_md, text)

    spacing_patterns = [
        (re.compile(r"(?<=第)(\d+)\s+号"), r"\1号"),
        (re.compile(r"(?<=知民终)(\d+)\s+号"), r"\1号"),
        (re.compile(r"(?<=民终)(\d+)\s+号"), r"\1号"),
        (re.compile(r"(?<=民初)(\d+)\s+号"), r"\1号"),
        (re.compile(r"(?<=行终)(\d+)\s+号"), r"\1号"),
        (re.compile(r"(?<=刑终)(\d+)\s+号"), r"\1号"),
        (re.compile(r"(?<=刑初)(\d+)\s+号"), r"\1号"),
    ]
    for pattern, repl in spacing_patterns:
        text = pattern.sub(repl, text)

    replacements = [
        ("一—", "——"),
        ("———", "——"),
        (" ,", "，"),
        (" 。", "。"),
        (" ；", "；"),
        (" ：", "："),
        (" ）", "）"),
        ("（ ", "（"),
    ]
    for old, new in replacements:
        text = text.replace(old, new)

    return normalize_markdown(text, title=title)


def variant_exists(source_path: Path, output_path: Path) -> bool:
    del source_path
    return output_path.exists()


def output_markdown_path(source_path: Path, output_dir: Path) -> Path:
    sibling_conflict = any(
        sibling != source_path
        and sibling.is_file()
        and sibling.stem == source_path.stem
        and sibling.suffix.lower() in SUPPORTED_EXTENSIONS
        for sibling in source_path.parent.iterdir()
    )
    if sibling_conflict:
        return output_dir / f"{source_path.name}.md"
    return output_dir / f"{source_path.stem}.md"


def discover_files(input_path: Path) -> list[Path]:
    if input_path.is_file():
        return [input_path] if input_path.suffix.lower() in SUPPORTED_EXTENSIONS else []
    return sorted(
        path
        for path in input_path.rglob("*")
        if path.is_file() and path.suffix.lower() in SUPPORTED_EXTENSIONS
    )


def resolve_base_output_dir(input_path: Path, explicit_output_dir: str | None) -> Path | None:
    if explicit_output_dir:
        return Path(explicit_output_dir).expanduser().resolve()
    return None


def output_dir_for_file(source_path: Path, input_root: Path, base_output_dir: Path | None) -> Path:
    if base_output_dir is None:
        return source_path.parent
    if input_root.is_file():
        return base_output_dir
    relative_parent = source_path.parent.relative_to(input_root)
    return base_output_dir / relative_parent


def extract_html_bytes(doc_path: Path) -> bytes:
    data = doc_path.read_bytes()
    start = data.find(HTML_START)
    if start != -1:
        end = data.find(HTML_END, start)
        if end == -1:
            raise ValueError("未找到内嵌 HTML 结束标记")
        return data[start : end + len(HTML_END)]

    start = data.lower().find(HTML_ANY_START)
    if start != -1:
        end = data.lower().rfind(HTML_END)
        if end == -1:
            return data[start:]
        return data[start : end + len(HTML_END)]

    raise ValueError("未找到可识别的 HTML 内容")


def detect_charset(raw_html: bytes) -> str | None:
    head = raw_html[:2048].decode("ascii", errors="ignore")
    match = CHARSET_RE.search(head)
    if not match:
        return None
    return match.group(1)


def decode_legacy_doc_html(doc_path: Path) -> str:
    raw_html = extract_html_bytes(doc_path)
    preferred = detect_charset(raw_html)
    candidates = []
    if preferred:
        candidates.append(preferred)
    candidates.extend(["gb18030", "gbk", "utf-8"])

    tried: set[str] = set()
    for encoding in candidates:
        normalized = encoding.lower()
        if normalized in tried:
            continue
        tried.add(normalized)
        try:
            return raw_html.decode(encoding)
        except UnicodeDecodeError:
            continue
    return raw_html.decode("gb18030", errors="replace")


def normalize_html(html: str) -> str:
    html = re.sub(r"\sstyle=\"[^\"]*\"", "", html, flags=re.IGNORECASE)
    html = re.sub(r"<div\b[^>]*>", "<p>", html, flags=re.IGNORECASE)
    html = re.sub(r"</div>", "</p>", html, flags=re.IGNORECASE)
    return html


def run_command(cmd: list[str]) -> None:
    try:
        subprocess.run(cmd, check=True)
    except FileNotFoundError as exc:
        raise ValueError(f"未找到系统命令: {cmd[0]}") from exc


def run_pandoc_to_md(
    input_path: Path,
    output_path: Path,
    input_format: str | None = None,
    media_dir: Path | None = None,
) -> None:
    cmd = ["pandoc", str(input_path)]
    if input_format:
        cmd.extend(["-f", input_format])
    cmd.extend(["-t", "gfm", "--wrap=none"])
    if media_dir:
        cmd.append(f"--extract-media={media_dir}")
    cmd.extend(["-o", str(output_path)])
    run_command(cmd)


def convert_legacy_word_to_md(source_path: Path, output_path: Path) -> None:
    fallback_errors: list[str] = []

    try:
        html = normalize_html(decode_legacy_doc_html(source_path))
        with tempfile.TemporaryDirectory(prefix="doc_html_") as temp_dir:
            html_path = Path(temp_dir) / f"{source_path.stem}.html"
            html_path.write_text(html, encoding="utf-8")
            run_pandoc_to_md(html_path, output_path, input_format="html")
            return
    except Exception as exc:  # noqa: BLE001
        fallback_errors.append(f"html 提取失败: {exc}")

    textutil = shutil.which("textutil")
    if textutil:
        try:
            result = subprocess.run(
                [textutil, "-convert", "txt", "-stdout", str(source_path)],
                check=True,
                capture_output=True,
                text=True,
            )
            text = result.stdout.strip()
            if text:
                output_path.write_text(normalize_markdown(text, title=source_path.stem), encoding="utf-8")
                return
        except Exception as exc:  # noqa: BLE001
            fallback_errors.append(f"textutil 提取失败: {exc}")

    antiword = shutil.which("antiword")
    if antiword:
        try:
            result = subprocess.run(
                [antiword, str(source_path)],
                check=True,
                capture_output=True,
                text=True,
            )
            text = result.stdout.strip()
            if text:
                output_path.write_text(normalize_markdown(text, title=source_path.stem), encoding="utf-8")
                return
        except Exception as exc:  # noqa: BLE001
            fallback_errors.append(f"antiword 提取失败: {exc}")

    catdoc = shutil.which("catdoc")
    if catdoc:
        try:
            result = subprocess.run(
                [catdoc, str(source_path)],
                check=True,
                capture_output=True,
                text=True,
            )
            text = result.stdout.strip()
            if text:
                output_path.write_text(normalize_markdown(text, title=source_path.stem), encoding="utf-8")
                return
        except Exception as exc:  # noqa: BLE001
            fallback_errors.append(f"catdoc 提取失败: {exc}")

    try:
        run_pandoc_to_md(source_path, output_path)
        if output_path.exists() and output_path.read_text(encoding="utf-8", errors="ignore").strip():
            return
    except Exception as exc:  # noqa: BLE001
        fallback_errors.append(f"pandoc 直接转换失败: {exc}")

    hints: list[str] = []
    if os.name == "nt":
        hints.append("Windows 上建议先将 .doc/.wps/.wpt 转为 .docx 再转换")
    if not shutil.which("pandoc"):
        hints.append("未安装 pandoc")
    if not textutil and not antiword and not catdoc:
        hints.append("未找到可用的旧版文档提取工具（textutil/antiword/catdoc）")

    detail = "; ".join(fallback_errors[:4]) if fallback_errors else "无详细错误"
    hint_text = f"；建议：{'；'.join(hints)}" if hints else ""
    raise ValueError(f"旧版文档转换失败：{detail}{hint_text}")


def convert_docx_or_epub_to_md(source_path: Path, output_path: Path) -> None:
    media_dir = output_path.parent / f"{output_path.stem}_assets"
    run_pandoc_to_md(source_path, output_path, media_dir=media_dir)


def get_hwp5txt_runner() -> list[str] | None:
    hwp5txt = shutil.which("hwp5txt")
    if hwp5txt:
        return [hwp5txt]

    probe = subprocess.run(
        [sys.executable, "-m", "hwp5.hwp5txt", "--help"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )
    if probe.returncode == 0:
        return [sys.executable, "-m", "hwp5.hwp5txt"]
    return None


def ensure_hwp5txt_runner() -> list[str]:
    runner = get_hwp5txt_runner()
    if runner:
        return runner

    install_cmd = [sys.executable, "-m", "pip", "install", "pyhwp", "six"]
    try:
        subprocess.run(install_cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    except Exception as exc:  # noqa: BLE001
        raise ValueError(
            "未找到 HWP 转换器（hwp5txt），且自动安装 pyhwp 失败。"
            "可手动执行: python -m pip install pyhwp six"
        ) from exc

    runner = get_hwp5txt_runner()
    if not runner:
        raise ValueError("已安装 pyhwp，但仍未找到可用 hwp5txt 命令")
    return runner


def convert_hwp_to_md(source_path: Path, output_path: Path) -> None:
    runner = ensure_hwp5txt_runner()

    with tempfile.TemporaryDirectory(prefix="hwp_text_") as temp_dir:
        txt_path = Path(temp_dir) / f"{source_path.stem}.txt"
        cmd = [*runner, str(source_path), "--output", str(txt_path)]
        try:
            subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True)
        except subprocess.CalledProcessError as exc:
            detail = (exc.stderr or "").strip()
            if detail:
                raise ValueError(f"HWP 提取失败: {detail}") from exc
            raise ValueError("HWP 提取失败（hwp5txt 执行异常）") from exc

        text = txt_path.read_text(encoding="utf-8", errors="ignore").strip()
        if not text:
            raise ValueError("HWP 提取到的文本为空")
        output_path.write_text(normalize_markdown(text, title=source_path.stem), encoding="utf-8")


def extract_pdf_text_with_pdftotext(pdf_path: Path) -> str:
    pdftotext = shutil.which("pdftotext")
    if not pdftotext:
        return ""

    with tempfile.TemporaryDirectory(prefix="pdf_text_") as temp_dir:
        txt_path = Path(temp_dir) / f"{pdf_path.stem}.txt"
        cmd = [pdftotext, "-enc", "UTF-8", str(pdf_path), str(txt_path)]
        subprocess.run(cmd, check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        return txt_path.read_text(encoding="utf-8", errors="ignore")


def extract_pdf_text_with_pdfplumber(pdf_path: Path) -> str:
    import pdfplumber

    pages: list[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            text = page.extract_text() or ""
            if text.strip():
                pages.append(text.strip())
    return "\n\n".join(pages)


def extract_pdf_text_with_pypdf(pdf_path: Path) -> str:
    from pypdf import PdfReader

    pages: list[str] = []
    reader = PdfReader(str(pdf_path))
    for page in reader.pages:
        text = page.extract_text() or ""
        if text.strip():
            pages.append(text.strip())
    return "\n\n".join(pages)


def mineru_headers(token: str) -> dict[str, str]:
    return {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json",
        "Accept": "*/*",
    }


def build_mineru_data_id(source_path: Path) -> str:
    suffix = source_path.suffix.lower() or ".bin"
    base = re.sub(r"[^A-Za-z0-9._-]+", "_", source_path.stem).strip("._-")
    if not base:
        base = "file"
    digest = hashlib.sha1(str(source_path).encode("utf-8")).hexdigest()[:16]
    data_id = f"{base[:40]}_{digest}{suffix}"
    return data_id[:128]


def extract_markdown_from_zip(zip_path: Path) -> str:
    with zipfile.ZipFile(zip_path) as zf:
        names = zf.namelist()
        preferred = [name for name in names if name.endswith("/full.md") or name == "full.md"]
        if not preferred:
            preferred = [name for name in names if name.endswith(".md")]
        if not preferred:
            raise ValueError("MinerU 结果压缩包中未找到 Markdown 文件")
        with zf.open(preferred[0]) as fp:
            return fp.read().decode("utf-8", errors="ignore")


def extract_markdown_with_mineru(source_path: Path) -> str:
    token = os.environ.get("MINERU_TOKEN", DEFAULT_MINERU_TOKEN)
    if not token:
        raise ValueError("未配置 MinerU token")

    with tempfile.TemporaryDirectory(prefix="mineru_") as temp_dir:
        temp_dir_path = Path(temp_dir)
        payload = {
            "files": [
                {
                    "name": source_path.name,
                    "data_id": build_mineru_data_id(source_path),
                    "is_ocr": True,
                }
            ],
            "model_version": "vlm",
            "enable_formula": False,
            "language": "ch",
        }
        response = requests.post(
            f"{MINERU_API_BASE}/file-urls/batch",
            headers=mineru_headers(token),
            json=payload,
            timeout=60,
        )
        response.raise_for_status()
        result = response.json()
        if result.get("code") != 0:
            raise ValueError(f"MinerU 申请上传链接失败: {result.get('msg', 'unknown error')}")

        batch_data = result["data"]
        batch_id = batch_data["batch_id"]
        upload_url = batch_data["file_urls"][0]

        with source_path.open("rb") as fp:
            upload_res = requests.put(upload_url, data=fp, timeout=300)
            upload_res.raise_for_status()

        status_url = f"{MINERU_API_BASE}/extract-results/batch/{batch_id}"
        for _ in range(120):
            status_res = requests.get(status_url, headers=mineru_headers(token), timeout=60)
            status_res.raise_for_status()
            status_json = status_res.json()
            if status_json.get("code") != 0:
                raise ValueError(f"MinerU 查询失败: {status_json.get('msg', 'unknown error')}")

            extract_results = status_json.get("data", {}).get("extract_result", [])
            if not extract_results:
                time.sleep(5)
                continue

            item = extract_results[0]
            state = item.get("state")
            if state == "done":
                zip_url = item.get("full_zip_url")
                if not zip_url:
                    raise ValueError("MinerU 结果缺少 full_zip_url")
                zip_path = temp_dir_path / f"{source_path.stem}.zip"
                zip_res = requests.get(zip_url, timeout=300)
                zip_res.raise_for_status()
                zip_path.write_bytes(zip_res.content)
                return extract_markdown_from_zip(zip_path)
            if state == "failed":
                raise ValueError(f"MinerU 解析失败: {item.get('err_msg', 'unknown error')}")
            time.sleep(5)

    raise ValueError("MinerU 解析超时")


def convert_with_mineru(source_path: Path, output_path: Path) -> bool:
    try:
        markdown = extract_markdown_with_mineru(source_path)
    except Exception:
        return False
    polished = polish_mineru_markdown(markdown, source_path.stem)
    output_path.write_text(polished, encoding="utf-8")
    return True


def get_pdf_page_count(pdf_path: Path) -> int:
    from pypdf import PdfReader

    reader = PdfReader(str(pdf_path))
    return len(reader.pages)


def should_prefer_direct_ocr(pdf_path: Path) -> bool:
    try:
        if pdf_path.stat().st_size >= DIRECT_OCR_FILESIZE_THRESHOLD:
            return True
    except OSError:
        pass

    try:
        return get_pdf_page_count(pdf_path) >= DIRECT_OCR_PAGE_THRESHOLD
    except Exception:
        return False


def split_pdf_for_ocr(source_path: Path, temp_dir: Path, chunk_size: int) -> list[Path]:
    from pypdf import PdfReader, PdfWriter

    reader = PdfReader(str(source_path))
    total_pages = len(reader.pages)
    chunks: list[Path] = []
    for start in range(0, total_pages, chunk_size):
        end = min(start + chunk_size, total_pages)
        writer = PdfWriter()
        for index in range(start, end):
            writer.add_page(reader.pages[index])
        chunk_path = temp_dir / f"{source_path.stem}.part_{start + 1:04d}_{end:04d}.pdf"
        with chunk_path.open("wb") as fp:
            writer.write(fp)
        chunks.append(chunk_path)
    return chunks


def convert_pdf_with_chunked_mineru(source_path: Path, output_path: Path) -> bool:
    try:
        total_pages = get_pdf_page_count(source_path)
    except Exception:
        return False

    if total_pages <= OCR_CHUNK_PAGE_THRESHOLD:
        return False

    with tempfile.TemporaryDirectory(prefix="pdf_ocr_chunks_") as temp_dir:
        chunk_paths = split_pdf_for_ocr(source_path, Path(temp_dir), OCR_CHUNK_SIZE)
        print(
            f"[INFO] 大 PDF 触发分段 OCR: {source_path} ({total_pages} 页, {len(chunk_paths)} 段)",
            file=sys.stderr,
        )
        parts: list[str] = []
        empty_chunks: list[int] = []
        for i, chunk_path in enumerate(chunk_paths, start=1):
            print(f"[INFO] OCR 第 {i}/{len(chunk_paths)} 段: {chunk_path.name}", file=sys.stderr)
            markdown = extract_markdown_with_mineru(chunk_path)
            cleaned = normalize_markdown(markdown).strip()
            if cleaned:
                parts.append(f"## Part {i}\n\n{cleaned}")
            else:
                empty_chunks.append(i)

    if not parts:
        return False
    if empty_chunks:
        raise ValueError(f"分段 OCR 结果缺失：第 {', '.join(map(str, empty_chunks))} 段为空")

    merged = "\n\n".join(parts)
    polished = polish_mineru_markdown(merged, source_path.stem)
    output_path.write_text(polished, encoding="utf-8")
    return True


def convert_pdf_to_md(source_path: Path, output_path: Path) -> None:
    text = ""
    errors: list[str] = []

    try:
        text = extract_pdf_text_with_pdftotext(source_path)
    except Exception as exc:  # noqa: BLE001
        errors.append(f"{extract_pdf_text_with_pdftotext.__name__}: {exc}")

    if text.strip():
        output_path.write_text(normalize_markdown(text, title=source_path.stem), encoding="utf-8")
        return

    if should_prefer_direct_ocr(source_path):
        print(f"[INFO] 大扫描 PDF，跳过慢速文本提取，直接进入 OCR: {source_path}", file=sys.stderr)
        try:
            if convert_pdf_with_chunked_mineru(source_path, output_path):
                return
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{convert_pdf_with_chunked_mineru.__name__}: {exc}")
        if convert_with_mineru(source_path, output_path):
            return
        detail = "; ".join(errors) if errors else "未提取到文本"
        raise ValueError(f"PDF 未提取到可用文本，且 OCR 失败。{detail}")

    extractors = (
        extract_pdf_text_with_pdfplumber,
        extract_pdf_text_with_pypdf,
    )

    for extractor in extractors:
        try:
            text = extractor(source_path)
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{extractor.__name__}: {exc}")
            continue
        if text.strip():
            break

    if not text.strip():
        try:
            if convert_pdf_with_chunked_mineru(source_path, output_path):
                return
        except Exception as exc:  # noqa: BLE001
            errors.append(f"{convert_pdf_with_chunked_mineru.__name__}: {exc}")
        if convert_with_mineru(source_path, output_path):
            return
        detail = "; ".join(errors) if errors else "未提取到文本"
        raise ValueError(f"PDF 未提取到可用文本，可能是扫描件，需要 OCR。{detail}")

    output_path.write_text(normalize_markdown(text, title=source_path.stem), encoding="utf-8")


def convert_file(source_path: Path, output_dir: Path) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    output_path = output_markdown_path(source_path, output_dir)
    if variant_exists(source_path, output_path):
        return []

    suffix = source_path.suffix.lower()
    if suffix in {".doc", ".wps", ".wpt"}:
        convert_legacy_word_to_md(source_path, output_path)
    elif suffix == ".hwp":
        convert_hwp_to_md(source_path, output_path)
    elif suffix in {".docx", ".epub"}:
        convert_docx_or_epub_to_md(source_path, output_path)
    elif suffix == ".pdf":
        convert_pdf_to_md(source_path, output_path)
    else:
        raise ValueError(f"暂不支持的格式: {suffix}")

    markdown = output_path.read_text(encoding="utf-8")
    output_path.write_text(normalize_markdown(markdown, title=source_path.stem), encoding="utf-8")
    return [output_path]


def main() -> int:
    parser = argparse.ArgumentParser(
        description="批量将 .doc / .docx / .epub / .pdf / .wps / .wpt / .hwp 转为 Markdown，并跳过已存在的同名 .md。"
    )
    parser.add_argument("input", help="单个文件或目录")
    parser.add_argument("-o", "--output-dir", help="输出目录，默认写回输入文件所在目录")
    args = parser.parse_args()

    input_path = Path(args.input).expanduser().resolve()
    base_output_dir = resolve_base_output_dir(input_path, args.output_dir)
    files = discover_files(input_path)

    if not files:
        print("没有找到可转换的文件（支持 .doc / .docx / .epub / .pdf / .wps / .wpt / .hwp）。", file=sys.stderr)
        return 1

    suffixes = {path.suffix.lower() for path in files}
    if suffixes.intersection({".doc", ".docx", ".epub", ".wps", ".wpt"}) and not shutil.which("pandoc"):
        print("[WARN] 未检测到 pandoc，doc/docx/epub/wps/wpt 可能转换失败。", file=sys.stderr)
    if ".pdf" in suffixes and not shutil.which("pdftotext"):
        print("[INFO] 未检测到 pdftotext，将使用 Python 提取器处理 PDF。", file=sys.stderr)
    if os.name == "nt" and suffixes.intersection({".doc", ".wps", ".wpt"}) and not shutil.which("textutil"):
        print("[INFO] Windows 不支持 textutil；旧版文档建议先转为 .docx。", file=sys.stderr)
    if ".hwp" in suffixes and not get_hwp5txt_runner():
        print("[INFO] 检测到 .hwp：首次转换会自动安装 pyhwp，耗时会稍长。", file=sys.stderr)

    failures: list[tuple[Path, str]] = []
    skipped = 0
    succeeded = 0

    for source_path in files:
        try:
            file_output_dir = output_dir_for_file(source_path, input_path, base_output_dir)
            outputs = convert_file(source_path, file_output_dir)
            if not outputs:
                skipped += 1
                print(f"[SKIP] {source_path}")
                continue
            succeeded += 1
            print(f"[OK] {source_path}")
            for output in outputs:
                print(f"     -> {output}")
        except Exception as exc:  # noqa: BLE001
            failures.append((source_path, str(exc)))
            print(f"[FAIL] {source_path}: {exc}", file=sys.stderr)

    if failures:
        print(
            f"\n完成，成功 {succeeded} 个，跳过 {skipped} 个，失败 {len(failures)} 个。",
            file=sys.stderr,
        )
        return 2

    print(f"\n完成，成功 {succeeded} 个，跳过 {skipped} 个。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
