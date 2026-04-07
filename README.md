# convert_mixed_to_md

批量把 `.doc`、`.docx`、`.epub`、`.pdf`、`.wps`、`.wpt`、`.hwp` 转成 Markdown。

> **必看：扫描版 PDF 需要 `MINERU_TOKEN`**
>
> - 普通 PDF（有文字层）通常不需要 token。  
> - 扫描版 / 纯图片 PDF 没有 `MINERU_TOKEN` 时，转换很可能失败。  
> - 这不是脚本 bug，是 OCR 服务需要鉴权。

## 支持格式

- `.doc`
- `.docx`
- `.epub`
- `.pdf`
- `.wps`
- `.wpt`
- `.hwp`

## 功能

- 自动识别文件类型
- 默认写回原目录
- 已有同名 `.md` 自动跳过
- 旧版 `.doc` 自动尝试两种解析方式
- 旧版 `.wps/.wpt` 按旧式文档方式尝试解析
- `.hwp` 自动调用 `hwp5txt`（若缺失会首次自动安装 `pyhwp`）
- 扫描版 `pdf` 可选使用 MinerU OCR
- 大扫描版 `pdf` 会自动分段 OCR 再合并

## 重要提醒（MinerU token）

- 普通 PDF（有文字层）通常不需要 token。
- 扫描版 / 纯图片 PDF 基本需要 OCR，建议配置 `MINERU_TOKEN`。
- 未配置 token 时，这类 PDF 可能直接失败（这是预期行为，不是脚本损坏）。

## 安装

### 1. 安装 Python

需要 Python 3.9+。

### 2. 安装系统命令

macOS:

```bash
brew install pandoc poppler
```

Windows（推荐任选一种）:

PowerShell + winget:

```powershell
winget install --id JohnMacFarlane.Pandoc -e
winget install --id oschwartz10612.Poppler -e
```

或 Chocolatey:

```powershell
choco install pandoc poppler -y
```

安装后请确认命令可用：

- `pandoc`
- `pdftotext`

### 3. Python 依赖（已自动化）

`run.command` 首次运行会自动创建本地 `.venv` 并安装 `requirements.txt` 依赖。  
一般不需要手动执行 `pip install`。

## 用法

### 最简单

直接把文件或文件夹拖到 `run.command` 上（推荐）。

也可以双击 `run.command`，然后手动输入路径。

支持格式：`doc / docx / pdf / epub / wps / wpt / hwp`

### Windows 一键版

Windows 用户请使用 `run_windows.bat`。

- 可双击运行
- 可拖拽文件/文件夹到 `.bat`
- 首次自动创建 `.venv` 并安装依赖
- 支持连续输入路径批量处理
- 支持粘贴多路径输入（含 `"C:\\a.docx" "D:\\b.pdf"` 和连写 `C:\a.docxC:\b.pdf`）

如果是旧版 `.doc/.wps/.wpt`，Windows 下建议先转为 `.docx` 再转换（成功率更高）。

### 命令行

#### 转单个文件

macOS / Linux:

```bash
python3 convert_mixed_to_md.py '/path/to/file.epub'
```

Windows:

```powershell
python .\convert_mixed_to_md.py "C:\path\to\file.epub"
```

### 转整个目录

macOS / Linux:

```bash
python3 convert_mixed_to_md.py '/path/to/folder'
```

Windows:

```powershell
python .\convert_mixed_to_md.py "C:\path\to\folder"
```

### 输出到另一个目录

macOS / Linux:

```bash
python3 convert_mixed_to_md.py '/path/to/folder' -o '/path/to/output'
```

Windows:

```powershell
python .\convert_mixed_to_md.py "C:\path\to\folder" -o "C:\path\to\output"
```

## 扫描版 PDF

如果 PDF 是扫描版或纯图片版，建议配置 MinerU token：

macOS / Linux:

```bash
export MINERU_TOKEN='your_token'
python3 convert_mixed_to_md.py '/path/to/folder'
```

Windows PowerShell:

```powershell
$env:MINERU_TOKEN='your_token'
python .\convert_mixed_to_md.py "C:\path\to\folder"
```

Windows CMD:

```bat
set MINERU_TOKEN=your_token
python .\convert_mixed_to_md.py "C:\path\to\folder"
```

如果没有配置 token，普通有文字层的 PDF 仍然可以转换。

大扫描版 PDF 会自动分段 OCR，并显示类似下面的进度：

```text
[INFO] 大 PDF 触发分段 OCR: ... (242 页, 10 段)
[INFO] OCR 第 1/10 段: ...
```

## 说明

- 同目录下如果同时有 `a.pdf` 和 `a.epub`，输出会分别命名为：
  - `a.pdf.md`
  - `a.epub.md`
- `epub` 转换时如果有图片，通常会生成一个 `_assets` 目录。
