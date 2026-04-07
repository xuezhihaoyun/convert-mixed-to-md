#!/bin/zsh
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
SCRIPT_PATH="$SCRIPT_DIR/convert_mixed_to_md.py"
REQUIREMENTS_PATH="$SCRIPT_DIR/requirements.txt"
VENV_DIR="$SCRIPT_DIR/.venv"
VENV_PYTHON="$VENV_DIR/bin/python3"
export PYTHONDONTWRITEBYTECODE=1

normalize_path() {
  local value="$1"
  value="${value%\"}"
  value="${value#\"}"
  value="${value%\'}"
  value="${value#\'}"
  value="${value//\\ / }"
  printf '%s' "$value"
}

if [ ! -f "$SCRIPT_PATH" ]; then
  echo "未找到脚本：$SCRIPT_PATH"
  echo "按回车退出。"
  read -r
  exit 1
fi

if ! command -v python3 >/dev/null 2>&1; then
  echo "未找到 python3，请先安装 Python 3。"
  echo "按回车退出。"
  read -r
  exit 1
fi

ensure_runtime() {
  if [ ! -d "$VENV_DIR" ]; then
    echo "[INIT] 首次运行，正在创建本地 Python 环境..."
    if ! python3 -m venv "$VENV_DIR"; then
      echo "[FAIL] 创建虚拟环境失败。"
      return 1
    fi
  fi

  local req_hash
  req_hash="$(shasum -a 256 "$REQUIREMENTS_PATH" | awk '{print $1}')"
  local hash_file="$VENV_DIR/.requirements.sha256"
  local old_hash=""
  if [ -f "$hash_file" ]; then
    old_hash="$(cat "$hash_file" 2>/dev/null)"
  fi

  if [ "$req_hash" != "$old_hash" ]; then
    echo "[INIT] 正在安装/更新 Python 依赖..."
    if ! "$VENV_PYTHON" -m pip install --upgrade pip >/dev/null 2>&1; then
      echo "[WARN] pip 升级失败，继续尝试安装依赖。"
    fi
    if ! "$VENV_PYTHON" -m pip install -r "$REQUIREMENTS_PATH"; then
      echo "[FAIL] 依赖安装失败。你可以手动执行："
      echo "       \"$VENV_PYTHON\" -m pip install -r \"$REQUIREMENTS_PATH\""
      return 1
    fi
    echo "$req_hash" > "$hash_file"
  fi

  if ! "$VENV_PYTHON" -c "import requests" >/dev/null 2>&1; then
    echo "[FAIL] requests 仍不可用，依赖环境异常。"
    return 1
  fi

  return 0
}

if ! ensure_runtime; then
  echo "按回车退出。"
  read -r
  exit 1
fi

if ! command -v pandoc >/dev/null 2>&1; then
  echo "[提示] 未检测到 pandoc，建议安装：brew install pandoc"
fi
if ! command -v pdftotext >/dev/null 2>&1; then
  echo "[提示] 未检测到 pdftotext（poppler），建议安装：brew install poppler"
fi

SUCCESS_COUNT=0
FAIL_COUNT=0
WARNED_TOKEN=0

process_one() {
  local raw_path="$1"
  local target_path
  target_path="$(normalize_path "$raw_path")"
  if [ -z "$target_path" ]; then
    return 0
  fi

  echo
  echo "开始处理：$target_path"
  echo
  if [ -z "$MINERU_TOKEN" ] && [ "$WARNED_TOKEN" -eq 0 ]; then
    echo "[提示] 未设置 MINERU_TOKEN。普通文档可继续转换；扫描版 PDF 可能失败。"
    WARNED_TOKEN=1
  fi
  if "$VENV_PYTHON" "$SCRIPT_PATH" "$target_path"; then
    SUCCESS_COUNT=$((SUCCESS_COUNT + 1))
  else
    FAIL_COUNT=$((FAIL_COUNT + 1))
    echo
    echo "处理失败：$target_path"
  fi
}

echo "convert_mixed_to_md"
echo
echo "支持格式：doc / docx / pdf / epub / wps / wpt / hwp"
echo "可连续处理：每次转完可继续输入路径，直接回车结束。"

if [ "$#" -gt 0 ]; then
  for RAW_PATH in "$@"; do
    process_one "$RAW_PATH"
  done
fi

while true; do
  echo
  echo "请输入文件或目录路径（可拖入，直接回车结束）："
  read -r TARGET_PATH_RAW
  TARGET_PATH_RAW="$(normalize_path "$TARGET_PATH_RAW")"
  if [ -z "$TARGET_PATH_RAW" ]; then
    break
  fi
  process_one "$TARGET_PATH_RAW"
  echo
  echo "当前累计：成功 $SUCCESS_COUNT 个，失败 $FAIL_COUNT 个。"
done

echo
echo "处理结束：成功 $SUCCESS_COUNT 个，失败 $FAIL_COUNT 个。"
echo "按回车退出。"
read -r

if [ "$FAIL_COUNT" -gt 0 ]; then
  exit 1
fi
