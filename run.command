#!/bin/zsh
set -eu
setopt null_glob
tool_dir="${0:A:h}"
for bundled_cli in "$tool_dir/../bin/precise-pptx/precise-pptx" "$tool_dir/PPTX Precise.app/Contents/Resources/bin/precise-pptx/precise-pptx"; do
  if [[ -x "$bundled_cli" ]]; then
    exec "$bundled_cli" "$@"
  fi
done
python_candidates=(
  "$tool_dir/.venv/bin/python3"
  "$HOME/Library/Application Support/PPTX-Precise/venv/bin/python3"
  /opt/homebrew/bin/python3 /usr/local/bin/python3 /usr/bin/python3
)
tool_python=""
for candidate in "${python_candidates[@]}"; do
  if [[ -x "$candidate" ]] && "$candidate" -c 'import sys; assert sys.version_info >= (3,10); import pypdf, PIL; from pypdf import PdfWriter; assert hasattr(PdfWriter, "clone_document_from_reader")' >/dev/null 2>&1; then
    tool_python="$candidate"
    break
  fi
done
if [[ -z "$tool_python" ]]; then
  print -u2 '缺少 Python 依赖。请先双击「安装依赖.command」，再运行转换工具。'
  exit 1
fi
exec "$tool_python" "$tool_dir/precise_pptx.py" "$@"
