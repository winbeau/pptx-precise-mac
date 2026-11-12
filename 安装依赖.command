#!/bin/zsh
set -eu
tool_dir="${0:A:h}"
tool_python="$(command -v python3 || true)"
if [[ -z "$tool_python" ]]; then
  print '请从 python.org 安装 Python 3.10 或更新版本。'
  exit 1
fi
env_dir="$HOME/Library/Application Support/PPTX-Precise/venv"
"$tool_python" -m venv "$env_dir"
"$env_dir/bin/python3" -m pip install -r "$tool_dir/requirements.txt"
print '依赖安装完成。现在可以双击 PPTX Precise.app。'
