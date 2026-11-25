#!/bin/bash
set -eu
cd "$(dirname "$0")"

if [ "$(sw_vers -productVersion | cut -d. -f1)" -lt 14 ]; then
    echo '这个版本需要 macOS 14 或更新版本。' >&2
    exit 1
fi

for candidate in python3.11 python3 python3.10 python3.12 python3.13 python3.14 python3.15; do
    if command -v "$candidate" >/dev/null && "$candidate" -c 'import sys; sys.exit(not ((3, 10) <= sys.version_info[:2] < (3, 16)))' 2>/dev/null; then
        interpreter="$candidate"
        break
    fi
done
if [ -z "${interpreter:-}" ]; then
    echo '请先从 https://www.python.org/downloads/macos/ 安装 Python 3.11，再打开此文件。' >&2
    exit 1
fi

exec "$interpreter" - <<'PY'
import fcntl
import os
from pathlib import Path
import platform
import subprocess
import sys
import venv

project = Path.cwd()
support = Path.home() / "Library" / "Application Support" / "ppt time"
support.mkdir(parents=True, exist_ok=True)
environment = support / f"venv-{platform.machine()}-{sys.version_info.major}.{sys.version_info.minor}"
python = environment / "bin" / "python"
requirements = project / "requirements.txt"
stamp = environment / ".ppt-time-requirements"

with (support / "install.lock").open("a") as lock:
    fcntl.flock(lock, fcntl.LOCK_EX)
    if not python.is_file():
        venv.EnvBuilder(with_pip=True).create(environment)
    current = requirements.read_bytes()
    if not stamp.is_file() or stamp.read_bytes() != current:
        print("正在准备首次运行所需的组件…", flush=True)
        subprocess.run([python, "-m", "pip", "--disable-pip-version-check", "install", "-r", requirements], check=True)
        subprocess.run([python, "-c", "import segno; from PySide6 import QtMultimedia, QtMultimediaWidgets, QtPdf, QtSvg, QtWidgets"], check=True)
        stamp.write_bytes(current)

os.execv(python, [str(python), str(project / "main.py")])
PY
