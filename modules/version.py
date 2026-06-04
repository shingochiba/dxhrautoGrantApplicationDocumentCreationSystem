"""アプリのバージョン情報 (最新Git commit メッセージ) を取得するモジュール。

Streamlit Cloud デプロイ環境でも `git` コマンドは利用可能 (ベースイメージに含まれる)。
取得に失敗した場合は空タプルを返す。
"""
from __future__ import annotations

import subprocess
from pathlib import Path
from typing import Optional, Tuple

_BASE_DIR = Path(__file__).resolve().parent.parent


def get_latest_commit_info() -> Tuple[Optional[str], Optional[str], Optional[str]]:
    """最新コミットの (subject, short_hash, iso_date) を返す。
    取得失敗時は (None, None, None)。
    """
    try:
        result = subprocess.run(
            ["git", "log", "-1", "--pretty=format:%s%n%h%n%ai"],
            cwd=str(_BASE_DIR),
            capture_output=True,
            text=True,
            timeout=3,
            check=False,
        )
        if result.returncode != 0:
            return (None, None, None)
        lines = result.stdout.strip().split("\n")
        if len(lines) >= 3:
            return (lines[0].strip(), lines[1].strip(), lines[2].strip())
        return (None, None, None)
    except (subprocess.TimeoutExpired, FileNotFoundError, OSError):
        return (None, None, None)
