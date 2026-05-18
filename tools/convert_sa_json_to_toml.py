"""Google サービスアカウント JSON を Streamlit Secrets 用 TOML に変換するヘルパー。

使い方:
    python tools/convert_sa_json_to_toml.py <path/to/service_account.json>

出力をそのままコピペで Streamlit Cloud の Secrets 欄に貼り付ければよい。
"""
from __future__ import annotations

import json
import sys
from pathlib import Path


def convert(json_path: Path) -> str:
    with open(json_path, "r", encoding="utf-8") as f:
        info = json.load(f)

    lines = ["[gcp_service_account]"]
    # private_key は三連クォートで囲み、\n を実改行に展開
    for key in [
        "type", "project_id", "private_key_id", "private_key",
        "client_email", "client_id",
        "auth_uri", "token_uri",
        "auth_provider_x509_cert_url", "client_x509_cert_url",
        "universe_domain",
    ]:
        if key not in info:
            continue
        value = info[key]
        if key == "private_key":
            # 三連クォートで囲んで実改行で展開
            unescaped = value.replace("\\n", "\n")
            lines.append(f'{key} = """{unescaped}"""')
        else:
            # 通常の文字列。ダブルクォート内のエスケープを最小限に
            escaped = str(value).replace("\\", "\\\\").replace('"', '\\"')
            lines.append(f'{key} = "{escaped}"')
    return "\n".join(lines) + "\n"


def main() -> int:
    if len(sys.argv) != 2:
        print(__doc__, file=sys.stderr)
        return 1
    path = Path(sys.argv[1])
    if not path.is_file():
        print(f"ファイルが見つかりません: {path}", file=sys.stderr)
        return 1
    toml = convert(path)
    print(toml)
    return 0


if __name__ == "__main__":
    sys.exit(main())
