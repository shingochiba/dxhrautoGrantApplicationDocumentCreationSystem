"""BTS (Bug Tracking System / バグ・タスク管理) パネル。

Google Sheets API + サービスアカウント認証で、非公開スプレッドシートを
読み込んで Streamlit 上に表示する。

認証情報の取得順:
  1. st.secrets["gcp_service_account"] (Streamlit Cloud デプロイ向け推奨)
  2. 環境変数 GOOGLE_APPLICATION_CREDENTIALS が指すJSONファイル
  3. プロジェクト直下の config/google_service_account.json
"""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import List, Optional, Tuple

import pandas as pd
import streamlit as st


# デフォルトの BTS スプレッドシート
BTS_SHEET_ID = "14_fUv5tcMidrtub6LmRKy8-YyQsiYTJha8McZ1Lasyo"
BTS_GID = 0  # 1枚目のシート
BTS_VIEW_URL = (
    f"https://docs.google.com/spreadsheets/d/{BTS_SHEET_ID}/edit?gid={BTS_GID}#gid={BTS_GID}"
)

# サービスアカウントJSONのローカル既定パス
_BASE_DIR = Path(__file__).resolve().parent.parent
_DEFAULT_SA_JSON = _BASE_DIR / "config" / "google_service_account.json"

# Google Sheets API のスコープ (読み取り専用)
_SCOPES = ["https://www.googleapis.com/auth/spreadsheets.readonly"]

# 対応状況の表示用設定: 表示順, アイコン
STATUS_ORDER = ["オープン", "テスト中", "FB待ち", "クローズ", "未設定", ""]
STATUS_ICON = {
    "オープン": "🟥",
    "テスト中": "🟧",
    "FB待ち": "🟨",
    "クローズ": "🟩",
}


class BTSCredentialsError(RuntimeError):
    """サービスアカウント認証情報が用意されていないことを示す例外。"""


def _load_credentials_info() -> Tuple[dict, str]:
    """サービスアカウントの credentials dict を返し、出所のラベルも返す。

    Returns:
        (credentials_dict, source_label)

    Raises:
        BTSCredentialsError: どこにも認証情報が見つからない場合。
    """
    # 1. st.secrets
    try:
        if "gcp_service_account" in st.secrets:
            info = dict(st.secrets["gcp_service_account"])
            if info.get("client_email") and info.get("private_key"):
                return info, "st.secrets[gcp_service_account]"
    except Exception:
        # st.secrets がそもそも未設定の環境では FileNotFoundError 等が出る
        pass

    # 2. 環境変数 GOOGLE_APPLICATION_CREDENTIALS
    env_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if env_path and Path(env_path).is_file():
        with open(env_path, "r", encoding="utf-8") as f:
            info = json.load(f)
        return info, f"環境変数 GOOGLE_APPLICATION_CREDENTIALS ({env_path})"

    # 3. ローカル既定パス
    if _DEFAULT_SA_JSON.is_file():
        with open(_DEFAULT_SA_JSON, "r", encoding="utf-8") as f:
            info = json.load(f)
        return info, str(_DEFAULT_SA_JSON)

    raise BTSCredentialsError(
        "サービスアカウントの認証情報が見つかりません。"
    )


@st.cache_resource(show_spinner=False)
def _get_gspread_client():
    """gspread クライアントを生成 (キャッシュ)。"""
    import gspread
    from google.oauth2.service_account import Credentials

    info, _source = _load_credentials_info()
    creds = Credentials.from_service_account_info(info, scopes=_SCOPES)
    return gspread.authorize(creds)


def _get_service_account_email() -> Optional[str]:
    """設定されたサービスアカウントのメールアドレスを返す (未設定なら None)。"""
    try:
        info, _ = _load_credentials_info()
    except BTSCredentialsError:
        return None
    return info.get("client_email")


@st.cache_data(ttl=120, show_spinner=False)
def _load_bts_dataframe(sheet_id: str, gid: int) -> pd.DataFrame:
    """Google Sheets API 経由でシートを取得し DataFrame を返す。

    キャッシュ TTL は 120 秒。手動更新ボタンで `st.cache_data.clear()` を呼べる。
    """
    client = _get_gspread_client()
    spreadsheet = client.open_by_key(sheet_id)
    worksheet = None
    for ws in spreadsheet.worksheets():
        if ws.id == gid:
            worksheet = ws
            break
    if worksheet is None:
        # gid が見つからない場合は最初のシートを使う
        worksheet = spreadsheet.get_worksheet(0)

    # 全セル値を二次元配列で取得 (空欄を None ではなく空文字で扱う)
    values = worksheet.get_all_values()
    if not values:
        return pd.DataFrame()
    header = [str(h).strip() for h in values[0]]
    df = pd.DataFrame(values[1:], columns=header)
    return df


def fetch_bts_items() -> pd.DataFrame:
    """BTS 項目を取得して、空行除去・整形済みの DataFrame を返す。"""
    df = _load_bts_dataframe(BTS_SHEET_ID, BTS_GID)
    if df.empty:
        return df
    # 「内容」列があれば、内容が完全に空の行は除外
    if "内容" in df.columns:
        df = df[df["内容"].astype(str).str.strip() != ""].copy()
    # 「対応状況」列があれば、ソート用キーを付与
    if "対応状況" in df.columns:
        df["_status_rank"] = df["対応状況"].map(
            lambda s: STATUS_ORDER.index(s) if s in STATUS_ORDER else len(STATUS_ORDER)
        )
    return df


def _status_badge(status: str) -> str:
    icon = STATUS_ICON.get(status, "⬜")
    return f"{icon} {status}" if status else "⬜ (未設定)"


def _render_credentials_help(extra_error: str = "") -> None:
    """認証情報が無いときの案内表示。"""
    st.error("⚠️ サービスアカウントの認証情報が設定されていません。")
    if extra_error:
        st.caption(f"詳細: {extra_error}")

    with st.expander("🔧 セットアップ手順", expanded=True):
        st.markdown(
            "**1. Google Cloud でサービスアカウントを作成**\n"
            "1. [Google Cloud Console](https://console.cloud.google.com/) にログイン\n"
            "2. 新規プロジェクト作成 (既存でもOK)\n"
            "3. **APIとサービス → ライブラリ** で「Google Sheets API」を検索→有効化\n"
            "4. **APIとサービス → 認証情報** → 「認証情報を作成」→「サービスアカウント」\n"
            "5. 任意の名前 (例: `bts-reader`) で作成→「キーを管理」→「鍵を追加」→「JSON」を選択してダウンロード\n"
            "\n"
            "**2. スプレッドシートを共有**\n"
            f"- [BTS スプレッドシート]({BTS_VIEW_URL}) を開く\n"
            "- 「共有」ボタン→ サービスアカウントの **client_email** (`xxxxx@xxxxx.iam.gserviceaccount.com`) を **閲覧者** として追加\n"
            "\n"
            "**3. 認証情報を配置 (いずれか1つ)**\n"
            "- **A. Streamlit Cloud (本番運用推奨)**: アプリ設定の `Secrets` 欄に下記をペースト\n"
            "```toml\n"
            "[gcp_service_account]\n"
            'type = "service_account"\n'
            'project_id = "xxx"\n'
            'private_key_id = "xxx"\n'
            'private_key = "-----BEGIN PRIVATE KEY-----\\nxxx\\n-----END PRIVATE KEY-----\\n"\n'
            'client_email = "xxx@xxx.iam.gserviceaccount.com"\n'
            'client_id = "xxx"\n'
            'auth_uri = "https://accounts.google.com/o/oauth2/auth"\n'
            'token_uri = "https://oauth2.googleapis.com/token"\n'
            'auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"\n'
            'client_x509_cert_url = "xxx"\n'
            "```\n"
            "- **B. ローカル開発**: ダウンロードしたJSONを `config/google_service_account.json` に置く (gitignore済)\n"
            "- **C. 環境変数**: `GOOGLE_APPLICATION_CREDENTIALS` にJSONファイルのフルパスを設定"
        )


def render_bts_panel() -> None:
    """BTS パネルを Streamlit に描画する。"""
    st.header("🐛 BTS (バグ・タスク管理)")
    st.caption(
        "Google スプレッドシートと連携して、バグ・機能要望の対応状況を確認できます。"
    )

    col_link, col_refresh = st.columns([4, 1])
    with col_link:
        st.markdown(
            f"📊 [スプレッドシートを開く]({BTS_VIEW_URL})  ｜ "
            "編集はスプレッドシート側で行ってください"
        )
    with col_refresh:
        if st.button("🔄 再取得", use_container_width=True, key="bts_refresh"):
            _load_bts_dataframe.clear()
            st.rerun()

    # サービスアカウントメールの表示 (共有用)
    sa_email = _get_service_account_email()
    if sa_email:
        st.caption(f"🔐 サービスアカウント: `{sa_email}`")

    # 取得
    try:
        df = fetch_bts_items()
    except BTSCredentialsError as e:
        _render_credentials_help(str(e))
        return
    except Exception as e:
        msg = str(e)
        if "PERMISSION_DENIED" in msg or "403" in msg:
            st.error(
                "❌ サービスアカウントにスプレッドシートへのアクセス権がありません。"
            )
            st.markdown(
                f"スプレッドシートの「共有」設定で、サービスアカウントのメール "
                f"(`{sa_email or '不明'}`) を **閲覧者** として追加してください。"
            )
            st.markdown(f"🔗 [スプレッドシートを開く]({BTS_VIEW_URL})")
        elif "APIError" in type(e).__name__ or "404" in msg:
            st.error("❌ スプレッドシートが見つかりません。シートIDが正しいか確認してください。")
        else:
            st.error("スプレッドシートを取得できませんでした。")
        with st.expander("エラー詳細", expanded=False):
            st.code(f"{type(e).__name__}: {e}")
        return

    if df.empty:
        st.info("BTS に項目はありません。")
        return

    # ---- サマリ (対応状況別件数) ----
    if "対応状況" in df.columns:
        status_counts = df["対応状況"].value_counts(dropna=False)
        ordered = [s for s in STATUS_ORDER if s in status_counts.index]
        ordered += [s for s in status_counts.index if s not in STATUS_ORDER]
        if ordered:
            cols = st.columns(max(4, len(ordered)))
            for i, status in enumerate(ordered):
                label = _status_badge(status) if status else "⬜ (未設定)"
                cols[i % len(cols)].metric(label, int(status_counts[status]))

    st.markdown("---")

    # ---- フィルタ ----
    filter_col1, filter_col2 = st.columns([2, 3])
    selected_status: List[str] = []
    if "対応状況" in df.columns:
        with filter_col1:
            all_statuses = sorted(
                [s for s in df["対応状況"].dropna().unique().tolist() if s],
                key=lambda s: STATUS_ORDER.index(s) if s in STATUS_ORDER else 99,
            )
            default = [s for s in all_statuses if s != "クローズ"]
            selected_status = st.multiselect(
                "対応状況で絞り込み",
                options=all_statuses,
                default=default,
                key="bts_filter_status",
            )
    keyword = ""
    with filter_col2:
        keyword = st.text_input(
            "キーワード検索 (内容・修正方針・修正内容)", value="", key="bts_filter_kw"
        )

    filtered = df.copy()
    if selected_status and "対応状況" in filtered.columns:
        filtered = filtered[filtered["対応状況"].isin(selected_status)]
    if keyword:
        kw = keyword.strip()
        search_cols = [c for c in ("内容", "修正方針", "修正内容", "起票者") if c in filtered.columns]
        if search_cols:
            mask = filtered[search_cols].apply(
                lambda col: col.astype(str).str.contains(kw, na=False), axis=0
            ).any(axis=1)
            filtered = filtered[mask]

    if "_status_rank" in filtered.columns:
        filtered = filtered.sort_values(["_status_rank", "項番"], na_position="last")

    # ---- 表示用に整形 ----
    display_cols = [c for c in [
        "項番", "処理区分", "内容", "起票者", "区分",
        "修正方針", "バグ", "対応状況",
        "対応完了の場合は、完了日", "ver", "修正内容",
    ] if c in filtered.columns]
    show_df = filtered[display_cols].reset_index(drop=True)

    st.caption(f"表示中: {len(show_df)} 件 / 全 {len(df)} 件")
    st.dataframe(
        show_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "項番": st.column_config.TextColumn("項番", width="small"),
            "処理区分": st.column_config.TextColumn("区分", width="small"),
            "内容": st.column_config.TextColumn("内容", width="large"),
            "起票者": st.column_config.TextColumn("起票", width="small"),
            "対応状況": st.column_config.TextColumn("状況", width="small"),
            "ver": st.column_config.TextColumn("ver", width="small"),
        },
    )

    # ---- 個別カードで詳細表示 (オープン項目のみ) ----
    if "対応状況" in filtered.columns:
        open_items = filtered[filtered["対応状況"] == "オープン"]
        if not open_items.empty:
            st.markdown("---")
            st.subheader(f"🟥 オープン項目の詳細 ({len(open_items)} 件)")
            for _, row in open_items.iterrows():
                title = f"#{row.get('項番', '')} [{row.get('処理区分', '')}] {row.get('内容', '')}"
                with st.expander(title, expanded=False):
                    for key in ("起票者", "区分", "修正方針", "バグ", "修正内容"):
                        if key in row and str(row[key]).strip():
                            st.markdown(f"**{key}**: {row[key]}")
