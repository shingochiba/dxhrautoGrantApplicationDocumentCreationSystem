"""BTS (Bug Tracking System / バグ・タスク管理) パネル。

Google Sheets API + サービスアカウント認証で、非公開スプレッドシートを
読み込み・編集して Streamlit 上に表示する。

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

# Google Sheets API スコープ (編集も行うため read/write)
_SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

# ヘッダー検出に使うキーワード (これらのうち2つ以上含まれる行をヘッダー行とみなす)
_HEADER_MARKERS = ("項番", "内容", "対応状況", "処理区分", "起票者")
_HEADER_MIN_MATCHES = 2

# 期待される列の順序 (表示・編集用)
_EXPECTED_COLS = [
    "項番", "処理区分", "内容", "起票者", "区分",
    "修正方針", "バグ", "対応状況",
    "対応完了の場合は、完了日", "ver", "修正内容",
]

# 対応状況の表示用設定: 表示順, アイコン
STATUS_ORDER = ["オープン", "テスト中", "FB待ち", "クローズ", "未設定", ""]
STATUS_ICON = {
    "オープン": "🟥",
    "テスト中": "🟧",
    "FB待ち": "🟨",
    "クローズ": "🟩",
}
STATUS_OPTIONS = ["", "オープン", "テスト中", "FB待ち", "クローズ"]


class BTSCredentialsError(RuntimeError):
    """サービスアカウント認証情報が用意されていないことを示す例外。"""


def _load_credentials_info() -> Tuple[dict, str]:
    """サービスアカウントの credentials dict を返し、出所のラベルも返す。"""
    try:
        if "gcp_service_account" in st.secrets:
            info = dict(st.secrets["gcp_service_account"])
            if info.get("client_email") and info.get("private_key"):
                return info, "st.secrets[gcp_service_account]"
    except Exception:
        pass

    env_path = os.environ.get("GOOGLE_APPLICATION_CREDENTIALS")
    if env_path and Path(env_path).is_file():
        with open(env_path, "r", encoding="utf-8") as f:
            info = json.load(f)
        return info, f"環境変数 GOOGLE_APPLICATION_CREDENTIALS ({env_path})"

    if _DEFAULT_SA_JSON.is_file():
        with open(_DEFAULT_SA_JSON, "r", encoding="utf-8") as f:
            info = json.load(f)
        return info, str(_DEFAULT_SA_JSON)

    raise BTSCredentialsError("サービスアカウントの認証情報が見つかりません。")


@st.cache_resource(show_spinner=False)
def _get_gspread_client():
    """gspread クライアントを生成 (キャッシュ)。"""
    import gspread
    from google.oauth2.service_account import Credentials

    info, _source = _load_credentials_info()
    creds = Credentials.from_service_account_info(info, scopes=_SCOPES)
    return gspread.authorize(creds)


def _get_service_account_email() -> Optional[str]:
    try:
        info, _ = _load_credentials_info()
    except BTSCredentialsError:
        return None
    return info.get("client_email")


def _get_worksheet():
    """編集対象の worksheet を返す。"""
    client = _get_gspread_client()
    spreadsheet = client.open_by_key(BTS_SHEET_ID)
    for ws in spreadsheet.worksheets():
        if ws.id == BTS_GID:
            return ws
    return spreadsheet.get_worksheet(0)


def _find_header_row(values: List[List[str]]) -> int:
    """_HEADER_MARKERS のうち _HEADER_MIN_MATCHES 個以上含む最初の行のインデックスを返す。
    見つからなければ 0 (1行目をヘッダーとして扱う)。
    """
    for i, row in enumerate(values):
        row_set = {str(c).strip() for c in row}
        hits = sum(1 for m in _HEADER_MARKERS if m in row_set)
        if hits >= _HEADER_MIN_MATCHES:
            return i
    return 0


@st.cache_data(ttl=120, show_spinner=False)
def _load_bts_dataframe() -> Tuple[pd.DataFrame, int]:
    """シート全体を取得し、(DataFrame, header_row_index_1based) を返す。

    header_row_index は gspread でセル更新する際の行番号 (1-indexed) として使う。
    """
    ws = _get_worksheet()
    values = ws.get_all_values()
    if not values:
        return pd.DataFrame(), 1

    header_idx = _find_header_row(values)
    header = [str(c).strip() for c in values[header_idx]]
    data_rows = values[header_idx + 1:]

    # 各行を header の長さに揃える
    normalized = []
    for r in data_rows:
        if len(r) < len(header):
            r = r + [""] * (len(header) - len(r))
        elif len(r) > len(header):
            r = r[: len(header)]
        normalized.append(r)

    df = pd.DataFrame(normalized, columns=header)
    # 元の行番号 (1-indexed) を保持
    df["_row"] = [header_idx + 2 + i for i in range(len(df))]

    # 内容が空、または「項番」セルが見出し文字 (二度目のヘッダー) の行を除外
    if "内容" in df.columns:
        is_empty = df["内容"].astype(str).str.strip() == ""
        is_dup_header = False
        if "項番" in df.columns:
            is_dup_header = df["項番"].astype(str).str.strip() == "項番"
        df = df[~(is_empty | is_dup_header)].reset_index(drop=True)

    return df, header_idx + 1  # 1-indexed header row


def fetch_bts_items() -> Tuple[pd.DataFrame, int]:
    """BTS 項目を取得して、(DataFrame, header_row_index) を返す。"""
    df, header_row = _load_bts_dataframe()
    if "対応状況" in df.columns:
        df["_status_rank"] = df["対応状況"].map(
            lambda s: STATUS_ORDER.index(s) if s in STATUS_ORDER else len(STATUS_ORDER)
        )
    return df, header_row


def _status_badge(status: str) -> str:
    icon = STATUS_ICON.get(status, "⬜")
    return f"{icon} {status}" if status else "⬜ (未設定)"


def _col_letter(n: int) -> str:
    """1-indexed の列番号を Excel 列文字 (A, B, ..., AA) に変換。"""
    out = ""
    while n > 0:
        n, rem = divmod(n - 1, 26)
        out = chr(ord("A") + rem) + out
    return out


def _apply_changes(
    edited_df: pd.DataFrame,
    original_df: pd.DataFrame,
    columns: List[str],
) -> int:
    """edited_df と original_df を比較し、変更されたセルをシートに書き戻す。
    変更されたセル数を返す。
    """
    ws = _get_worksheet()
    # 列名 -> 列番号 (1-indexed) のマップ
    # 元の header に対応する列番号を取得するため、もう一度シートのヘッダーを読む
    sheet_values = ws.get_all_values()
    header_idx = _find_header_row(sheet_values)
    header = [str(c).strip() for c in sheet_values[header_idx]]
    col_to_num = {name: i + 1 for i, name in enumerate(header)}

    # gspread の batch_update 用にセル更新を集める
    updates = []
    for i in range(len(edited_df)):
        row_num = int(edited_df.iloc[i]["_row"])
        for col in columns:
            if col not in col_to_num:
                continue
            new_val = str(edited_df.iloc[i].get(col, "") or "")
            old_val = str(original_df.iloc[i].get(col, "") or "")
            if new_val != old_val:
                cell = f"{_col_letter(col_to_num[col])}{row_num}"
                updates.append({"range": cell, "values": [[new_val]]})

    if updates:
        ws.batch_update(updates, value_input_option="USER_ENTERED")
    return len(updates)


def _append_bts_item(item: dict) -> int:
    """新規 BTS 項目をスプレッドシートに追記する。

    動作:
      - ヘッダー行を検出
      - データブロック内で「内容」が空の最初の行を探して、そこに上書き
      - 既存の placeholder 行に 項番 があればそれを採用、なければ自動採番
      - 空き行が無ければ、シート末尾に append_row で追加

    Returns:
        書き込んだ行の 項番 (int)。失敗時は例外。
    """
    ws = _get_worksheet()
    sheet_values = ws.get_all_values()
    if not sheet_values:
        raise RuntimeError("シートが空です。ヘッダー行が存在しません。")

    header_idx = _find_header_row(sheet_values)
    header = [str(c).strip() for c in sheet_values[header_idx]]
    col_to_num = {name: i + 1 for i, name in enumerate(header)}

    content_col = col_to_num.get("内容")
    han_col = col_to_num.get("項番")
    if not content_col:
        raise RuntimeError("シートに「内容」列が見つかりません。")

    # 1) データブロック内で「内容」が空の最初の行を探す
    target_row = None
    block_end = len(sheet_values)
    for i in range(header_idx + 1, len(sheet_values)):
        row = sheet_values[i]
        # padding
        if len(row) < len(header):
            row = row + [""] * (len(header) - len(row))
        # 二度目のヘッダー行に到達したらブロック終わり
        if han_col and str(row[han_col - 1]).strip() == "項番":
            block_end = i
            break
        if str(row[content_col - 1]).strip() == "":
            target_row = i + 1  # 1-indexed
            break

    # 2) 自動採番: 既存項目の最大 項番 + 1
    next_han = 1
    if han_col:
        used_nums = []
        for i in range(header_idx + 1, block_end):
            row = sheet_values[i]
            if len(row) < len(header):
                continue
            content_val = str(row[content_col - 1]).strip()
            han_val = str(row[han_col - 1]).strip()
            if content_val and han_val.isdigit():
                used_nums.append(int(han_val))
        if used_nums:
            next_han = max(used_nums) + 1

    # 3) 書き込み行の決定: target_row が見つからなければ末尾に append
    if target_row is None:
        # シート末尾に新規行を追加
        new_row_values = [""] * len(header)
        for col_name, val in item.items():
            if col_name in col_to_num:
                new_row_values[col_to_num[col_name] - 1] = str(val or "")
        if han_col:
            new_row_values[han_col - 1] = str(next_han)
        ws.append_row(new_row_values, value_input_option="USER_ENTERED")
        return next_han

    # 4) 既存の placeholder 行に書き込み (項番は placeholder の値を尊重)
    if han_col:
        placeholder_han = str(sheet_values[target_row - 1][han_col - 1]).strip()
        if placeholder_han.isdigit():
            next_han = int(placeholder_han)

    updates = []
    for col_name, val in item.items():
        if col_name not in col_to_num:
            continue
        col_num = col_to_num[col_name]
        cell = f"{_col_letter(col_num)}{target_row}"
        updates.append({"range": cell, "values": [[str(val or "")]]})
    # 項番セルも書き込む (placeholder と同じ値 or 新規採番)
    if han_col:
        cell = f"{_col_letter(han_col)}{target_row}"
        updates.append({"range": cell, "values": [[str(next_han)]]})

    if updates:
        ws.batch_update(updates, value_input_option="USER_ENTERED")
    return next_han


def _render_new_entry_form(default_author: str = "") -> None:
    """新規項目登録フォームを描画する。送信時にシートに書き込む。"""
    with st.form("bts_new_entry_form", clear_on_submit=True):
        st.markdown("#### ➕ 新規項目を登録")
        col1, col2 = st.columns([1, 1])
        with col1:
            kubun = st.selectbox(
                "処理区分",
                options=["", "バックエンド", "フロントエンド", "書式", "UI", "その他"],
                key="bts_new_kubun",
            )
            author = st.text_input("起票者", value=default_author, key="bts_new_author")
            category = st.selectbox(
                "区分",
                options=["", "バグ", "設定ミス", "マッピングミス", "UIの問題", "未設定", "その他"],
                key="bts_new_category",
            )
        with col2:
            status = st.selectbox(
                "対応状況",
                options=STATUS_OPTIONS,
                index=STATUS_OPTIONS.index("オープン") if "オープン" in STATUS_OPTIONS else 0,
                key="bts_new_status",
            )
            ver = st.text_input("ver (任意)", value="", key="bts_new_ver")
            bug_note = st.text_input("バグ (任意)", value="", key="bts_new_bug")

        content = st.text_area(
            "内容 *",
            placeholder="例: 様式X-X で 〇〇 が反映されない",
            key="bts_new_content",
        )
        policy = st.text_area(
            "修正方針 (任意)",
            placeholder="どう直すか / 確認事項など",
            key="bts_new_policy",
        )
        fix_detail = st.text_area(
            "修正内容 (任意)",
            placeholder="具体的にどう修正したか",
            key="bts_new_fix",
        )

        submitted = st.form_submit_button(
            "📝 登録", type="primary", width='stretch'
        )

    if submitted:
        if not content.strip():
            st.error("「内容」は必須です。")
            return
        item = {
            "処理区分": kubun,
            "内容": content.strip(),
            "起票者": author.strip(),
            "区分": category,
            "修正方針": policy.strip(),
            "バグ": bug_note.strip(),
            "対応状況": status,
            "ver": ver.strip(),
            "修正内容": fix_detail.strip(),
        }
        try:
            han = _append_bts_item(item)
            st.success(f"✅ 項番 {han} として登録しました。")
            _load_bts_dataframe.clear()
            st.rerun()
        except Exception as e:
            st.error(f"登録に失敗しました: {type(e).__name__}: {e}")


def _render_credentials_help(extra_error: str = "") -> None:
    st.error("⚠️ サービスアカウントの認証情報が設定されていません。")
    if extra_error:
        st.caption(f"詳細: {extra_error}")
    with st.expander("🔧 セットアップ手順", expanded=True):
        st.markdown(
            "DEPLOY.md の「BTS (バグ・タスク管理) 連携の設定」セクションを参照してください。\n\n"
            "ローカル開発の場合: ダウンロードした JSON を "
            "`config/google_service_account.json` に配置してください。"
        )


def render_bts_panel() -> None:
    """BTS パネルを Streamlit に描画する。"""
    st.header("🐛 BTS (バグ・タスク管理)")
    st.caption(
        "Google スプレッドシートと連携して、バグ・機能要望の対応状況を確認・編集できます。"
    )

    col_link, col_refresh = st.columns([4, 1])
    with col_link:
        st.markdown(
            f"📊 [スプレッドシートを開く]({BTS_VIEW_URL})  ｜ "
            "編集はアプリ内のテーブルでも、スプレッドシート側でも可能です"
        )
    with col_refresh:
        if st.button("🔄 再取得", width='stretch', key="bts_refresh"):
            _load_bts_dataframe.clear()
            st.rerun()

    sa_email = _get_service_account_email()
    if sa_email:
        st.caption(f"🔐 サービスアカウント: `{sa_email}`")

    # 取得
    try:
        df, _header_row = fetch_bts_items()
    except BTSCredentialsError as e:
        _render_credentials_help(str(e))
        return
    except Exception as e:
        msg = str(e)
        if "PERMISSION_DENIED" in msg or "403" in msg:
            st.error("❌ サービスアカウントにスプレッドシートへのアクセス権がありません。")
            st.markdown(
                f"スプレッドシートの「共有」設定で、サービスアカウントのメール "
                f"(`{sa_email or '不明'}`) を **編集者** として追加してください。"
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
        with st.expander("🔍 デバッグ: 取得状況", expanded=True):
            try:
                ws = _get_worksheet()
                raw = ws.get_all_values()
                st.write(f"シート総行数: {len(raw)}")
                if raw:
                    st.write("先頭5行 (生データ):")
                    st.dataframe(pd.DataFrame(raw[:5]), width='stretch')
            except Exception as ex:
                st.code(f"{type(ex).__name__}: {ex}")
        return

    # デバッグ情報 (常時表示・閉じている)
    with st.expander("🔍 デバッグ: 取得した列・先頭行", expanded=False):
        st.write("**取得列名**:", [c for c in df.columns if c != "_row"])
        st.write(f"**ヘッダー行 (シート上の行番号)**: {_header_row}")
        st.dataframe(df.head(3), width='stretch')

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

    # ---- モード切替 (新規登録 / 編集) ----
    mode_col1, mode_col2 = st.columns([1, 1])
    with mode_col1:
        new_entry_mode = st.toggle(
            "➕ 新規登録",
            value=False,
            key="bts_new_entry_mode",
            help="ONにすると新規項目登録フォームを表示します",
        )
    with mode_col2:
        edit_mode = st.toggle(
            "✏️ 編集モード",
            value=False,
            key="bts_edit_mode",
            help="ONにするとテーブル上で直接編集→保存できます",
            disabled=new_entry_mode,
        )

    # ---- 新規登録フォーム ----
    if new_entry_mode:
        # 起票者のデフォルトはログイン中ユーザー
        try:
            from .auth import get_current_user
            current_user = get_current_user() or ""
            # email 形式なら @ より前を初期値とする (例: chiba@dxhr.inc → chiba)
            default_author = current_user.split("@")[0] if "@" in current_user else current_user
        except Exception:
            default_author = ""
        _render_new_entry_form(default_author=default_author)
        st.markdown("---")
        st.caption("下に既存の一覧も表示されます。")

    # ---- フィルタ ----
    filter_col1, filter_col2 = st.columns([2, 3])
    selected_status: List[str] = []
    if "対応状況" in df.columns and not edit_mode:
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
    if not edit_mode:
        with filter_col2:
            keyword = st.text_input(
                "キーワード検索 (内容・修正方針・修正内容)", value="", key="bts_filter_kw"
            )

    filtered = df.copy()
    if not edit_mode:
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
    filtered = filtered.reset_index(drop=True)

    # 表示・編集対象の列
    display_cols = [c for c in _EXPECTED_COLS if c in filtered.columns]
    if not display_cols:
        # 期待した列名が1つもマッチしない場合のフォールバック:
        # シートから取得した全列を使う (_row は除外、文字列で表示)
        display_cols = [c for c in filtered.columns if c != "_row"]
        st.warning(
            "⚠️ 期待した列名 (項番・処理区分・内容 等) と一致しませんでした。"
            "シートから取得した列をそのまま表示しています。下のデバッグ情報をご確認ください。"
        )
        with st.expander("🔍 デバッグ: 取得した列名と先頭行", expanded=False):
            st.write("**取得列**:", list(filtered.columns))
            st.dataframe(filtered.head(3), width='stretch')

    cols_for_show = display_cols + (["_row"] if "_row" in filtered.columns else [])
    show_df = filtered[cols_for_show].copy()

    st.caption(f"表示中: {len(show_df)} 件 / 全 {len(df)} 件")

    if edit_mode:
        # ---- 編集モード: data_editor + 保存ボタン ----
        editable_cols = [c for c in display_cols if c != "項番"]  # 項番は固定
        column_config = {
            "_row": None,  # 非表示
            "項番": st.column_config.TextColumn("項番", width="small", disabled=True),
            "処理区分": st.column_config.SelectboxColumn(
                "処理区分",
                options=["", "バックエンド", "フロントエンド", "書式", "UI", "その他"],
                width="small",
            ),
            "内容": st.column_config.TextColumn("内容", width="large"),
            "対応状況": st.column_config.SelectboxColumn(
                "対応状況", options=STATUS_OPTIONS, width="small"
            ),
            "ver": st.column_config.TextColumn("ver", width="small"),
        }
        edited = st.data_editor(
            show_df,
            width='stretch',
            hide_index=True,
            column_config=column_config,
            num_rows="fixed",
            key="bts_editor",
        )
        col_save, col_cancel = st.columns([1, 4])
        with col_save:
            if st.button("💾 変更を保存", type="primary", width='stretch'):
                try:
                    n = _apply_changes(edited, show_df, editable_cols)
                    if n == 0:
                        st.info("変更はありませんでした。")
                    else:
                        st.success(f"✅ {n} セルを更新しました。")
                        _load_bts_dataframe.clear()
                        st.rerun()
                except Exception as e:
                    st.error(f"保存に失敗しました: {type(e).__name__}: {e}")
        with col_cancel:
            st.caption(
                "※ 列ヘッダをクリックすると並び替えできます。"
                "編集後は必ず「変更を保存」を押してください。"
            )
        return

    # ---- 閲覧モード ----
    view_cols = display_cols
    st.dataframe(
        show_df[view_cols],
        width='stretch',
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
