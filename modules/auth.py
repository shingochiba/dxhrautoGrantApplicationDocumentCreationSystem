"""
シンプルなパスワード認証モジュール。

認証情報は st.secrets から読み込む。
- ローカル開発: .streamlit/secrets.toml
- Streamlit Cloud: アプリ設定の「Secrets」で設定

secrets.toml の形式:
    [passwords]
    "user@example.com" = "password"
"""
from __future__ import annotations

import hashlib
from typing import Dict, Optional

import streamlit as st


def _get_credentials() -> Dict[str, str]:
    """st.secrets から [passwords] セクションを取得"""
    try:
        pw_section = st.secrets.get("passwords", None)
        if pw_section is None:
            return {}
        return {str(k): str(v) for k, v in pw_section.items()}
    except Exception:
        return {}


def check_password(username: str, password: str) -> bool:
    """ユーザー名とパスワードが正しいか確認"""
    creds = _get_credentials()
    # メールアドレスは大文字小文字を区別しない
    username_lower = username.strip().lower()
    creds_lower = {k.lower(): v for k, v in creds.items()}
    return creds_lower.get(username_lower) == password


def get_user_id(username: str) -> str:
    """ユーザー識別子（ファイル名用、短いハッシュ）"""
    h = hashlib.md5(username.strip().lower().encode("utf-8")).hexdigest()
    return h[:12]


def get_current_user() -> Optional[str]:
    """ログイン中のユーザー名（メールアドレス）"""
    return st.session_state.get("_authenticated_user")


def is_authenticated() -> bool:
    return bool(st.session_state.get("_authenticated_user"))


def logout() -> None:
    """ログアウト: 認証情報を session_state から除去"""
    # 全セッションをクリア（他のセッションデータも含めて）
    for k in list(st.session_state.keys()):
        del st.session_state[k]


def render_login_gate() -> bool:
    """
    ログインゲートを表示。
    認証済みなら True、未認証ならログインフォームを表示して False を返す。
    """
    if is_authenticated():
        return True

    creds = _get_credentials()
    if not creds:
        st.error("⚠️ サーバー設定エラー: 認証情報（secrets）が設定されていません。管理者にお問い合わせください。")
        st.caption("開発者向け: `.streamlit/secrets.toml` もしくは Streamlit Cloud のSecretsに [passwords] セクションを設定してください。")
        return False

    st.title("📄 書類作成ツール")
    st.markdown("### ログイン")
    st.markdown("メールアドレスとパスワードを入力してください。")

    # st.form で囲むことで、入力値とクリックイベントを atomically に送信する
    # （Streamlit Cloud の遅延環境で「1回目失敗→2回目成功」になる問題を回避）
    with st.form("login_form", clear_on_submit=False):
        username = st.text_input(
            "メールアドレス", key="_login_username", autocomplete="username"
        )
        password = st.text_input(
            "パスワード", type="password", key="_login_password",
            autocomplete="current-password"
        )
        clicked = st.form_submit_button(
            "ログイン", type="primary", use_container_width=True
        )

    if clicked:
        if check_password(username, password):
            st.session_state["_authenticated_user"] = username.strip().lower()
            # ログインフォームの入力値をクリア
            for k in ("_login_username", "_login_password"):
                st.session_state.pop(k, None)
            st.rerun()
        else:
            st.error("メールアドレスまたはパスワードが正しくありません")

    return False
