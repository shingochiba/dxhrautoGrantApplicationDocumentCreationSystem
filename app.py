"""
人材開発支援助成金 書類自動作成ツール
Streamlitメインアプリケーション
"""
import warnings
# openpyxl の無害な UserWarning（Data Validation extension 等）を全セッションで抑制
warnings.filterwarnings("ignore", category=UserWarning, module=r"openpyxl\..*")

import streamlit as st
from datetime import datetime
from pathlib import Path
import os

# モジュールのインポート
from modules.company_form import (
    render_company_form, render_sr_form,
    get_saved_company_info, get_saved_sr_info,
    CompanyInfo, SocialInsuranceLabor
)
from modules.upload_handler import (
    render_upload_form, get_saved_participant_groups
)
from modules.document_generator import DocumentGenerator, FORMAT_CONFIGS, get_format_label
from modules.excel_writer import is_teigaku_course
from modules.storage import load_saved_company, load_saved_sr, clear_saved
from modules.auth import render_login_gate, get_current_user, logout

# アプリの設定
st.set_page_config(
    page_title="人材開発支援助成金 書類自動作成ツール",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Streamlit Cloud が注入する右上アイコン（Share / Star / Edit / GitHub）を非表示にする
st.markdown(
    """
    <style>
    /* ヘッダーごと非表示（Streamlit Cloud の viewer badge / GitHub リンク等を含む） */
    header[data-testid="stHeader"] {display: none !important;}
    /* 念のためツールバー系も非表示 */
    [data-testid="stToolbar"] {display: none !important;}
    [data-testid="stToolbarActions"] {display: none !important;}
    [data-testid="stDecoration"] {display: none !important;}
    /* Streamlit Cloud の viewer badge（古いバージョン向け） */
    .viewerBadge_container__1QSob,
    .viewerBadge_link__1S137 {display: none !important;}
    </style>
    """,
    unsafe_allow_html=True,
)

# ベースディレクトリ
BASE_DIR = Path(__file__).parent


def init_session_state():
    """セッション状態の初期化（前回入力をJSONから復元）"""
    if 'current_step' not in st.session_state:
        st.session_state.current_step = 1
    if 'participant_groups' not in st.session_state:
        st.session_state.participant_groups = None

    # アプリ起動時に前回入力をファイルから復元
    if '_restored_from_storage' not in st.session_state:
        st.session_state['_restored_from_storage'] = True

        # LABOR_BUREAUS 参照のため遅延import
        from modules.company_form import LABOR_BUREAUS

        saved_company = load_saved_company()
        if saved_company:
            st.session_state['company_info'] = saved_company
            text_fields = ['company_name', 'insurance_office_number', 'postal_code',
                           'address', 'contact_name', 'contact_department', 'contact_email',
                           'phone_number', 'main_business',
                           'representative_name', 'representative_title', 'corporate_number']
            for field_name in text_fields:
                val = getattr(saved_company, field_name, None)
                if val is not None and val != "":
                    st.session_state[field_name] = val
                    st.session_state[f'_{field_name}'] = val

            # employee_count（number_input は min_value=1）
            emp = getattr(saved_company, 'employee_count', 0)
            if emp and emp >= 1:
                st.session_state['employee_count'] = emp
                st.session_state['_employee_count'] = emp

            # labor_bureau（selectbox は有効な選択肢のみ）
            lb = getattr(saved_company, 'labor_bureau', '')
            if lb and lb in LABOR_BUREAUS:
                st.session_state['labor_bureau'] = lb
                st.session_state['_labor_bureau'] = lb

        saved_sr = load_saved_sr()
        if saved_sr:
            st.session_state['sr_info'] = saved_sr
            for src, dst in [
                (saved_sr.office_name, 'sr_office_name'),
                (saved_sr.sr_name, 'sr_name'),
                (saved_sr.postal_code, 'sr_postal_code'),
                (saved_sr.address, 'sr_address'),
                (saved_sr.phone_number, 'sr_phone_number'),
            ]:
                if src:
                    st.session_state[dst] = src
                    st.session_state[f'_{dst}'] = src

    if 'company_info' not in st.session_state:
        st.session_state.company_info = None
    if 'sr_info' not in st.session_state:
        st.session_state.sr_info = None


def render_sidebar():
    """サイドバーの描画"""
    with st.sidebar:
        st.title("📄 書類作成ツール")

        # ログインユーザー情報とログアウト
        user = get_current_user()
        if user:
            st.caption(f"👤 {user}")
            if st.button("🚪 ログアウト", use_container_width=True, key="sidebar_logout"):
                logout()
                st.rerun()

        st.markdown("---")

        # ステップ表示
        steps = [
            ("1️⃣", "会社情報入力"),
            ("2️⃣", "社労士情報入力"),
            ("3️⃣", "受講者一覧アップロード"),
            ("4️⃣", "生成設定"),
            ("5️⃣", "書類生成")
        ]

        for i, (icon, name) in enumerate(steps, 1):
            if i == st.session_state.current_step:
                st.markdown(f"**{icon} {name}** ◀")
            elif i < st.session_state.current_step:
                st.markdown(f"✅ {name}")
            else:
                st.markdown(f"⬜ {name}")

        st.markdown("---")

        # 入力状況
        st.subheader("入力状況")
        company = get_saved_company_info()
        sr = get_saved_sr_info()
        groups = get_saved_participant_groups()

        if company:
            st.success(f"✅ 会社: {company.company_name}")
        else:
            st.warning("⬜ 会社情報未入力")

        if sr:
            st.success(f"✅ 社労士: {sr.sr_name}")
        else:
            st.warning("⬜ 社労士情報未入力")

        if groups:
            st.success(f"✅ カリキュラム: {len(groups)}件")
        else:
            st.warning("⬜ 受講者未アップロード")

        st.markdown("---")

        # リセット系ボタン
        if st.button("🔄 最初からやり直す", use_container_width=True,
                     help="ステップをリセットします（前回入力値は残ります）"):
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.rerun()

        if st.button("🗑️ 保存した入力データを削除", use_container_width=True,
                     help="永続化された会社情報・社労士情報を完全に削除します"):
            clear_saved()
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            st.success("保存データを削除しました")
            st.rerun()


def render_step1():
    """Step 1: 会社情報入力"""
    st.header("Step 1: 会社情報入力")
    st.markdown("申請に必要な会社の基本情報を入力してください。")

    # render_company_form は送信成功時に内部で current_step=2 にして st.rerun() する
    render_company_form()

    # 既に保存済みなら「次へ進む」ボタンを表示（フォールバック）
    saved = st.session_state.get('company_info')
    if saved:
        st.markdown("---")
        st.info(f"✅ 前回保存された情報: {saved.company_name}")
        if st.button("Step 2（社労士情報）へ進む ▶", type="primary",
                     use_container_width=True, key="step1_next_button"):
            st.session_state.current_step = 2
            st.rerun()


def render_step2():
    """Step 2: 社労士情報入力"""
    st.header("Step 2: 社労士情報入力")
    st.markdown("提出代行を行う社労士の情報を入力してください。")

    col1, col2 = st.columns([3, 1])
    with col2:
        if st.button("◀ 戻る", use_container_width=True):
            st.session_state.current_step = 1
            st.rerun()

    # render_sr_form は送信成功時に内部で current_step=3 にして st.rerun() する
    render_sr_form()

    # 既に保存済みならフォールバックボタン
    saved_sr = st.session_state.get('sr_info')
    if saved_sr:
        st.markdown("---")
        st.info(f"✅ 前回保存された情報: {saved_sr.sr_name}")
        if st.button("Step 3（受講者一覧）へ進む ▶", type="primary",
                     use_container_width=True, key="step2_next_button"):
            st.session_state.current_step = 3
            st.rerun()


def render_step3():
    """Step 3: 受講者一覧アップロード"""
    st.header("Step 3: 受講者一覧アップロード")
    st.markdown("受講者一覧のExcelファイルをアップロードしてください。")

    col1, col2 = st.columns([3, 1])
    with col2:
        if st.button("◀ 戻る", use_container_width=True):
            st.session_state.current_step = 2
            st.rerun()

    # テンプレートダウンロード
    template_path = BASE_DIR / "templates" / "受講者一覧_入力テンプレート.xlsx"
    if template_path.exists():
        with open(template_path, "rb") as f:
            st.download_button(
                label="📥 受講者一覧テンプレートをダウンロード",
                data=f.read(),
                file_name="受講者一覧_入力テンプレート.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    st.markdown("---")

    result = render_upload_form()

    if result:
        if st.button("次へ進む ▶", type="primary", use_container_width=True):
            st.session_state.current_step = 4
            st.rerun()


def render_step4():
    """Step 4: 生成設定"""
    st.header("Step 4: 生成設定")
    st.markdown("生成する書類の種類とカリキュラムを選択してください。")

    col1, col2 = st.columns([3, 1])
    with col2:
        if st.button("◀ 戻る", use_container_width=True):
            st.session_state.current_step = 3
            st.rerun()

    groups = get_saved_participant_groups()

    if not groups:
        st.error("受講者一覧がアップロードされていません。")
        return

    st.subheader("書式の選択")
    format_options = list(FORMAT_CONFIGS.keys())
    format_id = st.selectbox(
        "書式を選択",
        options=format_options,
        format_func=get_format_label,
        index=0,
        key="format_id_select",
        help="生成する書類の書式を選択してください"
    )
    # 自動入力非対応の書式の場合は注意書き
    if not FORMAT_CONFIGS[format_id]["auto_fill"]:
        st.warning(
            f"⚠️ **{get_format_label(format_id)}** は現在自動入力に対応していません。"
            "テンプレートファイル一式をZIPで出力します（中身は手動で入力してください）。"
        )

    st.subheader("書類種類の選択")
    col1, col2 = st.columns(2)
    with col1:
        generate_plan = st.checkbox("📋 計画申請書類を生成", value=True)
    with col2:
        generate_payment = st.checkbox("💰 支給申請書類を生成", value=True)

    st.subheader("カリキュラムの選択")
    selected_curricula = []
    for key, group in groups.items():
        teigaku = is_teigaku_course(group.subsidy_course)
        teigaku_mark = "🟢 **定額制として認識**" if teigaku else "⚪ 通常コース"
        if st.checkbox(
            f"📚 {group.curriculum_name}（{group.subsidy_course}）- {len(group.participants)}名  ｜  {teigaku_mark}",
            value=True,
            key=f"curriculum_{key}"
        ):
            selected_curricula.append(key)

    st.subheader("提出日")
    submit_date = st.date_input("提出予定日", value=datetime.now())

    # セッションに保存
    st.session_state['generate_plan'] = generate_plan
    st.session_state['generate_payment'] = generate_payment
    st.session_state['selected_curricula'] = selected_curricula
    st.session_state['submit_date'] = datetime.combine(submit_date, datetime.min.time())
    st.session_state['format_id'] = format_id

    if selected_curricula and (generate_plan or generate_payment):
        if st.button("📄 書類を生成する", type="primary", use_container_width=True):
            st.session_state.current_step = 5
            st.rerun()
    else:
        st.warning("生成する書類種類とカリキュラムを選択してください。")


def render_step5():
    """Step 5: 書類生成"""
    st.header("Step 5: 書類生成")

    col1, col2 = st.columns([3, 1])
    with col2:
        if st.button("◀ 設定に戻る", use_container_width=True):
            st.session_state.current_step = 4
            st.rerun()

    company = get_saved_company_info()
    sr = get_saved_sr_info()
    groups = get_saved_participant_groups()

    if not company or not sr or not groups:
        st.error("必要な情報が不足しています。前のステップを確認してください。")
        return

    generate_plan = st.session_state.get('generate_plan', True)
    generate_payment = st.session_state.get('generate_payment', True)
    selected_curricula = st.session_state.get('selected_curricula', list(groups.keys()))
    submit_date = st.session_state.get('submit_date', datetime.now())
    format_id = st.session_state.get('format_id', 'current')

    # 選択された書式を表示
    st.info(f"📋 書式: **{get_format_label(format_id)}**")
    if not FORMAT_CONFIGS[format_id]["auto_fill"]:
        st.warning(
            "この書式は自動入力に対応していないため、テンプレート一式をそのまま出力します。"
        )

    # 各カリキュラムの定額制判定を表示
    with st.expander("🔍 生成対象カリキュラム (定額制判定の確認)", expanded=True):
        for key in selected_curricula:
            grp = groups.get(key)
            if not grp:
                continue
            teigaku = is_teigaku_course(grp.subsidy_course)
            if teigaku:
                st.success(
                    f"🟢 **{grp.curriculum_name}** （助成コース: `{grp.subsidy_course}`） "
                    f"→ **定額制として認識** → 様式3-2号 / 様式11号(定額制版) 等を生成"
                )
            else:
                st.info(
                    f"⚪ **{grp.curriculum_name}** （助成コース: `{grp.subsidy_course}`） "
                    f"→ 通常コース → 様式3-1号 / 様式11号(通常版) 等を生成"
                )

    # 生成実行
    if 'generated_files' not in st.session_state:
        with st.spinner("書類を生成中..."):
            progress_bar = st.progress(0)
            status_text = st.empty()

            try:
                generator = DocumentGenerator(str(BASE_DIR))

                total_curricula = len(selected_curricula)
                generated_files = {}

                for idx, curriculum_key in enumerate(selected_curricula):
                    group = groups[curriculum_key]
                    status_text.text(f"生成中: {group.curriculum_name}")
                    progress_bar.progress((idx + 1) / total_curricula)

                    result = generator.generate_documents(
                        company=company,
                        sr=sr,
                        groups={curriculum_key: group},
                        generate_plan=generate_plan,
                        generate_payment=generate_payment,
                        selected_curricula=[curriculum_key],
                        submit_date=submit_date,
                        format_id=format_id,
                    )
                    generated_files.update(result)

                st.session_state['generated_files'] = generated_files
                progress_bar.progress(1.0)
                status_text.text("生成完了！")

            except Exception as e:
                st.error(f"書類生成中にエラーが発生しました: {str(e)}")
                import traceback
                st.code(traceback.format_exc())
                return

    # 結果表示
    generated_files = st.session_state.get('generated_files', {})

    if generated_files:
        st.success(f"✅ {len(generated_files)}件のZIPファイルを生成しました！")

        st.subheader("生成されたファイル")
        import zipfile as _zipfile
        for curriculum_key, zip_path in generated_files.items():
            group = groups.get(curriculum_key)
            if group:
                teigaku = is_teigaku_course(group.subsidy_course)
                badge = "🟢 定額制" if teigaku else "⚪ 通常"
                st.markdown(f"### 📚 {group.curriculum_name}  ｜  {badge}")

            if os.path.exists(zip_path):
                # ZIP の中身を表示（定額制かどうかが分かる）
                try:
                    with _zipfile.ZipFile(zip_path) as zf:
                        names = sorted(zf.namelist())
                    with st.expander(f"📂 ZIPの中身を確認 ({len(names)}件)", expanded=False):
                        for n in names:
                            marker = ""
                            if "3-2" in n: marker = "  ← 様式3-2号(定額制)"
                            elif "3-1" in n: marker = "  ← 様式3-1号(通常)"
                            elif "14-1" in n: marker = "  ← 様式14-1号(定額制)"
                            elif "6-3" in n: marker = "  ← 様式6-3号(定額制)"
                            elif "_定額制" in n: marker = "  ← 定額制版"
                            st.text(f"  • {n}{marker}")
                except Exception:
                    pass

                with open(zip_path, "rb") as f:
                    file_name = os.path.basename(zip_path)
                    st.download_button(
                        label=f"📥 {file_name} をダウンロード",
                        data=f.read(),
                        file_name=file_name,
                        mime="application/zip",
                        key=f"download_{curriculum_key}"
                    )
            else:
                st.error(f"ファイルが見つかりません: {zip_path}")

        st.markdown("---")
        if st.button("🔄 新しい書類を生成する", use_container_width=True):
            if 'generated_files' in st.session_state:
                del st.session_state['generated_files']
            st.session_state.current_step = 4
            st.rerun()


def render_sidebar_toggle_button():
    """サイドバー表示/非表示を切り替えるボタンを描画。
    session_state でトグル状態を管理し、CSS で直接 display を切り替える。
    JS クリック方式より確実。
    """
    if '_sidebar_hidden' not in st.session_state:
        st.session_state['_sidebar_hidden'] = False

    col_btn, _ = st.columns([1, 9])
    with col_btn:
        hidden_now = st.session_state['_sidebar_hidden']
        label = "▶ サイドバー表示" if hidden_now else "☰ サイドバー非表示"
        if st.button(
            label,
            key="_sidebar_toggle_global",
            help="サイドバーの表示・非表示を切り替えます",
        ):
            st.session_state['_sidebar_hidden'] = not hidden_now
            st.rerun()

    # 非表示状態の場合は CSS でサイドバーを完全に隠す
    if st.session_state.get('_sidebar_hidden'):
        st.markdown(
            """
            <style>
            /* サイドバー本体を非表示 */
            section[data-testid="stSidebar"] {
                display: none !important;
                width: 0 !important;
                min-width: 0 !important;
                max-width: 0 !important;
            }
            /* Streamlit が表示する「折りたたみ時の展開ボタン」も隠す（自前のボタンで管理するため） */
            [data-testid="stSidebarCollapsedControl"],
            [data-testid="collapsedControl"] {
                display: none !important;
            }
            /* メインコンテンツのマージン調整 */
            section[data-testid="stMain"] > div:first-child,
            .main > .block-container {
                padding-left: 1rem !important;
                margin-left: 0 !important;
            }
            </style>
            """,
            unsafe_allow_html=True,
        )


def main():
    """メイン関数"""
    # ログインゲート: 認証されるまで以降を表示しない
    if not render_login_gate():
        return

    init_session_state()
    render_sidebar()

    # メインコンテンツ
    render_sidebar_toggle_button()
    st.title("人材開発支援助成金 書類自動作成ツール")
    st.markdown("会社情報と受講者一覧を入力して、計画申請・支給申請の書類を自動生成します。")
    st.markdown("---")

    # 現在のステップに応じた画面を表示
    current_step = st.session_state.current_step

    if current_step == 1:
        render_step1()
    elif current_step == 2:
        render_step2()
    elif current_step == 3:
        render_step3()
    elif current_step == 4:
        render_step4()
    elif current_step == 5:
        render_step5()


if __name__ == "__main__":
    main()
