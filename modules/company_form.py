"""会社情報・社労士情報入力フォームモジュール"""
import streamlit as st
from dataclasses import dataclass, field
from typing import Optional, List

# storage は循環import回避のためフォーム送信時に遅延import


@dataclass
class Office:
    """事業所情報（本社および従たる事業所）"""
    name: str = ""                      # 事業所名
    insurance_number: str = ""          # 雇用保険適用事業所番号 (4桁-6桁-1桁)
    employee_count: int = 0             # 常時雇用する労働者数

@dataclass
class CompanyInfo:
    """会社情報を保持するデータクラス"""
    company_name: str = ""          # 雇用保険適用事業所名
    insurance_office_number: str = "" # 雇用保険適用事業所番号（4桁-6桁-1桁）
    postal_code: str = ""           # 郵便番号
    address: str = ""               # 所在地
    contact_name: str = ""          # 担当者氏名
    contact_department: str = ""    # 担当者の所属・役職
    contact_email: str = ""         # 担当者メール
    phone_number: str = ""          # 電話番号
    main_business: str = ""         # 主たる事業
    employee_count: int = 0         # 常時雇用労働者数
    labor_bureau: str = ""          # 管轄労働局
    representative_name: str = ""   # 代表者氏名
    representative_title: str = ""  # 代表者役職
    corporate_number: str = ""      # 法人番号（13桁）
    # 事業所一覧（[0]=申請事業所(本社)、[1:]=従たる事業所）
    offices: List[Office] = field(default_factory=list)

@dataclass
class SocialInsuranceLabor:
    """社労士情報を保持するデータクラス"""
    office_name: str = ""           # 事務所名
    sr_name: str = ""               # 社労士氏名
    postal_code: str = ""           # 郵便番号
    address: str = ""               # 所在地
    phone_number: str = ""          # 電話番号


# 都道府県の労働局リスト
LABOR_BUREAUS = [
    "北海道労働局", "青森労働局", "岩手労働局", "宮城労働局", "秋田労働局",
    "山形労働局", "福島労働局", "茨城労働局", "栃木労働局", "群馬労働局",
    "埼玉労働局", "千葉労働局", "東京労働局", "神奈川労働局", "新潟労働局",
    "富山労働局", "石川労働局", "福井労働局", "山梨労働局", "長野労働局",
    "岐阜労働局", "静岡労働局", "愛知労働局", "三重労働局", "滋賀労働局",
    "京都労働局", "大阪労働局", "兵庫労働局", "奈良労働局", "和歌山労働局",
    "鳥取労働局", "島根労働局", "岡山労働局", "広島労働局", "山口労働局",
    "徳島労働局", "香川労働局", "愛媛労働局", "高知労働局", "福岡労働局",
    "佐賀労働局", "長崎労働局", "熊本労働局", "大分労働局", "宮崎労働局",
    "鹿児島労働局", "沖縄労働局"
]


def _apply_imported_company(imported: 'CompanyInfo') -> None:
    """インポートされたCompanyInfoの値をsession_stateに反映
    ウィジェットのkey (_xxx) と、汎用保存キー (xxx) の両方に書き込む。
    """
    text_fields = ['company_name', 'insurance_office_number', 'postal_code',
                   'address', 'contact_name', 'contact_department', 'contact_email',
                   'phone_number', 'main_business',
                   'representative_name', 'representative_title', 'corporate_number']
    for field_name in text_fields:
        val = getattr(imported, field_name, None)
        if val is not None and val != "":
            st.session_state[field_name] = val
            st.session_state[f'_{field_name}'] = val  # ウィジェットkey にも書き込み

    # employee_count (number_input は min_value=1 のため 0 は避ける)
    emp = getattr(imported, 'employee_count', 0)
    if emp and emp >= 1:
        st.session_state['employee_count'] = emp
        st.session_state['_employee_count'] = emp

    # labor_bureau (selectbox は有効な選択肢のみ許容)
    lb = getattr(imported, 'labor_bureau', '')
    if lb and lb in LABOR_BUREAUS:
        st.session_state['labor_bureau'] = lb
        st.session_state['_labor_bureau'] = lb


def render_company_form() -> Optional[CompanyInfo]:
    """会社情報入力フォームを表示し、入力されたデータを返す"""
    st.subheader("会社情報入力")

    # === Excel一括インポート ===
    with st.expander("📂 会社情報シートから一括読み込み（任意）", expanded=False):
        st.markdown(
            "「**【助成金申請】書類準備用_会社情報シート.xlsx**」を"
            "アップロードすると、各項目が自動入力されます。"
        )
        uploaded = st.file_uploader(
            "会社情報シートを選択",
            type=['xlsx'],
            key='company_info_sheet_upload',
            label_visibility='collapsed'
        )
        if uploaded is not None:
            from .company_importer import import_company_from_excel
            imported, warnings = import_company_from_excel(uploaded.getvalue())
            if imported:
                _apply_imported_company(imported)
                st.success("✅ 会社情報を読み込みました。下のフォームで確認・修正してから保存してください。")
                for w in warnings:
                    st.info(f"ℹ️ {w}")
                # アップロードウィジェットをリセットし、フォームに値を反映
                del st.session_state['company_info_sheet_upload']
                st.rerun()
            else:
                for w in warnings:
                    st.error(w)

    # デフォルト値を session_state に初期化（widget 作成前、かつ初回のみ）
    # Streamlit の仕様: value= と key= を同時指定かつ session_state に値があると警告 + 誤動作
    if '_representative_title' not in st.session_state:
        st.session_state['_representative_title'] = "代表取締役"
    if '_labor_bureau' not in st.session_state:
        st.session_state['_labor_bureau'] = "東京労働局"
    if '_employee_count' not in st.session_state:
        st.session_state['_employee_count'] = 1

    # st.form で囲むことで、送信ボタン押下時だけ rerun する（入力中の rerun と
    # 送信クリックの取りこぼしを防ぐ。Streamlit Cloud 環境で必須）
    with st.form("company_info_form", clear_on_submit=False):
        col1, col2 = st.columns(2)

        with col1:
            company_name = st.text_input(
                "雇用保険適用事業所名 *",
                key="_company_name",
                help="会社名を正式名称で入力してください"
            )
            insurance_office_number = st.text_input(
                "雇用保険適用事業所番号 *",
                key="_insurance_office_number",
                placeholder="0000-000000-0",
                help="11桁（4桁-6桁-1桁）の形式で入力してください"
            )
            postal_code = st.text_input(
                "郵便番号 *",
                key="_postal_code",
                placeholder="000-0000",
            )
            address = st.text_area(
                "所在地 *",
                key="_address",
                height=80,
            )
            representative_name = st.text_input(
                "代表者氏名 *",
                key="_representative_name",
            )
            corporate_number = st.text_input(
                "法人番号（13桁）",
                key="_corporate_number",
                max_chars=13,
            )

        with col2:
            contact_name = st.text_input(
                "担当者氏名 *",
                key="_contact_name",
            )
            contact_department = st.text_input(
                "担当者の所属・役職",
                key="_contact_department",
            )
            contact_email = st.text_input(
                "担当者メール",
                key="_contact_email",
                placeholder="example@company.co.jp",
            )
            phone_number = st.text_input(
                "電話番号 *",
                key="_phone_number",
                placeholder="00-0000-0000",
            )
            main_business = st.text_input(
                "主たる事業 *",
                key="_main_business",
                help="例：情報サービス業、製造業など",
            )
            representative_title = st.text_input(
                "代表者役職",
                key="_representative_title",
            )
            employee_count = st.number_input(
                "常時雇用労働者数 *",
                min_value=1,
                key="_employee_count",
            )
            labor_bureau = st.selectbox(
                "管轄労働局 *",
                options=LABOR_BUREAUS,
                key="_labor_bureau",
            )

        # 送信ボタン（form_submit_button で確実に送信）
        st.markdown("---")
        submit_clicked = st.form_submit_button(
            "💾 保存して次へ", type="primary", use_container_width=True
        )

    if not submit_clicked:
        return None

    # バリデーション
    errors = []
    if not company_name:
        errors.append("雇用保険適用事業所名は必須です")
    if not insurance_office_number:
        errors.append("雇用保険適用事業所番号は必須です")
    else:
        parts = insurance_office_number.replace('−', '-').replace('ー', '-').split('-')
        if len(parts) != 3 or len(parts[0]) != 4 or len(parts[1]) != 6 or len(parts[2]) != 1:
            errors.append("雇用保険適用事業所番号は4桁-6桁-1桁の形式で入力してください")
    if not postal_code:
        errors.append("郵便番号は必須です")
    if not address:
        errors.append("所在地は必須です")
    if not representative_name:
        errors.append("代表者氏名は必須です")
    if not contact_name:
        errors.append("担当者氏名は必須です")
    if not phone_number:
        errors.append("電話番号は必須です")
    if not main_business:
        errors.append("主たる事業は必須です")

    if errors:
        for error in errors:
            st.error(error)
        return None

    # 既存のofficesを保持
    existing = st.session_state.get('company_info')
    existing_offices = list(existing.offices) if existing and getattr(existing, 'offices', None) else []

    company_info = CompanyInfo(
        company_name=company_name,
        insurance_office_number=insurance_office_number,
        postal_code=postal_code,
        address=address,
        contact_name=contact_name,
        contact_department=contact_department,
        contact_email=contact_email,
        phone_number=phone_number,
        main_business=main_business,
        employee_count=employee_count,
        labor_bureau=labor_bureau,
        representative_name=representative_name,
        representative_title=representative_title,
        corporate_number=corporate_number,
        offices=existing_offices,
    )

    for _fld in ['company_name', 'insurance_office_number', 'postal_code', 'address', 'contact_name',
                 'contact_department', 'contact_email',
                 'phone_number', 'main_business', 'employee_count', 'labor_bureau',
                 'representative_name', 'representative_title', 'corporate_number']:
        st.session_state[_fld] = getattr(company_info, _fld)
    st.session_state['company_info'] = company_info

    try:
        from . import storage
        storage.save_company(company_info)
    except Exception as e:
        st.warning(f"⚠️ 永続化ファイルへの保存に失敗しました: {e}")

    st.session_state.current_step = 2
    st.rerun()
    return company_info


def render_sr_form() -> Optional[SocialInsuranceLabor]:
    """社労士情報入力フォームを表示し、入力されたデータを返す"""
    st.subheader("社労士情報入力")

    # st.form で囲むことで、送信ボタン押下時だけ rerun する
    with st.form("sr_info_form", clear_on_submit=False):
        col1, col2 = st.columns(2)

        with col1:
            office_name = st.text_input(
                "事務所名 *",
                key="_sr_office_name",
            )
            sr_name = st.text_input(
                "社労士氏名 *",
                key="_sr_name",
            )
            sr_postal_code = st.text_input(
                "郵便番号 *",
                key="_sr_postal_code",
                placeholder="000-0000"
            )

        with col2:
            sr_address = st.text_area(
                "所在地 *",
                key="_sr_address",
                height=80
            )
            sr_phone_number = st.text_input(
                "電話番号 *",
                key="_sr_phone_number",
                placeholder="00-0000-0000"
            )

        submit_clicked = st.form_submit_button(
            "保存して次へ", type="primary", use_container_width=True
        )

    if not submit_clicked:
        return None

    errors = []
    if not office_name:
        errors.append("事務所名は必須です")
    if not sr_name:
        errors.append("社労士氏名は必須です")
    if not sr_postal_code:
        errors.append("郵便番号は必須です")
    if not sr_address:
        errors.append("所在地は必須です")
    if not sr_phone_number:
        errors.append("電話番号は必須です")

    if errors:
        for error in errors:
            st.error(error)
        return None

    sr_info = SocialInsuranceLabor(
        office_name=office_name,
        sr_name=sr_name,
        postal_code=sr_postal_code,
        address=sr_address,
        phone_number=sr_phone_number
    )

    st.session_state['sr_office_name'] = office_name
    st.session_state['sr_name'] = sr_name
    st.session_state['sr_postal_code'] = sr_postal_code
    st.session_state['sr_address'] = sr_address
    st.session_state['sr_phone_number'] = sr_phone_number
    st.session_state['sr_info'] = sr_info

    try:
        from . import storage
        storage.save_sr(sr_info)
    except Exception as e:
        st.warning(f"⚠️ 永続化ファイルへの保存に失敗しました: {e}")

    st.session_state.current_step = 3
    st.rerun()
    return sr_info


def get_saved_company_info() -> Optional[CompanyInfo]:
    """セッションから保存された会社情報を取得"""
    return st.session_state.get('company_info')


def get_saved_sr_info() -> Optional[SocialInsuranceLabor]:
    """セッションから保存された社労士情報を取得"""
    return st.session_state.get('sr_info')
