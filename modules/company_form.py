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

    # offices (申請事業所と従たる事業所の一覧) を session_state に格納
    # → 後段の save 処理 (render_company_form 内) が existing_offices として参照する
    imported_offices = getattr(imported, 'offices', None) or []
    if imported_offices:
        st.session_state['_imported_offices'] = list(imported_offices)
        # 既存の company_info に offices をマージしておく
        # (まだフォーム未送信でも、後で _imported_offices が優先採用される)
        existing_company = st.session_state.get('company_info')
        if existing_company is not None:
            existing_company.offices = list(imported_offices)


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
            # 同じファイルを毎 rerun で再処理しないよう file_id でガード。
            # （del session_state では file_uploader 表示は消えないため、フラグで管理する）
            file_id = getattr(uploaded, 'file_id', None) or uploaded.name
            last_imported = st.session_state.get('_last_imported_file_id')
            if file_id != last_imported:
                from .company_importer import import_company_from_excel
                imported, warnings = import_company_from_excel(uploaded.getvalue())
                if imported:
                    _apply_imported_company(imported)
                    st.session_state['_last_imported_file_id'] = file_id
                    st.success("✅ 会社情報を読み込みました。下のフォームで確認・修正してから保存してください。")
                    for w in warnings:
                        st.info(f"ℹ️ {w}")
                    st.rerun()
                else:
                    for w in warnings:
                        st.error(w)
            else:
                # 既にインポート済みのファイルが残っている状態（×で消すまで）
                st.success("✅ 会社情報を読み込み済みです。下のフォームで確認・修正してください。")
        else:
            # ファイルが取り除かれたら処理済みフラグをリセット（同じファイル再アップロードに備える）
            st.session_state.pop('_last_imported_file_id', None)

    # === 産業分類コード一覧（参照用・静的表示）===
    # 注: ここに text_input 等の rerun を起こすウィジェットを置くと
    # フォーム内の常時雇用労働者数（number_input）等の入力中の値がリセットされる。
    # そのため一覧は st.dataframe（rerun を起こさない静的表示）で提供する。
    # 自動反映はフォーム内の「産業分類コード」欄で行う。
    from . import storage as _storage
    _industry_map = _storage.load_industry_codes()
    with st.expander("📊 産業分類コード一覧（参照・99件）", expanded=False):
        st.caption(
            "中分類2桁コード一覧です。下のフォーム内の「産業分類コード」欄にコードを"
            "入力すると、保存時に「主たる事業」へ自動反映されます。"
            "[出典: ハローワークインターネットサービス](https://www.hellowork.mhlw.go.jp/info/industry_list02.html)"
        )
        if _industry_map:
            import pandas as _pd
            _df_industry = _pd.DataFrame(
                [{"コード": k, "分類名": v} for k, v in _industry_map.items()]
            )
            st.dataframe(
                _df_industry,
                width='stretch',
                hide_index=True,
                height=300,
            )

    # デフォルト値を session_state に初期化（widget 作成前、かつ初回のみ）
    # Streamlit の仕様: value= と key= を同時指定かつ session_state に値があると警告 + 誤動作
    if '_representative_title' not in st.session_state:
        st.session_state['_representative_title'] = "代表取締役"
    if '_labor_bureau' not in st.session_state:
        st.session_state['_labor_bureau'] = "東京労働局"
    if '_employee_count' not in st.session_state:
        st.session_state['_employee_count'] = 1

    # st.form で囲むことで、送信ボタン押下時だけ rerun する（入力中の rerun と
    # 申請事業所(本社) の情報が未取り込みの場合は注意喚起
    # (各書式の「雇用保険適用事業所名」欄は 会社情報シート ②事業所情報 B9 を反映するため、
    #  シートをインポートしないと該当欄が空欄になる)
    existing_company = st.session_state.get('company_info')
    has_offices = bool(getattr(existing_company, 'offices', None)) or bool(
        st.session_state.get('_imported_offices')
    )
    if not has_offices:
        st.warning(
            "⚠️ **事業所情報が未取り込みです。**\n\n"
            "各書式 (様式4-2 / 様式5 / 様式6-2 / 様式11 等) の「雇用保険適用事業所名」欄は、"
            "**会社情報シートの ②事業所情報シート → 申請事業所 → 事業所名 (B9)** から自動転記されます。\n\n"
            "上の **「📂 会社情報シートから一括読み込み」** を開いて、"
            "シートをアップロードしてください (再アップロード可)。"
        )

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
            industry_code_input = st.text_input(
                "産業分類コード（2桁・任意）",
                key="_industry_code_form",
                max_chars=2,
                placeholder="例: 39",
                help=(
                    "中分類2桁コードを入力すると、保存時に「主たる事業」へ"
                    "対応する分類名を自動的に反映します。"
                    "上の参照エクスパンダーでコードを確認できます。"
                ),
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
            "💾 保存して次へ", type="primary", width='stretch'
        )

    if not submit_clicked:
        return None

    # === バリデーション前の自動補完 ===
    # 産業分類コードが入力されていれば、主たる事業を分類名で上書き
    # （バリデーションより先に行うことで、コード入力時に主たる事業が空でもエラーにならない）
    _code_trim = (industry_code_input or "").strip()
    _industry_applied = False
    if _code_trim and len(_code_trim) == 2 and _code_trim.isdigit():
        _mapped_name = _industry_map.get(_code_trim)
        if _mapped_name:
            main_business = _mapped_name
            _industry_applied = True
        else:
            st.warning(f"⚠️ 産業分類コード `{_code_trim}` は存在しません。入力された主たる事業を使用します。")

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
        errors.append("主たる事業は必須です（または産業分類コードを入力してください）")

    if errors:
        for error in errors:
            st.error(error)
        return None

    # 自動補完の通知（バリデーション通過後に表示）
    if _industry_applied:
        st.info(f"📊 産業分類コード `{_code_trim}` → 「{main_business}」を主たる事業に反映しました。")

    # 住所の都道府県から労働局を自動判定（自動入力ロジック）
    # 住所から都道府県が明確に読み取れる場合は、ドロップダウンの値より住所優先
    from .company_importer import detect_prefecture, prefecture_to_labor_bureau
    detected_pref = detect_prefecture(address)
    if detected_pref:
        suggested_bureau = prefecture_to_labor_bureau(detected_pref)
        if suggested_bureau in LABOR_BUREAUS and suggested_bureau != labor_bureau:
            labor_bureau = suggested_bureau
            st.info(f"📍 住所「{detected_pref}」から管轄労働局を「{suggested_bureau}」に自動設定しました。")

    # 既存のofficesを保持: インポート直後の _imported_offices があればそれを優先、
    # 無ければ session_state['company_info'].offices を流用
    imported = st.session_state.get('_imported_offices')
    if imported:
        existing_offices = list(imported)
    else:
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

    # === 社労士マスタからの自動入力 ===
    # st.form の外に置く（form 内のウィジェット変更では rerun されないため）
    from . import storage
    master_list = storage.load_sr_master()
    if master_list:
        master_options = ["（手動入力）"] + [sr.office_name for sr in master_list]
        selected_office = st.selectbox(
            "📋 社労士マスタから選択（任意）",
            options=master_options,
            key="_sr_master_select",
            help="マスタから選ぶと、社労士氏名・郵便番号・所在地・電話番号が自動入力されます。"
                 "選択後にフィールドを手動編集することも可能です。"
        )
        # 選択が変わったときだけフィールドに反映する（手動編集を保護）
        prev_selection = st.session_state.get('_sr_master_prev_selection')
        if selected_office != prev_selection and selected_office != "（手動入力）":
            sr_from_master = next(
                (s for s in master_list if s.office_name == selected_office), None
            )
            if sr_from_master:
                st.session_state['_sr_office_name'] = sr_from_master.office_name
                st.session_state['_sr_name'] = sr_from_master.sr_name
                st.session_state['_sr_postal_code'] = sr_from_master.postal_code
                st.session_state['_sr_address'] = sr_from_master.address
                st.session_state['_sr_phone_number'] = sr_from_master.phone_number
        st.session_state['_sr_master_prev_selection'] = selected_office

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
            "保存して次へ", type="primary", width='stretch'
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
