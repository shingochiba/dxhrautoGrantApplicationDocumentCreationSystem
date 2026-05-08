"""受講者一覧Excelファイルの処理モジュール"""
import pandas as pd
import streamlit as st
from typing import Optional, Dict, List
from dataclasses import dataclass
from datetime import datetime
import io


@dataclass
class Participant:
    """受講者情報を保持するデータクラス"""
    no: int
    name: str                        # 氏名
    insurance_number: str            # 雇用保険被保険者番号
    employment_type: str             # 雇用区分（旧：雇用形態）
    curriculum_name: str             # カリキュラム名
    start_date: datetime             # 訓練開始日
    end_date: datetime               # 訓練終了日
    tuition_fee: float               # 受講料
    subsidy_course: str              # 助成コース
    # 拡張フィールド（新テンプレート対応、既存ファイルでは空のまま）
    furigana: str = ""               # フリガナ
    gender: str = ""                 # 性別 (男/女)
    email: str = ""                  # メールアドレス（eラーニング時必須）
    office_name: str = ""            # 所属事業所（本社以外の場合）
    location_home: str = ""          # 受講場所（在宅の場合）
    location_name: str = ""          # 受講場所（名称・住所）
    contract_checked: str = ""       # 雇用契約書チェック


@dataclass
class CurriculumGroup:
    """カリキュラム毎のグループ情報"""
    curriculum_name: str             # カリキュラム名
    subsidy_course: str              # 助成コース
    start_date: datetime             # 訓練開始日
    end_date: datetime               # 訓練終了日
    total_fee: float                 # 合計受講料
    participants: List[Participant]  # 受講者リスト


def parse_date(date_val) -> Optional[datetime]:
    """日付をパースする"""
    if pd.isna(date_val):
        return None
    if isinstance(date_val, datetime):
        return date_val
    if isinstance(date_val, str):
        for fmt in ['%Y/%m/%d', '%Y-%m-%d', '%Y年%m月%d日']:
            try:
                return datetime.strptime(date_val, fmt)
            except ValueError:
                continue
    return None


def validate_insurance_number(number: str) -> bool:
    """雇用保険被保険者番号の形式を検証（4桁-6桁-1桁）"""
    if not number or pd.isna(number):
        return False
    parts = str(number).replace('−', '-').replace('ー', '-').split('-')
    if len(parts) != 3:
        return False
    return len(parts[0]) == 4 and len(parts[1]) == 6 and len(parts[2]) == 1


def read_participant_excel(uploaded_file) -> Optional[pd.DataFrame]:
    """アップロードされたExcelファイルを読み込む"""
    try:
        df = pd.read_excel(uploaded_file, sheet_name=0)

        # カラム名の正規化（改行や空白を削除）
        df.columns = [str(col).replace('\n', '').replace(' ', '').strip() for col in df.columns]

        # 必須カラム + オプションカラム
        # 「雇用区分」と旧「雇用形態」を相互互換にする
        required_columns = ['氏名', '雇用保険被保険者番号',
                            'カリキュラム名', '訓練開始日', '訓練終了日',
                            '受講料', '助成コース']
        optional_columns = ['No', 'フリガナ', '性別', 'メールアドレス',
                            '所属事業所', '受講場所(在宅の場合)', '受講場所(名称・住所)',
                            '雇用契約書チェック']

        def find_column(target: str) -> Optional[str]:
            """カラム名を緩く照合（カッコ・全半角差を吸収）"""
            t_norm = target.replace('(', '（').replace(')', '）')
            for actual in df.columns:
                a_norm = actual.replace('(', '（').replace(')', '）')
                if t_norm in a_norm or target in actual:
                    return actual
            return None

        # 雇用区分（または旧称: 雇用形態）
        emp_col = find_column('雇用区分') or find_column('雇用形態')
        column_mapping = {}
        if emp_col:
            column_mapping[emp_col] = '雇用区分'

        missing = []
        for col in required_columns:
            actual = find_column(col)
            if actual:
                column_mapping[actual] = col
            else:
                missing.append(col)
        if not emp_col:
            missing.append('雇用区分（または雇用形態）')

        if missing:
            st.error(f"以下のカラムが見つかりません: {', '.join(missing)}")
            return None

        for col in optional_columns:
            actual = find_column(col)
            if actual and actual not in column_mapping:
                column_mapping[actual] = col

        df = df.rename(columns=column_mapping)
        df = df.dropna(how='all')
        df = df[df['氏名'].notna() & (df['氏名'] != '')]
        # 例行（氏名が「例）」「(例)」を含むなど）を除去
        df = df[~df['氏名'].astype(str).str.match(r'^[\(（]?例[\)）]?\s')]
        return df

    except Exception as e:
        st.error(f"Excelファイルの読み込みエラー: {str(e)}")
        return None


def _normalize_gender(value) -> str:
    """性別の値を '男'/'女'/'' に正規化"""
    if value is None or pd.isna(value):
        return ''
    s = str(value).strip()
    if s in ('男', '男性', 'M', 'm', 'male', 'Male'):
        return '男'
    if s in ('女', '女性', 'F', 'f', 'female', 'Female'):
        return '女'
    return s  # 不明値はそのまま保持


def _safe_str(value) -> str:
    """NaN/None を空文字に変換"""
    if value is None or pd.isna(value):
        return ''
    return str(value).strip()


def validate_participants(df: pd.DataFrame) -> tuple[List[str], List[Participant]]:
    """受講者データのバリデーションを行う"""
    errors = []
    participants = []

    for idx, row in df.iterrows():
        row_num = idx + 2  # ヘッダー行を考慮

        # 氏名チェック
        if pd.isna(row.get('氏名')) or str(row.get('氏名')).strip() == '':
            errors.append(f"行{row_num}: 氏名が入力されていません")
            continue

        # 被保険者番号チェック
        insurance_num = str(row.get('雇用保険被保険者番号', ''))
        if not validate_insurance_number(insurance_num):
            errors.append(f"行{row_num}: 被保険者番号の形式が不正です（4桁-6桁-1桁の形式で入力してください）")

        # 雇用区分チェック（「正規雇用労働者等」「有期契約労働者等」など各種表記を許容）
        emp_type = _safe_str(row.get('雇用区分', '') or row.get('雇用形態', ''))
        if not emp_type:
            errors.append(f"行{row_num}: 雇用区分が入力されていません")
        elif not ('正規' in emp_type or '有期' in emp_type):
            errors.append(f"行{row_num}: 雇用区分は「正規雇用」「有期契約」等で入力してください（実際の値: {emp_type}）")

        # カリキュラム名チェック
        curriculum = row.get('カリキュラム名')
        if pd.isna(curriculum) or str(curriculum).strip() == '':
            errors.append(f"行{row_num}: カリキュラム名が入力されていません")

        # 日付チェック
        start_date = parse_date(row.get('訓練開始日'))
        end_date = parse_date(row.get('訓練終了日'))
        if not start_date:
            errors.append(f"行{row_num}: 訓練開始日の形式が不正です")
        if not end_date:
            errors.append(f"行{row_num}: 訓練終了日の形式が不正です")
        if start_date and end_date and start_date > end_date:
            errors.append(f"行{row_num}: 訓練開始日が終了日より後になっています")

        # 受講料チェック
        tuition = row.get('受講料(税抜)', row.get('受講料', 0))
        try:
            tuition = float(tuition) if not pd.isna(tuition) else 0
        except (ValueError, TypeError):
            errors.append(f"行{row_num}: 受講料は数値で入力してください")
            tuition = 0

        # 助成コースチェック
        subsidy = _safe_str(row.get('助成コース', ''))
        if subsidy not in ['リスキリング', '定額制']:
            errors.append(f"行{row_num}: 助成コースは「リスキリング」または「定額制」を入力してください")

        # 拡張フィールド（オプション）
        gender = _normalize_gender(row.get('性別'))
        furigana = _safe_str(row.get('フリガナ'))
        email = _safe_str(row.get('メールアドレス'))
        office_name = _safe_str(row.get('所属事業所'))
        location_home = _safe_str(row.get('受講場所(在宅の場合)'))
        location_name = _safe_str(row.get('受講場所(名称・住所)'))
        contract_checked = _safe_str(row.get('雇用契約書チェック'))

        # Participantオブジェクト作成
        if start_date and end_date:
            participant = Participant(
                no=row_num - 1,
                name=str(row.get('氏名', '')).strip(),
                insurance_number=insurance_num.replace('−', '-').replace('ー', '-'),
                employment_type=emp_type,
                curriculum_name=str(curriculum).strip() if curriculum else '',
                start_date=start_date,
                end_date=end_date,
                tuition_fee=tuition,
                subsidy_course=subsidy,
                furigana=furigana,
                gender=gender,
                email=email,
                office_name=office_name,
                location_home=location_home,
                location_name=location_name,
                contract_checked=contract_checked,
            )
            participants.append(participant)

    return errors, participants


def group_by_curriculum(participants: List[Participant]) -> Dict[str, CurriculumGroup]:
    """受講者をカリキュラム毎にグループ化する"""
    groups = {}

    for p in participants:
        key = f"{p.curriculum_name}_{p.subsidy_course}"

        if key not in groups:
            groups[key] = CurriculumGroup(
                curriculum_name=p.curriculum_name,
                subsidy_course=p.subsidy_course,
                start_date=p.start_date,
                end_date=p.end_date,
                total_fee=0,
                participants=[]
            )

        groups[key].participants.append(p)
        groups[key].total_fee += p.tuition_fee

        # 日付の範囲を更新
        if p.start_date < groups[key].start_date:
            groups[key].start_date = p.start_date
        if p.end_date > groups[key].end_date:
            groups[key].end_date = p.end_date

    return groups


def render_upload_form() -> Optional[Dict[str, CurriculumGroup]]:
    """受講者一覧アップロードフォームを表示"""
    st.subheader("受講者一覧アップロード")

    # テンプレートダウンロードリンク
    st.info("📁 受講者一覧テンプレートをダウンロードして、受講者情報を入力してください。")

    # ファイルアップロード
    uploaded_file = st.file_uploader(
        "受講者一覧Excelファイルをアップロード",
        type=['xlsx', 'xls'],
        help="受講者一覧テンプレートに入力したファイルをアップロードしてください"
    )

    if uploaded_file is not None:
        # ファイル読み込み
        df = read_participant_excel(uploaded_file)

        if df is not None:
            st.success(f"✅ {len(df)}件のデータを読み込みました")

            # プレビュー表示
            st.subheader("データプレビュー")
            st.dataframe(df, use_container_width=True)

            # バリデーション
            errors, participants = validate_participants(df)

            if errors:
                st.warning("⚠️ 以下のエラーがあります。修正してください。")
                for error in errors[:10]:  # 最大10件表示
                    st.error(error)
                if len(errors) > 10:
                    st.error(f"...他 {len(errors) - 10}件のエラー")
                return None

            # カリキュラム毎にグループ化
            groups = group_by_curriculum(participants)

            # カリキュラム一覧表示
            st.subheader("カリキュラム一覧")
            for key, group in groups.items():
                with st.expander(f"📚 {group.curriculum_name}（{group.subsidy_course}）- {len(group.participants)}名", expanded=True):
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("受講者数", f"{len(group.participants)}名")
                    with col2:
                        st.metric("訓練期間", f"{group.start_date.strftime('%Y/%m/%d')} ～ {group.end_date.strftime('%Y/%m/%d')}")
                    with col3:
                        st.metric("合計受講料", f"¥{group.total_fee:,.0f}")

                    # 受講者リスト
                    st.write("**受講者:**")
                    for p in group.participants:
                        st.write(f"  - {p.name}（{p.insurance_number}）")

            # セッションに保存
            st.session_state['participant_groups'] = groups
            st.session_state['all_participants'] = participants

            return groups

    return None


def get_saved_participant_groups() -> Optional[Dict[str, CurriculumGroup]]:
    """セッションから保存された受講者グループを取得"""
    return st.session_state.get('participant_groups')
