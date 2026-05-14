"""
Excelテンプレートへのデータ書き込みモジュール (2025年度 新書式対応)

ZIPレベル操作でセル値だけを書き換えるため、テンプレート内のチェックボックスや
フォームコントロール、画像、書式、マージセルなどは全て完全に保持される。
"""
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Any

from .company_form import CompanyInfo, SocialInsuranceLabor
from .upload_handler import CurriculumGroup, Participant
from .xlsx_patcher import patch_xlsx


# 定額制コースの識別キーワード
_TEIGAKU_KEYWORDS = ("定額制", "人への投資", "人材投資")


def is_teigaku_course(subsidy_course: str) -> bool:
    if not subsidy_course:
        return False
    return any(kw in subsidy_course for kw in _TEIGAKU_KEYWORDS)


def split_postal(postal_code: str) -> tuple:
    if not postal_code:
        return ("", "")
    parts = postal_code.replace('−', '-').replace('ー', '-').replace('―', '-').split('-')
    if len(parts) >= 2:
        return (parts[0], parts[1])
    return (postal_code, "")


def split_phone(phone: str) -> tuple:
    if not phone:
        return ("", "", "")
    parts = phone.replace('−', '-').replace('ー', '-').replace('―', '-').split('-')
    if len(parts) >= 3:
        return (parts[0], parts[1], parts[2])
    return (phone, "", "")


def split_insurance_number(num: str) -> tuple:
    if not num:
        return ("", "", "")
    parts = num.replace('−', '-').replace('ー', '-').replace('―', '-').split('-')
    if len(parts) == 3:
        return (parts[0], parts[1], parts[2])
    return (num, "", "")


def labor_bureau_short(bureau: str) -> str:
    """「東京労働局」→「東京」（県名のみ）。テンプレート側に「労働局長 殿」が既にあるため重複を避ける。
    北海道労働局のみ「北海道」に保つ。
    """
    if not bureau:
        return ""
    if bureau == "北海道労働局":
        return "北海道"
    return bureau.replace("労働局", "").strip()


def get_insurance_office_name(company) -> str:
    """雇用保険適用事業所の名称を取得。
    offices[0].name があればそれを優先、なければ company.company_name を返す。
    （会社名と保険適用事業所名が異なる場合への対応：個人事業主が UNISYS幸手 のような
      事業所名を持つケースなど）
    """
    offices = getattr(company, 'offices', None) or []
    if offices and offices[0].name:
        return offices[0].name
    return company.company_name


def get_4_2_teigaku_checkboxes(subsidy_course: str) -> Dict[int, bool]:
    """様式4-2号（支給申請書）の定額制契約途中解約禁止チェックボックス

    対象: 「本申請の定額制サービスに係る契約について契約期間の終了前に途中解約しません。…」
    対応する ctrlProp（3書式共通）:
      ctrlProp14 = Check Box 20

    定額制コースの場合のみチェックする。
    """
    teigaku = is_teigaku_course(subsidy_course)
    return {14: teigaku}


def get_3_1_employment_checkboxes(participants, max_participants: int = 80) -> Dict[int, bool]:
    """様式3-1号（対象労働者一覧）の雇用形態チェックボックス状態を生成

    現行・R80302 書式共通のマッピング規則:
      参加者 idx (0-indexed) → 行 (13 + idx*2)
      - 正規雇用労働者等 チェックボックス: ctrlProp(1 + idx*2)
      - 有期契約労働者等 チェックボックス: ctrlProp(2 + idx*2)

    各参加者の employment_type に「正規」を含む → 正規側 True
                                    「有期」を含む → 有期側 True
    """
    states: Dict[int, bool] = {}
    for idx, p in enumerate(participants):
        if idx >= max_participants:
            break
        emp = getattr(p, 'employment_type', '') or ''
        regular_ctrl = 1 + idx * 2  # ctrlProp1, 3, 5, ...
        fixed_ctrl = 2 + idx * 2    # ctrlProp2, 4, 6, ...
        is_regular = '正規' in emp
        is_fixed = '有期' in emp
        states[regular_ctrl] = is_regular
        states[fixed_ctrl] = is_fixed
    return states


def get_training_method_checkboxes(subsidy_course: str) -> Dict[int, bool]:
    """助成コース → 様式1-1号 「12 訓練の実施方法」のチェックボックス状態

    対応するチェックボックス（全書式共通）:
      ctrlProp6 = ①通学制
      ctrlProp7 = ②同時双方向型の通信訓練
      ctrlProp8 = ③eラーニング
      ctrlProp9 = ④通信制

    マッピング:
      インタラクティブ（通学制）   → ①通学制
      インタラクティブ（同時双方向）→ ②同時双方向
      e-ラーニング                → ③eラーニング
      定額制                      → ③eラーニング (定額制サービスは通常 e-Learning)
    """
    s = (subsidy_course or "").strip()
    states = {6: False, 7: False, 8: False, 9: False}
    if 'インタラクティブ' in s and '通学' in s:
        states[6] = True
    elif 'インタラクティブ' in s and '双方向' in s:
        states[7] = True
    elif 'eラーニング' in s or 'e-ラーニング' in s or 'eラー' in s:
        states[8] = True
    elif '定額制' in s:
        states[8] = True
    return states


class ExcelWriter:
    """Excelテンプレートへのデータ書き込みクラス (ZIPレベル操作)"""

    def __init__(self, template_dir: str, output_dir: str):
        self.template_dir = Path(template_dir)
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def _template(self, relative_path: str) -> Path:
        p = self.template_dir / relative_path
        if not p.exists():
            raise FileNotFoundError(f"テンプレートが見つかりません: {p}")
        return p

    def _patch(self, template_rel: str, output_filename: str,
               cell_values: Dict[str, Any],
               checkbox_states: Dict[int, bool] | None = None) -> str:
        """指定テンプレートをコピーし、cell_values を書き込んで出力

        Args:
            checkbox_states: ctrlProp番号 → チェック状態 の辞書
        """
        tpl = self._template(template_rel)
        out = self.output_dir / output_filename
        return patch_xlsx(tpl, out, cell_values, checkbox_states=checkbox_states)

    # ==================================================================
    # 計画申請書類
    # ==================================================================

    def write_訓練実施計画届(self, company: CompanyInfo, sr: SocialInsuranceLabor,
                              group: CurriculumGroup, submit_date: datetime) -> str:
        """01_職業訓練実施計画届(様式第1-1号)"""
        p1, p2 = split_postal(company.postal_code)
        sp1, sp2 = split_postal(sr.postal_code) if sr else ("", "")
        cp1, cp2, cp3 = split_phone(company.phone_number)
        spp1, spp2, spp3 = split_phone(sr.phone_number) if sr else ("", "", "")
        i1, i2, i3 = split_insurance_number(company.insurance_office_number)

        male = sum(1 for p in group.participants if getattr(p, 'gender', '') in ('男', '男性'))
        female = sum(1 for p in group.participants if getattr(p, 'gender', '') in ('女', '女性'))
        if male + female == 0:
            male = len(group.participants)

        teigaku = is_teigaku_course(group.subsidy_course)
        insurance_office_name = get_insurance_office_name(company)

        # 受講場所: 1人目の受講者情報から組み立て、無ければ会社住所
        first_p = group.participants[0] if group.participants else None
        location_value = ""
        if first_p:
            if getattr(first_p, 'location_home', '').strip():
                location_value = first_p.location_home.strip()
            elif getattr(first_p, 'location_name', '').strip():
                location_value = first_p.location_name.strip()
        if not location_value:
            location_value = f"本社({company.address})"

        values = {
            # 提出日
            'AL5': submit_date.year, 'AR5': submit_date.month, 'AU5': submit_date.day,
            # 管轄労働局（県名のみ）
            'B7': labor_bureau_short(company.labor_bureau),
            # 事業主 郵便番号/住所/名称/代表者/法人番号
            'AG9': p1, 'AL9': p2,
            'AF10': company.address,
            'AF12': company.company_name,
            'AF13': company.representative_name,
            'AF14': company.corporate_number,
            # 社労士
            'AG16': sp1, 'AL16': sp2,
            'AF17': sr.address if sr else "",
            'AF19': sr.office_name if sr else "",
            'AF20': sr.sr_name if sr else "",
            'AF21': spp1, 'AM21': spp2, 'AT21': spp3,
            # 雇用保険適用事業所（offices[0].name 優先）
            'K26': insurance_office_name,
            'AN26': i1, 'AS26': i2, 'AZ26': i3,
            'M27': p1, 'Q27': p2,
            'K28': company.address,
            # 担当者
            'R29': company.contact_name,
            'AM29': company.contact_department,
            'R30': cp1, 'W30': cp2, 'AB30': cp3,
            'AM30': company.contact_email,
            # 訓練コース
            'K42': group.curriculum_name,
            'AN42': len(group.participants),
            # 訓練の実施場所（11番）
            'K47': location_value,
            # 男女別受講者数
            'S80': male, 'AN80': female,
        }
        if group.start_date:
            values['N43'] = group.start_date.year
            values['T43'] = group.start_date.month
            values['Z43'] = group.start_date.day
        if group.end_date:
            values['AI43'] = group.end_date.year
            values['AO43'] = group.end_date.month
            values['AU43'] = group.end_date.day

        # 定額制の場合: 9番「定額制サービスの契約期間」= 8番「訓練の実施期間」と同じ
        if teigaku:
            if group.start_date:
                values['N44'] = group.start_date.year
                values['T44'] = group.start_date.month
                values['Z44'] = group.start_date.day
            if group.end_date:
                values['AI44'] = group.end_date.year
                values['AO44'] = group.end_date.month
                values['AU44'] = group.end_date.day
            # 13番「訓練の時間数」の標準学習時間プレースホルダを空欄化
            # （定額制の場合は記載不要）
            values['R55'] = ""
            values['Y55'] = ""

        # 12 訓練の実施方法のチェックボックスを助成コースに応じて設定
        checkbox_states = get_training_method_checkboxes(group.subsidy_course)

        fname = f"01_職業訓練実施計画届_{company.company_name}_{group.curriculum_name}.xlsx"
        return self._patch(
            "計画申請/01_職業訓練実施計画届(様式第1-1号).xlsx", fname, values,
            checkbox_states=checkbox_states,
        )

    def write_事業展開等実施計画(self, company: CompanyInfo, group: CurriculumGroup,
                                  submit_date: datetime) -> str:
        """02_事業展開等実施計画(様式第1-3号)"""
        values = {
            # 申請事業主の証明欄の日付 (●●●●年 ●●月 ●●日)
            'H73': submit_date.year,
            'L73': submit_date.month,
            'O73': submit_date.day,
            # 代表者役職名・氏名
            'K76': company.representative_title,
            'K77': company.representative_name,
        }
        fname = f"02_事業展開等実施計画_{company.company_name}_{group.curriculum_name}.xlsx"
        return self._patch("計画申請/02_事業展開等実施計画(様式第1-3号).xlsx", fname, values)

    def write_対象者一覧_3_1(self, company: CompanyInfo, group: CurriculumGroup) -> str:
        """03_対象者一覧(様式第3-1号) 通常用
        列構成: A=No(pre-filled), B=氏名, C=被保険者番号4桁, F=6桁, I=1桁,
               J=雇用形態(チェックボックス), L=採用予定日, S=対象労働者属性
        受講者1件につき2行使用 (行13-14 が1件目、行15-16 が2件目...)
        雇用形態のチェックボックスは employment_type から自動チェック。
        """
        values = {
            'C8': company.company_name,
            'C9': group.curriculum_name,
        }
        start_row = 13
        rows_per_entry = 2
        for idx, p in enumerate(group.participants):
            r = start_row + idx * rows_per_entry
            # B列(B13:B14 merged): 氏名
            values[f'B{r}'] = p.name
            # 雇用保険被保険者番号: C(4桁) / F(6桁) / I(1桁)
            parts = (p.insurance_number or "").replace('−', '-').replace('ー', '-').split('-')
            if len(parts) == 3:
                values[f'C{r}'] = parts[0]
                values[f'F{r}'] = parts[1]
                values[f'I{r}'] = parts[2]

        # 雇用形態チェックボックス（正規雇用/有期契約）の状態を生成
        checkbox_states = get_3_1_employment_checkboxes(group.participants)

        fname = f"03_対象者一覧(3-1)_{company.company_name}_{group.curriculum_name}.xlsx"
        return self._patch(
            "計画申請/03_対象者一覧(様式第3-1号).xlsx",
            fname, values,
            checkbox_states=checkbox_states,
        )

    def write_対象者一覧_電子申請(self, company: CompanyInfo, group: CurriculumGroup) -> str:
        """03_対象者一覧(様式第3号)※電子申請のみ - 現行 R8.4 新規

        構造:
          - 雇用保険適用事業所の名称: B13 (B13:C13 merged)
          - 雇用保険適用事業所番号: B15 (B15:C15 merged)
          - データ行: row 21 から開始（row 20 は記入例）
          - B: 番号(自動採番), C: 氏名, D: フリガナ, E: 雇用保険番号(11桁),
            F: 性別, G: 生年月日, H: 取得日, I: 訓練コース名,
            J: 雇用形態, K: 採用予定日, L: 属性, M: 所属事業主名
        """
        insurance_office_name = get_insurance_office_name(company)
        # 保険番号を 11桁ハイフン区切り文字列で正規化
        ins_raw = (company.insurance_office_number or "").replace('−', '-').replace('ー', '-').strip()

        values = {
            'B13': insurance_office_name,
            'B15': ins_raw,
        }
        start_row = 21
        for idx, p in enumerate(group.participants):
            r = start_row + idx
            # 番号は B21=1, B22=B21+1 ... と自動採番されるので idx==0 のみ書き込まなくてもOK
            # ただし上書きの可能性に備えて 1人目に 1 を入れておく
            if idx == 0:
                values[f'B{r}'] = 1
            values[f'C{r}'] = p.name
            values[f'D{r}'] = getattr(p, 'furigana', '')
            values[f'E{r}'] = (p.insurance_number or "").replace('−', '-').replace('ー', '-')
            values[f'F{r}'] = getattr(p, 'gender', '')
            # G: 生年月日, H: 取得日 → 未取得のため空
            values[f'I{r}'] = group.curriculum_name
            values[f'J{r}'] = p.employment_type
            # K: 採用予定日, L: 属性 → 未取得のため空
            values[f'M{r}'] = getattr(p, 'office_name', '')

        fname = f"03_対象者一覧(電子申請)_{company.company_name}_{group.curriculum_name}.xlsx"
        return self._patch(
            "計画申請/03_対象者一覧(様式第3号)_企業名※電子申請のみ.xlsx",
            fname, values,
        )

    def write_対象者一覧_3_2(self, company: CompanyInfo, group: CurriculumGroup) -> str:
        """03_対象者一覧(様式第3-2号) 定額制用

        テンプレート列構造:
          A: 番号（自動採番済み）, B: 氏名, C-F(merged): 正規雇用マーク, G-J(merged): 有期契約マーク
        ※ このフォームには保険番号欄は無い
        """
        values = {
            'C11': company.company_name,
            'C12': group.curriculum_name,
        }
        start_row = 16
        for idx, p in enumerate(group.participants):
            r = start_row + idx
            # A{r} は事前採番済みなのでスキップ。B{r} に氏名。
            values[f'B{r}'] = p.name
            # 雇用形態に応じてマーク列を選択（C=正規雇用、G=有期契約）
            if '正規' in p.employment_type:
                values[f'C{r}'] = '○'
            elif '有期' in p.employment_type:
                values[f'G{r}'] = '○'

        fname = f"03_対象者一覧(3-2)_{company.company_name}_{group.curriculum_name}.xlsx"
        return self._patch("計画申請/03_対象者一覧(様式第3-2号)_定額制.xlsx", fname, values)

    def write_事前確認書(self, company: CompanyInfo, sr: SocialInsuranceLabor,
                         group: CurriculumGroup, submit_date: datetime,
                         teigaku: bool = False) -> str:
        """04_事前確認書(様式第11号) 通常版 / 定額制版"""
        if teigaku:
            # 定額制の場合は ※定額制 とついた書類を優先
            template = "計画申請/04_事前確認書(様式第11号)※定額制.xlsx"
            suffix = "_定額制"
        else:
            template = "計画申請/04_事前確認書(様式第11号).xlsx"
            suffix = ""

        p1, p2 = split_postal(company.postal_code)
        sp1, sp2 = split_postal(sr.postal_code) if sr else ("", "")
        cp1, cp2, cp3 = split_phone(company.phone_number)
        spp1, spp2, spp3 = split_phone(sr.phone_number) if sr else ("", "", "")

        values = {
            'E12': submit_date.year, 'I12': submit_date.month, 'L12': submit_date.day,
            'S13': p1, 'W13': p2,
            'R14': company.address,
            'R16': company.company_name,
            'R18': company.representative_name,
            'R19': cp1, 'V19': cp2, 'Z19': cp3,
            'S21': sp1, 'W21': sp2,
            'R22': sr.address if sr else "",
            'R24': sr.office_name if sr else "",
            'R26': sr.sr_name if sr else "",
            'R27': spp1, 'V27': spp2, 'Z27': spp3,
            'A30': labor_bureau_short(company.labor_bureau),
        }

        fname = f"04_事前確認書{suffix}_{company.company_name}_{group.curriculum_name}.xlsx"
        return self._patch(template, fname, values)

    # ==================================================================
    # 支給申請書類
    # ==================================================================

    def write_支給申請書(self, company: CompanyInfo, sr: SocialInsuranceLabor,
                         group: CurriculumGroup, submit_date: datetime) -> str:
        """01_支給申請書(様式第4-2号)"""
        p1, p2 = split_postal(company.postal_code)
        sp1, sp2 = split_postal(sr.postal_code) if sr else ("", "")
        cp1, cp2, cp3 = split_phone(company.phone_number)
        spp1, spp2, spp3 = split_phone(sr.phone_number) if sr else ("", "", "")
        i1, i2, i3 = split_insurance_number(company.insurance_office_number)

        male = sum(1 for p in group.participants if getattr(p, 'gender', '') in ('男', '男性'))
        female = sum(1 for p in group.participants if getattr(p, 'gender', '') in ('女', '女性'))
        if male + female == 0:
            male = len(group.participants)

        values = {
            'AL5': submit_date.year, 'AR5': submit_date.month, 'AU5': submit_date.day,
            'B7': labor_bureau_short(company.labor_bureau),
            'AG9': p1, 'AL9': p2,
            'AF10': company.address,
            'AF12': company.company_name,
            'AF13': company.representative_name,
            'AG16': sp1, 'AL16': sp2,
            'AF17': sr.address if sr else "",
            'AF19': sr.office_name if sr else "",
            'AF20': sr.sr_name if sr else "",
            'AF21': spp1, 'AM21': spp2, 'AT21': spp3,
            'K26': company.main_business,
            'K27': company.employee_count,
            # 雇用保険適用事業所名（offices[0].name 優先）
            'K30': get_insurance_office_name(company),
            'AN30': i1, 'AS30': i2, 'AZ30': i3,
            'R31': company.contact_name,
            'AM31': company.contact_department,
            'R32': cp1, 'W32': cp2, 'AB32': cp3,
            'AM32': company.contact_email,
            'M37': male,
            'AM37': female,
        }
        # 法人番号を1桁ずつ AF14〜AR14 (13桁)
        if company.corporate_number:
            digits = str(company.corporate_number).replace('-', '').strip()
            cols = ['AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR']
            for i, d in enumerate(digits[:13]):
                values[f'{cols[i]}14'] = d

        # 定額制の場合は契約途中解約禁止チェックボックスをチェック
        checkbox_states = get_4_2_teigaku_checkboxes(group.subsidy_course)

        fname = f"01_支給申請書_{group.curriculum_name}_{company.company_name}.xlsx"
        return self._patch(
            "支給申請/01_支給申請書(様式第4-2号).xlsx", fname, values,
            checkbox_states=checkbox_states,
        )

    def write_経費助成内訳(self, company: CompanyInfo, group: CurriculumGroup) -> str:
        """02_経費助成の内訳(様式第6-2号)"""
        values = {
            'J7': labor_bureau_short(company.labor_bureau),
            'AE7': get_insurance_office_name(company),
        }
        fname = f"02_経費助成の内訳_{group.curriculum_name}_{company.company_name}.xlsx"
        return self._patch("支給申請/02_経費助成の内訳(様式第6-2号).xlsx", fname, values)

    def write_経費助成内訳_定額制(self, company: CompanyInfo, group: CurriculumGroup) -> str:
        """04_定額制サービスによる訓練に関する経費助成の内訳(様式第6-3号) - 現行 R8.4 新規

        定額制ケース用。様式6-2号 の代わりに使用。
        セル位置（R070401・R80302 と共通）:
          - AG6: 訓練コース名
          - J7: 支給対象労働者数
          - M8/R8/W8: 訓練期間 始日（年/月/日）
          - AF8/AK8/AP8: 訓練期間 終日（年/月/日）
        """
        values = {
            'AG6': group.curriculum_name,
            'J7': len(group.participants),
        }
        if group.start_date:
            values['M8'] = group.start_date.year
            values['R8'] = group.start_date.month
            values['W8'] = group.start_date.day
        if group.end_date:
            values['AF8'] = group.end_date.year
            values['AK8'] = group.end_date.month
            values['AP8'] = group.end_date.day

        fname = f"04_様式6-3号(定額制)_{group.curriculum_name}_{company.company_name}.xlsx"
        return self._patch(
            "支給申請/04_定額制サービスによる訓練に関する経費助成の内訳(様式第6-3号).xlsx",
            fname, values,
        )

    def write_賃金助成内訳(self, company: CompanyInfo, group: CurriculumGroup) -> str:
        """03_賃金助成の内訳(様式第5号)

        修正点（PPT指摘事項）:
          - P7 (受付番号エリア) への labor_bureau 書き込みは誤りのため削除
          - BA7 は雇用保険適用事業所の名称 → offices[0].name 優先
          - 行22以降の支給対象労働者一覧にデータを書き込む
        """
        values = {
            # 計画届の受付番号 (P7) は外部から付与される番号なので空のまま
            # 雇用保険適用事業所の名称
            'BA7': get_insurance_office_name(company),
        }
        # 支給対象労働者一覧 (行22から)
        # A: No, B: 対象労働者名, K: フリガナ, T: 雇用保険被保険者番号
        start_row = 22
        for idx, p in enumerate(group.participants):
            r = start_row + idx
            values[f'B{r}'] = p.name
            values[f'K{r}'] = getattr(p, 'furigana', '')
            values[f'T{r}'] = (p.insurance_number or "").replace('−', '-').replace('ー', '-')

        fname = f"03_賃金助成の内訳_{group.curriculum_name}_{company.company_name}.xlsx"
        return self._patch("支給申請/03_賃金助成の内訳(様式第5号).xlsx", fname, values)

    def write_事業所確認票(self, company: CompanyInfo, group: CurriculumGroup,
                          submit_date: datetime) -> str:
        """05_事業所確認票(様式第13号)

        事業所欄の構造:
          申請事業所(本社) 行16: A=名, D=4桁, I=6桁, P=1桁, Q=労働者数
          従たる事業所1-10 行21,23,25,27,29,31,33,35,37,39: B/D/I/P/Q
        """
        values = {
            'N3': submit_date.year, 'R3': submit_date.month, 'T3': submit_date.day,
            'B4': labor_bureau_short(company.labor_bureau),
            'L8': company.company_name,
            'L10': company.address,
        }

        offices = company.offices or []

        # 本社(申請事業所) = offices[0] があればそれ、なければ CompanyInfo から構築
        if offices:
            head = offices[0]
            head_name = head.name or company.company_name
            head_ins = head.insurance_number or company.insurance_office_number
            head_emp = head.employee_count
        else:
            head_name = company.company_name
            head_ins = company.insurance_office_number
            head_emp = company.employee_count

        i1, i2, i3 = split_insurance_number(head_ins)
        values['A16'] = head_name
        values['D16'] = i1
        values['I16'] = i2
        values['P16'] = i3
        if head_emp:
            values['Q16'] = head_emp

        # 従たる事業所 (最大10件、行21から2行ずつ)
        subsidiary = offices[1:] if len(offices) > 1 else []
        office_rows = [21, 23, 25, 27, 29, 31, 33, 35, 37, 39]
        total_employee = head_emp or 0
        for idx, office in enumerate(subsidiary[:10]):
            r = office_rows[idx]
            s1, s2, s3 = split_insurance_number(office.insurance_number)
            values[f'B{r}'] = office.name
            values[f'D{r}'] = s1
            values[f'I{r}'] = s2
            values[f'P{r}'] = s3
            if office.employee_count:
                values[f'Q{r}'] = office.employee_count
                total_employee += office.employee_count

        # 事業所数 (L12) = 本社含む全事業所数
        total_offices = max(1, len(offices)) if offices else 1
        values['L12'] = total_offices

        # 申請事業所と申請事業所以外の常時雇用する労働者数の合計 (Q42)
        # company.employee_count があればそれを優先、なければ全事業所の合計
        values['Q42'] = company.employee_count if company.employee_count else total_employee

        fname = f"05_事業所確認票_{group.curriculum_name}_{company.company_name}.xlsx"
        return self._patch("支給申請/05_事業所確認票(様式第13号).xlsx", fname, values)

    def write_支給申請承諾書(self, company: CompanyInfo, group: CurriculumGroup,
                            submit_date: datetime) -> str:
        """41_支給申請承諾書(様式第12号)
        対象訓練欄 (1件目のみ): F40=コース名, S40/W40/Y40=開始 年/月/日,
                               S42/W42/Y42=終了 年/月/日
        申請事業主欄: E51=所在地, E53=名称, E55=氏名
        """
        values = {
            # 確認日
            'R24': submit_date.year, 'V24': submit_date.month, 'Y24': submit_date.day,
            # 対象訓練 (1件目)
            'F40': group.curriculum_name,
            # 申請事業主
            'E51': company.address,
            'E53': company.company_name,
            'E55': company.representative_name,
        }
        # 訓練実施期間
        if group.start_date:
            values['S40'] = group.start_date.year
            values['W40'] = group.start_date.month
            values['Y40'] = group.start_date.day
        if group.end_date:
            values['S42'] = group.end_date.year
            values['W42'] = group.end_date.month
            values['Y42'] = group.end_date.day

        fname = f"41_支給申請承諾書_{group.curriculum_name}_{company.company_name}.xlsx"
        return self._patch("支給申請/41_支給申請承諾書(様式第12号).xlsx", fname, values)

    # ==================================================================
    # 書類一式生成
    # ==================================================================

    def generate_plan_documents(self, company: CompanyInfo, sr: SocialInsuranceLabor,
                                 group: CurriculumGroup, submit_date: datetime) -> List[str]:
        generated = []
        teigaku = is_teigaku_course(group.subsidy_course)

        jobs = [
            ("計画届", lambda: self.write_訓練実施計画届(company, sr, group, submit_date)),
            ("事業展開等実施計画", lambda: self.write_事業展開等実施計画(company, group, submit_date)),
        ]
        if teigaku:
            jobs.append(("対象者一覧(3-2定額制)", lambda: self.write_対象者一覧_3_2(company, group)))
        else:
            jobs.append(("対象者一覧(3-1)", lambda: self.write_対象者一覧_3_1(company, group)))
        # 様式第3号 (※電子申請のみ) - 常に追加生成
        jobs.append(("対象者一覧(電子申請)",
                     lambda: self.write_対象者一覧_電子申請(company, group)))
        jobs.append(("事前確認書",
                     lambda: self.write_事前確認書(company, sr, group, submit_date, teigaku=teigaku)))

        for label, fn in jobs:
            try:
                generated.append(fn())
            except Exception as e:
                print(f"[{label}] 生成失敗: {e}")
                import traceback
                traceback.print_exc()
        return generated

    def generate_payment_documents(self, company: CompanyInfo, sr: SocialInsuranceLabor,
                                    group: CurriculumGroup, submit_date: datetime) -> List[str]:
        generated = []
        teigaku = is_teigaku_course(group.subsidy_course)

        jobs = [
            ("支給申請書", lambda: self.write_支給申請書(company, sr, group, submit_date)),
        ]
        # 経費助成内訳: 定額制は 様式6-3号、通常は 様式6-2号
        if teigaku:
            jobs.append(("経費助成内訳(6-3号定額制)",
                         lambda: self.write_経費助成内訳_定額制(company, group)))
        else:
            jobs.append(("経費助成内訳(6-2号)",
                         lambda: self.write_経費助成内訳(company, group)))
        jobs.extend([
            ("賃金助成内訳", lambda: self.write_賃金助成内訳(company, group)),
            ("事業所確認票", lambda: self.write_事業所確認票(company, group, submit_date)),
            ("支給申請承諾書", lambda: self.write_支給申請承諾書(company, group, submit_date)),
        ])
        for label, fn in jobs:
            try:
                generated.append(fn())
            except Exception as e:
                print(f"[{label}] 生成失敗: {e}")
                import traceback
                traceback.print_exc()
        return generated
