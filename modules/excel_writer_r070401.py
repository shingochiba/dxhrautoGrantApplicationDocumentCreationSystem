"""
R070401（令和7年4月1日）書式用 Excel Writer

ExcelWriter のサブクラスとして R070401 書式特有の差分を実装する。
- 様式1-1号: ヘッダー同じ、訓練コース系 -1 行、男女別 -7 行
- 様式1-3号: レイアウト大幅変更（form 短縮）
- 様式3-1号: 完全新構造（マルチシート、データ20行目以降、4列構成）
- 様式3-2号: 同じ位置（テンプレートパスのみ変更）
- 様式11号: ヘッダー部は同じ位置
- 様式4-2号: row 26-27 同じ、雇用保険番号・担当者・男女別 -1 行
- 様式6-2号: row 7 同じ
- 様式5号: 新構造（company name は BA11）
- 様式13号: 同じ位置
- 様式12号 / 様式14-1号 は R070401 には無いのでスキップ
"""
from datetime import datetime
from typing import List

from .company_form import CompanyInfo, SocialInsuranceLabor
from .upload_handler import CurriculumGroup
from .excel_writer import (
    ExcelWriter, is_teigaku_course,
    split_postal, split_phone, split_insurance_number,
    labor_bureau_short, get_insurance_office_name,
    get_training_method_checkboxes,
    get_digital_training_checkboxes,
    get_4_2_teigaku_checkboxes,
    chunk_participants, chunk_suffix, _PARTICIPANTS_PER_FORM,
)


def _normalize_insurance(num: str) -> str:
    """雇用保険被保険者番号からハイフンを除去して 11桁文字列にする"""
    if not num:
        return ""
    return (num.replace('-', '')
            .replace('−', '')
            .replace('ー', '')
            .replace('―', '')
            .strip())


class ExcelWriterR070401(ExcelWriter):
    """R070401 書式用 Writer"""

    # ==================================================================
    # 計画申請書類
    # ==================================================================

    def write_訓練実施計画届(self, company: CompanyInfo, sr: SocialInsuranceLabor,
                              group: CurriculumGroup, submit_date: datetime,
                              training_company=None) -> str:
        """01_職業訓練実施計画届(様式第1-1号) - R070401版

        現行との差分:
          - ヘッダー部 (rows 1-30) は同じ
          - 訓練コース・期間 (row 41-42) は -1 行
          - 男女別受講者数 (row 73) は -7 行
        """
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

        # 受講場所
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
            'AL5': submit_date.year, 'AR5': submit_date.month, 'AU5': "",
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
            # 雇用保険適用事業所
            'K26': insurance_office_name,
            'AN26': i1, 'AS26': i2, 'AZ26': i3,
            'M27': p1, 'Q27': p2,
            'K28': company.address,
            # 担当者
            'R29': company.contact_name,
            'AM29': company.contact_department,
            'R30': cp1, 'W30': cp2, 'AB30': cp3,
            'AM30': company.contact_email,
            # 訓練コース・受講者数 (R070401 では -1 行)
            'K41': group.curriculum_name,
            'AN41': len(group.participants),
            # 訓練の実施場所 (-1 行: row 46)
            'K46': location_value,
            # 男女別受講者数 (R070401 では行 73)
            'S73': male, 'AN73': female,
        }
        # 訓練期間 (R070401 では -1 行)
        if group.start_date:
            values['N42'] = group.start_date.year
            values['T42'] = group.start_date.month
            values['Z42'] = group.start_date.day
        if group.end_date:
            values['AI42'] = group.end_date.year
            values['AO42'] = group.end_date.month
            values['AU42'] = group.end_date.day

        # 定額制の場合: 9番「契約期間」= 8番「実施期間」と同じ (-1 行: row 43)
        if teigaku:
            if group.start_date:
                values['N43'] = group.start_date.year
                values['T43'] = group.start_date.month
                values['Z43'] = group.start_date.day
            if group.end_date:
                values['AI43'] = group.end_date.year
                values['AO43'] = group.end_date.month
                values['AU43'] = group.end_date.day
            # 標準学習時間プレースホルダ空欄化（R070401 では行 54）
            values['R54'] = ""
            values['Y54'] = ""

        # 12 訓練の実施方法のチェックボックスを助成コースに応じて設定
        checkbox_states = get_training_method_checkboxes(group.subsidy_course)
        # 17 デジタル人材の育成 (R070401 は ctrlProp26〜31)
        checkbox_states.update(get_digital_training_checkboxes(group.curriculum_name, base_ctrl=26))

        # 16欄 教育訓練機関 (R070401: 名称 R61, 代表者 AM61, 所在地 R62)
        if training_company is not None:
            rep_full = f"{training_company.representative_title}　{training_company.representative_name}".strip()
            values['R61'] = training_company.name
            values['AM61'] = rep_full
            values['R62'] = training_company.address

        fname = f"01_職業訓練実施計画届_{company.company_name}_{group.curriculum_name}.xlsx"
        return self._patch(
            "計画申請/01_職業訓練実施計画届(様式第1-1号)_企業名.xlsx",
            fname, values,
            checkbox_states=checkbox_states,
        )

    def write_事業展開等実施計画(self, company: CompanyInfo, group: CurriculumGroup,
                                  submit_date: datetime) -> str:
        """02_事業展開等実施計画(様式第1-3号) - R070401版

        現行との差分: form 大幅短縮、申請事業主の証明欄が row 37-41 に移動
        """
        values = {
            # 申請事業主の証明欄の日付
            'H37': submit_date.year,
            'L37': submit_date.month,
            'O37': "",  # 日は空欄
            # 代表者役職名・氏名
            'K40': company.representative_title,
            'K41': company.representative_name,
        }
        fname = f"02_事業展開等実施計画_{company.company_name}_{group.curriculum_name}.xlsx"
        return self._patch(
            "計画申請/02_事業展開等実施計画(様式第1-3号)_企業名.xlsx",
            fname, values
        )

    def _write_対象労働者一覧_新形式(self, group: CurriculumGroup,
                                       template_rel: str, fname: str) -> str:
        """様式3-1号 R070401 の新レイアウト用 共通実装

        新形式の特徴:
          - データ行は 20 行目から（17行目以下はヘッダー＋例）
          - 1 行 = 1 受講者（現行のような 2 行 1 セット ではない）
          - A: No, B: 氏名, C: 雇用保険番号(11桁文字列), D: 雇用形態
          - E〜G は条件付き項目のため空欄（手動入力）
        """
        values = {}
        start_row = 20
        for idx, p in enumerate(group.participants):
            r = start_row + idx
            values[f'A{r}'] = idx + 1
            values[f'B{r}'] = p.name
            values[f'C{r}'] = _normalize_insurance(p.insurance_number)
            values[f'D{r}'] = p.employment_type
        return self._patch(template_rel, fname, values)

    def write_対象者一覧_3_1(self, company: CompanyInfo, group: CurriculumGroup) -> str:
        """03_対象労働者一覧(様式第3-1号) - R070401版（新レイアウト）"""
        fname = f"03_対象労働者一覧(3-1)_{company.company_name}_{group.curriculum_name}.xlsx"
        return self._write_対象労働者一覧_新形式(
            group,
            "計画申請/03_対象労働者一覧_企業名(様式第3-1号).xlsx",
            fname
        )

    def write_対象者一覧_3_2(self, company: CompanyInfo, group: CurriculumGroup,
                              submit_date: datetime = None) -> List[str]:
        """03_対象労働者一覧(様式第3-2号)定額制 - R070401版

        列構造: A=番号(自動), B=氏名, C-F(merged)=正規雇用マーク, G-J(merged)=有期契約マーク
        41人を超える場合は書式を分割し、ファイル名に ①②... を付与。
        """
        if submit_date is None:
            from datetime import datetime as _dt
            submit_date = _dt.now()
        rep = f"{company.representative_title}　{company.representative_name}".strip()
        cert_text = f"{company.address}\n{rep}"

        chunks = chunk_participants(group.participants, _PARTICIPANTS_PER_FORM)
        output_paths: List[str] = []
        for chunk_idx, chunk in enumerate(chunks):
            offset = chunk_idx * _PARTICIPANTS_PER_FORM
            values = {
                'E8': submit_date.year, 'G8': submit_date.month, 'I8': "",
                'E9': cert_text,
                'C11': company.company_name,
                'C12': group.curriculum_name,
                'C50': company.company_name,
                'C51': group.curriculum_name,
            }
            # A列「No.」を 40 スロット全てに連番で上書き
            for slot in range(_PARTICIPANTS_PER_FORM):
                r = (16 + slot) if slot < 20 else (55 + (slot - 20))
                values[f'A{r}'] = offset + slot + 1
            for idx, p in enumerate(chunk):
                r = (16 + idx) if idx < 20 else (55 + (idx - 20))
                values[f'B{r}'] = p.name
                if '正規' in p.employment_type:
                    values[f'C{r}'] = '○'
                elif '有期' in p.employment_type:
                    values[f'G{r}'] = '○'

            suffix = chunk_suffix(chunk_idx, len(chunks))
            suffix_part = f"_{suffix}" if suffix else ""
            fname = (f"03_対象労働者一覧(3-2){suffix_part}_"
                     f"{company.company_name}_{group.curriculum_name}.xlsx")
            path = self._patch(
                "計画申請/03_人材開発支援助成金（事業展開等リスキリング支援コース）(様式第3-2号)※定額制.xlsx",
                fname, values
            )
            output_paths.append(path)
        return output_paths

    def write_様式14_1(self, company: CompanyInfo, group: CurriculumGroup,
                       submit_date: datetime) -> str:
        """05_様式第14-1号 ※定額制 定額制サービスによる訓練に関する事業所確認票（R070401 新規）

        セル位置は R80302 の様式14-1号 と同じ。
        """
        offices = company.offices or []
        if offices:
            head = offices[0]
            head_name = head.name or company.company_name
            head_ins = head.insurance_number or company.insurance_office_number
            head_emp = head.employee_count or 0
        else:
            head_name = company.company_name
            head_ins = company.insurance_office_number
            head_emp = 0

        i1, i2, i3 = split_insurance_number(head_ins)

        # マージセル分析に基づくデータセル位置
        values = {
            # 提出日 (R3:S3 merged=年, U3=月, W3=日)
            'R3': submit_date.year,
            'U3': submit_date.month,
            'W3': "",  # 日は空欄
            # 事業主名 (M8:X8 merged)
            'M8': company.company_name,
            # 所在地 (M10:X10 merged)
            'M10': company.address,
            # 訓練コース名 (F12:P12 merged)
            'F12': group.curriculum_name,
            # 申請事業所（本社）行16-17
            'A16': head_name,
            'G16': i1, 'L16': i2, 'S16': i3,
            'T16': head_emp if head_emp else "",
        }

        subsidiary = offices[1:] if len(offices) > 1 else []
        office_rows = [21, 23, 25, 27, 29, 31, 33, 35, 37, 39]
        for idx, office in enumerate(subsidiary[:10]):
            r = office_rows[idx]
            s1, s2, s3 = split_insurance_number(office.insurance_number)
            values[f'B{r}'] = office.name
            values[f'G{r}'] = s1
            values[f'L{r}'] = s2
            values[f'S{r}'] = s3
            if office.employee_count:
                values[f'T{r}'] = office.employee_count

        # 事業所数 (U12:V12 merged)
        total_offices = max(1, len(offices)) if offices else 1
        values['U12'] = total_offices

        fname = f"05_様式14-1号_{company.company_name}_{group.curriculum_name}.xlsx"
        return self._patch(
            "計画申請/05_定額制サービスによる訓練に関する事業所確認票(様式第14-1号)※定額制.xlsx",
            fname, values
        )

    def write_事前確認書(self, company: CompanyInfo, sr: SocialInsuranceLabor,
                         group: CurriculumGroup, submit_date: datetime,
                         teigaku: bool = False) -> str:
        """04_事前確認書(様式第11号) - R070401版

        R070401 には定額制版テンプレートが無いため、teigaku に関わらず
        同じテンプレートを使用する。ヘッダー部（rows 12-30）は現行と同じ位置。
        """
        p1, p2 = split_postal(company.postal_code)
        sp1, sp2 = split_postal(sr.postal_code) if sr else ("", "")
        cp1, cp2, cp3 = split_phone(company.phone_number)
        spp1, spp2, spp3 = split_phone(sr.phone_number) if sr else ("", "", "")

        values = {
            'E12': submit_date.year, 'I12': submit_date.month, 'L12': "",
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
        suffix = "_定額制" if teigaku else ""
        fname = f"04_事前確認書{suffix}_{company.company_name}_{group.curriculum_name}.xlsx"
        return self._patch(
            "計画申請/04_事前確認書(様式第11号)_企業名.xlsx",
            fname, values
        )

    # ==================================================================
    # 支給申請書類
    # ==================================================================

    def write_支給申請書(self, company: CompanyInfo, sr: SocialInsuranceLabor,
                         group: CurriculumGroup, submit_date: datetime) -> str:
        """01_支給申請書(様式第4-2号) - R070401版

        現行との差分:
          - ヘッダー・社労士・主たる事業・雇用人数 (rows 1-27) は同じ位置
          - 雇用保険適用事業所 (row 30 → 29) は -1 行
          - 担当者 (row 31-32 → 30-31) は -1 行
          - 男女別 (row 37 → 36) は -1 行
        """
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
            'AL5': submit_date.year, 'AR5': submit_date.month, 'AU5': "",
            'B7': labor_bureau_short(company.labor_bureau),
            'AG9': p1, 'AL9': p2,
            'AF10': company.address,
            'AF12': company.company_name,
            'AF13': company.representative_name,
            # 社労士
            'AG16': sp1, 'AL16': sp2,
            'AF17': sr.address if sr else "",
            'AF19': sr.office_name if sr else "",
            'AF20': sr.sr_name if sr else "",
            'AF21': spp1, 'AM21': spp2, 'AT21': spp3,
            # 主たる事業・雇用人数 (row 26-27 - 同じ位置)
            'K26': company.main_business,
            'K27': company.employee_count,
            # 雇用保険適用事業所名 (row 29 - R070401 では -1 行) - offices[0].name 優先
            'K29': get_insurance_office_name(company),
            'AN29': i1, 'AS29': i2, 'AZ29': i3,
            # 担当者 (row 30-31 - R070401 では -1 行)
            'R30': company.contact_name,
            'AM30': company.contact_department,
            'R31': cp1, 'W31': cp2, 'AB31': cp3,
            'AM31': company.contact_email,
            # 男女別 (row 36 - R070401 では -1 行)
            'M36': male,
            'AM36': female,
        }
        # 法人番号 1桁ずつ AF14〜AR14 (13桁) - row 14 は同じ
        if company.corporate_number:
            digits = str(company.corporate_number).replace('-', '').strip()
            cols = ['AF','AG','AH','AI','AJ','AK','AL','AM','AN','AO','AP','AQ','AR']
            for i, d in enumerate(digits[:13]):
                values[f'{cols[i]}14'] = d

        # 定額制の場合は契約途中解約禁止チェックボックスをチェック
        checkbox_states = get_4_2_teigaku_checkboxes(group.subsidy_course)

        fname = f"01_支給申請書_{group.curriculum_name}_{company.company_name}.xlsx"
        return self._patch(
            "支給申請/01_支給申請書(様式第4-2号)_研修名_企業名.xlsx",
            fname, values,
            checkbox_states=checkbox_states,
        )

    def write_経費助成内訳(self, company: CompanyInfo, group: CurriculumGroup) -> str:
        """02_経費助成の内訳(様式第6-2号) - R070401版（row 7 は同じ位置）"""
        values = {
            'J7': labor_bureau_short(company.labor_bureau),
            'AE7': get_insurance_office_name(company),
        }
        fname = f"02_経費助成の内訳_{group.curriculum_name}_{company.company_name}.xlsx"
        return self._patch(
            "支給申請/02_経費助成の内訳(様式第6-2号)_研修名_企業名.xlsx",
            fname, values
        )

    def write_経費助成内訳_定額制(self, company: CompanyInfo, group: CurriculumGroup) -> str:
        """06_様式第6-3号 ※定額制 定額制サービスによる訓練に関する経費助成の内訳（R070401）

        定額制ケース用。自動入力は行わず、未入力の書式 (空欄) として出力する。
        テンプレートに残ったサンプル値も空欄にクリアして提供する。
        ファイル名に「未生成」を含めることでユーザーに手入力を促す。
        """
        values = {
            'AG6': '',
            'J7': '', 'AG7': '',
            'M8': '', 'R8': '', 'W8': '',
            'AF8': '', 'AK8': '', 'AP8': '',
            'M9': '', 'R9': '', 'W9': '',
            'AF9': '', 'AK9': '', 'AP9': '',
            'AB13': '',
        }
        fname = f"06_様式6-3号_未生成_{group.curriculum_name}_{company.company_name}.xlsx"
        return self._patch(
            "支給申請/06_定額制サービスによる訓練に関する経費助成の内訳(様式第6-3号)※定額制.xlsx",
            fname, values,
        )

    def write_対象者一覧_支給(self, company: CompanyInfo, group: CurriculumGroup) -> str:
        """03_対象労働者一覧 - R070401版（計画申請の3-1号と同じ新レイアウト）"""
        fname = f"03_対象労働者一覧(支給)_{group.curriculum_name}_{company.company_name}.xlsx"
        return self._write_対象労働者一覧_新形式(
            group,
            "支給申請/03_対象労働者一覧_企業名.xlsx",
            fname
        )

    def write_賃金助成内訳(self, company: CompanyInfo, group: CurriculumGroup) -> str:
        """03_賃金助成の内訳(様式第5号) - R070401版

        新レイアウト: 雇用保険適用事業所の名称は row 11 に移動。
        labor_bureau は本様式に存在しない。
        """
        values = {
            'BA11': get_insurance_office_name(company),
        }
        fname = f"03_賃金助成の内訳_{group.curriculum_name}_{company.company_name}.xlsx"
        return self._patch(
            "支給申請/03_賃金助成の内訳(様式第5号)_研修名_企業名.xlsx",
            fname, values
        )

    def write_事業所確認票(self, company: CompanyInfo, group: CurriculumGroup,
                          submit_date: datetime) -> str:
        """05_事業所確認票(様式第13号) - R070401版（位置は現行と同じ）"""
        values = {
            'N3': submit_date.year, 'R3': submit_date.month, 'T3': "",
            'B4': labor_bureau_short(company.labor_bureau),
            'L8': company.company_name,
            'L10': company.address,
        }
        offices = company.offices or []
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

        total_offices = max(1, len(offices)) if offices else 1
        values['L12'] = total_offices

        # 常時雇用労働者数の合計 (Q42)
        values['Q42'] = company.employee_count if company.employee_count else total_employee

        fname = f"05_事業所確認票_{group.curriculum_name}_{company.company_name}.xlsx"
        return self._patch(
            "支給申請/05_事業所確認票(様式第13号)_研修名_企業名.xlsx",
            fname, values
        )

    # ==================================================================
    # 書類一式生成
    # ==================================================================

    def generate_plan_documents(self, company: CompanyInfo, sr: SocialInsuranceLabor,
                                 group: CurriculumGroup, submit_date: datetime,
                                 training_company=None) -> List[str]:
        """計画申請書類の一括生成（R070401）

        現行との差分:
          - Word書類（事業内職業能力開発計画）は無し
          - 定額制の場合のみ 様式14-1号 (※定額制) を追加
        """
        generated = []
        teigaku = is_teigaku_course(group.subsidy_course)

        jobs = [
            ("計画届", lambda: self.write_訓練実施計画届(company, sr, group, submit_date,
                                                       training_company=training_company)),
            ("事業展開等実施計画", lambda: self.write_事業展開等実施計画(company, group, submit_date)),
        ]
        if teigaku:
            jobs.append(("対象者一覧(3-2定額制)", lambda: self.write_対象者一覧_3_2(company, group, submit_date)))
        else:
            jobs.append(("対象者一覧(3-1)", lambda: self.write_対象者一覧_3_1(company, group)))
        jobs.append(("事前確認書",
                     lambda: self.write_事前確認書(company, sr, group, submit_date, teigaku=teigaku)))
        if teigaku:
            jobs.append(("様式14-1号", lambda: self.write_様式14_1(company, group, submit_date)))

        for label, fn in jobs:
            try:
                result = fn()
                if isinstance(result, list):
                    generated.extend(result)
                else:
                    generated.append(result)
            except Exception as e:
                print(f"[{label}] 生成失敗: {e}")
                import traceback
                traceback.print_exc()
        return generated

    def generate_payment_documents(self, company: CompanyInfo, sr: SocialInsuranceLabor,
                                    group: CurriculumGroup, submit_date: datetime,
                                    training_company=None) -> List[str]:
        """支給申請書類の一括生成（R070401）

        現行との差分:
          - 支給申請承諾書(様式12号) は R070401 には無いのでスキップ
          - 対象者一覧(様式3-1号 新形式) を支給申請にも含める
          - 定額制の場合は 様式6-3号 (※定額制) を使用、通常は 様式6-2号 を使用
        """
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
            ("対象者一覧(支給)", lambda: self.write_対象者一覧_支給(company, group)),
            ("賃金助成内訳", lambda: self.write_賃金助成内訳(company, group)),
            ("事業所確認票", lambda: self.write_事業所確認票(company, group, submit_date)),
        ])
        for label, fn in jobs:
            try:
                result = fn()
                if isinstance(result, list):
                    generated.extend(result)
                else:
                    generated.append(result)
            except Exception as e:
                print(f"[{label}] 生成失敗: {e}")
                import traceback
                traceback.print_exc()
        return generated
