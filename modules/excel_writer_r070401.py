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
                              group: CurriculumGroup, submit_date: datetime) -> str:
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

        values = {
            # 提出日
            'AL5': submit_date.year, 'AR5': submit_date.month, 'AU5': submit_date.day,
            # 管轄労働局
            'B7': company.labor_bureau,
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
            'K26': company.company_name,
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

        fname = f"01_職業訓練実施計画届_{company.company_name}_{group.curriculum_name}.xlsx"
        return self._patch(
            "計画申請/01_職業訓練実施計画届(様式第1-1号)_企業名.xlsx",
            fname, values
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
            'O37': submit_date.day,
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

    def write_対象者一覧_3_2(self, company: CompanyInfo, group: CurriculumGroup) -> str:
        """03_対象労働者一覧(様式第3-2号)定額制 - R070401版（位置は現行と同じ）"""
        values = {
            'C11': company.company_name,
            'C12': group.curriculum_name,
        }
        start_row = 16
        for idx, p in enumerate(group.participants):
            r = start_row + idx
            values[f'B{r}'] = idx + 1
            values[f'C{r}'] = p.name
            values[f'G{r}'] = p.insurance_number

        fname = f"03_対象労働者一覧(3-2)_{company.company_name}_{group.curriculum_name}.xlsx"
        return self._patch(
            "計画申請/03_人材開発支援助成金（事業展開等リスキリング支援コース）(様式第3-2号)※定額制.xlsx",
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
            'A30': company.labor_bureau,
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
            'AL5': submit_date.year, 'AR5': submit_date.month, 'AU5': submit_date.day,
            'B7': company.labor_bureau,
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
            # 雇用保険適用事業所 (row 29 - R070401 では -1 行)
            'K29': company.company_name,
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

        fname = f"01_支給申請書_{group.curriculum_name}_{company.company_name}.xlsx"
        return self._patch(
            "支給申請/01_支給申請書(様式第4-2号)_研修名_企業名.xlsx",
            fname, values
        )

    def write_経費助成内訳(self, company: CompanyInfo, group: CurriculumGroup) -> str:
        """02_経費助成の内訳(様式第6-2号) - R070401版（row 7 は同じ位置）"""
        values = {
            'J7': company.labor_bureau,
            'AE7': company.company_name,
        }
        fname = f"02_経費助成の内訳_{group.curriculum_name}_{company.company_name}.xlsx"
        return self._patch(
            "支給申請/02_経費助成の内訳(様式第6-2号)_研修名_企業名.xlsx",
            fname, values
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
            'BA11': company.company_name,
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
            'N3': submit_date.year, 'R3': submit_date.month, 'T3': submit_date.day,
            'B4': company.labor_bureau,
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
        for idx, office in enumerate(subsidiary[:10]):
            r = office_rows[idx]
            s1, s2, s3 = split_insurance_number(office.insurance_number)
            values[f'B{r}'] = office.name
            values[f'D{r}'] = s1
            values[f'I{r}'] = s2
            values[f'P{r}'] = s3
            if office.employee_count:
                values[f'Q{r}'] = office.employee_count

        total_offices = max(1, len(offices)) if offices else 1
        values['L12'] = total_offices

        fname = f"05_事業所確認票_{group.curriculum_name}_{company.company_name}.xlsx"
        return self._patch(
            "支給申請/05_事業所確認票(様式第13号)_研修名_企業名.xlsx",
            fname, values
        )

    # ==================================================================
    # 書類一式生成
    # ==================================================================

    def generate_plan_documents(self, company: CompanyInfo, sr: SocialInsuranceLabor,
                                 group: CurriculumGroup, submit_date: datetime) -> List[str]:
        """計画申請書類の一括生成（R070401）

        現行との差分:
          - Word書類（事業内職業能力開発計画）は無し
          - 様式14-1号 は無し
        """
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
        """支給申請書類の一括生成（R070401）

        現行との差分:
          - 支給申請承諾書(様式12号) は R070401 には無いのでスキップ
          - 対象者一覧(様式3-1号 新形式) を支給申請にも含める
        """
        generated = []
        jobs = [
            ("支給申請書", lambda: self.write_支給申請書(company, sr, group, submit_date)),
            ("経費助成内訳", lambda: self.write_経費助成内訳(company, group)),
            ("対象者一覧(支給)", lambda: self.write_対象者一覧_支給(company, group)),
            ("賃金助成内訳", lambda: self.write_賃金助成内訳(company, group)),
            ("事業所確認票", lambda: self.write_事業所確認票(company, group, submit_date)),
        ]
        for label, fn in jobs:
            try:
                generated.append(fn())
            except Exception as e:
                print(f"[{label}] 生成失敗: {e}")
                import traceback
                traceback.print_exc()
        return generated
