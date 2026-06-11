"""
R80302（令和8年3月2日）書式用 Excel Writer

ExcelWriter のサブクラスとして、R80302 書式特有の差分のみを実装する。
- 大半の様式は現行と同じセル位置 → テンプレートパスのみ変更（オーバーライド）
- 様式1-1号: 男女別受講者数の行が 80 → 74 へ移動
- 様式1-3号: 全セルが +1 行ずれている
- 様式14-1号: R80302 新規（定額制サービスによる訓練の事業所確認票）
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
    get_3_1_employment_checkboxes,
    get_4_2_teigaku_checkboxes,
    chunk_participants, chunk_suffix, _PARTICIPANTS_PER_FORM,
)


class ExcelWriterR80302(ExcelWriter):
    """R80302 書式用 Writer"""

    # ==================================================================
    # 計画申請書類
    # ==================================================================

    def write_訓練実施計画届(self, company: CompanyInfo, sr: SocialInsuranceLabor,
                              group: CurriculumGroup, submit_date: datetime,
                              training_company=None) -> str:
        """01_職業訓練実施計画届(様式第1-1号) - R80302版

        現行との差分: 男女別受講者数が S80/AN80 → S74/AN74 に移動
        16欄 教育訓練機関: 名称 R62, 代表者 AM62, 所在地 R63 (R070401 から +1 行ずれ)
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

        # 受講場所: 1人目の受講者情報から組み立て
        first_p = group.participants[0] if group.participants else None
        location_value = ""
        if first_p:
            if getattr(first_p, 'location_home', '').strip():
                location_value = first_p.location_home.strip()
            elif getattr(first_p, 'location_name', '').strip():
                location_value = first_p.location_name.strip()
        if not location_value:
            location_value = f"本社({company.address})"
        # K47 (上段=受講場所) には「受信先：」プレフィックスを付与
        location_value = "受信先：" + location_value
        # GAID 選択時: 下段「送信元」セル (K48) を上書きする
        gaid_sender_value = None
        if training_company is not None and 'GAID' in training_company.name:
            gaid_sender_value = "送信元：" + training_company.address

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
            # 訓練コース
            'K42': group.curriculum_name,
            'AN42': len(group.participants),
            # 訓練の実施場所
            'K47': location_value,
            # 男女別受講者数（R80302 では行 74 に移動）
            'S74': male, 'AN74': female,
        }
        # GAID 選択時のみ K48 (送信元セル) を上書き
        if gaid_sender_value is not None:
            values['K48'] = gaid_sender_value
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
            # 13番「標準学習時間」プレースホルダ空欄化（R80302 でも -1 行）
            values['R54'] = ""
            values['Y54'] = ""

        # 12 訓練の実施方法のチェックボックスを助成コースに応じて設定
        checkbox_states = get_training_method_checkboxes(group.subsidy_course)
        # 17 デジタル人材の育成 (R80302 は ctrlProp24〜29、現行と同じ)
        checkbox_states.update(get_digital_training_checkboxes(group.curriculum_name, base_ctrl=24))

        # 16欄 教育訓練機関 (R80302: 名称 R62, 代表者 AM62, 所在地 R63)
        if training_company is not None:
            rep_full = f"{training_company.representative_title}　{training_company.representative_name}".strip()
            values['R62'] = training_company.name
            values['AM62'] = rep_full
            values['R63'] = training_company.address

        fname = f"01_職業訓練実施計画届_{company.company_name}_{group.curriculum_name}.xlsx"
        return self._patch(
            "計画申請/01_職業訓練実施計画届(様式第1-1号)_企業名.xlsx",
            fname, values,
            checkbox_states=checkbox_states,
        )

    def write_事業展開等実施計画(self, company: CompanyInfo, group: CurriculumGroup,
                                  submit_date: datetime) -> str:
        """02_事業展開等実施計画(様式第1-3号) - R80302版

        現行との差分: 全セルが +1 行ずれている
        """
        values = {
            # 申請事業主の証明欄の日付（行73 → 行74）
            'H74': submit_date.year,
            'L74': submit_date.month,
            'O74': "",  # 日は空欄
            # 代表者役職名・氏名（行76,77 → 行77,78）
            'K77': company.representative_title,
            'K78': company.representative_name,
        }
        fname = f"02_事業展開等実施計画_{company.company_name}_{group.curriculum_name}.xlsx"
        return self._patch(
            "計画申請/02_事業展開等実施計画(様式第1-3号)_企業名.xlsx",
            fname, values
        )

    def write_対象者一覧_3_1(self, company: CompanyInfo, group: CurriculumGroup) -> List[str]:
        """03_対象労働者一覧(様式第3-1号) - R80302版（セル位置は現行と同じ）

        41人を超える場合は書式を分割し、ファイル名に ①②... を付与。
        """
        chunks = chunk_participants(group.participants, _PARTICIPANTS_PER_FORM)
        output_paths: List[str] = []
        for chunk_idx, chunk in enumerate(chunks):
            offset = chunk_idx * _PARTICIPANTS_PER_FORM
            values = {
                'C8': company.company_name,
                'C9': group.curriculum_name,
                'C63': company.company_name,
                'C64': group.curriculum_name,
            }
            # A列「No.」を 40 スロット全てに連番で上書き
            for slot in range(_PARTICIPANTS_PER_FORM):
                r = 13 + slot * 2 if slot < 20 else 68 + (slot - 20) * 2
                values[f'A{r}'] = offset + slot + 1
            for idx, p in enumerate(chunk):
                r = 13 + idx * 2 if idx < 20 else 68 + (idx - 20) * 2
                values[f'B{r}'] = p.name
                parts = (p.insurance_number or "").replace('−', '-').replace('ー', '-').split('-')
                if len(parts) == 3:
                    values[f'C{r}'] = parts[0]
                    values[f'F{r}'] = parts[1]
                    values[f'I{r}'] = parts[2]

            checkbox_states = get_3_1_employment_checkboxes(chunk)

            suffix = chunk_suffix(chunk_idx, len(chunks))
            suffix_part = f"_{suffix}" if suffix else ""
            fname = (f"03_対象労働者一覧(3-1){suffix_part}_"
                     f"{company.company_name}_{group.curriculum_name}.xlsx")
            path = self._patch(
                "計画申請/03_人材開発支援助成金（事業展開等リスキリング支援コース）対象労働者一覧(様式第3-1号).xlsx",
                fname, values,
                checkbox_states=checkbox_states,
            )
            output_paths.append(path)
        return output_paths

    def write_対象者一覧_3_2(self, company: CompanyInfo, group: CurriculumGroup,
                              submit_date: datetime = None) -> List[str]:
        """03_対象労働者一覧(様式第3-2号)定額制 - R80302版

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
                "計画申請/03_人材開発支援助成金（事業展開等リスキリング支援コース）対象労働者一覧(様式第3-2号)※定額制.xlsx",
                fname, values
            )
            output_paths.append(path)
        return output_paths

    def write_事前確認書(self, company: CompanyInfo, sr: SocialInsuranceLabor,
                         group: CurriculumGroup, submit_date: datetime,
                         teigaku: bool = False) -> str:
        """04_事前確認書(様式第11号) - R80302版（セル位置は現行と同じ）"""
        if teigaku:
            template = "計画申請/04_事前確認書(様式第11号)_企業名 ※定額制.xlsx"
            suffix = "_定額制"
        else:
            template = "計画申請/04_事前確認書(様式第11号)_企業名.xlsx"
            suffix = ""

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
        fname = f"04_事前確認書{suffix}_{company.company_name}_{group.curriculum_name}.xlsx"
        return self._patch(template, fname, values)

    def write_様式14_1(self, company: CompanyInfo, group: CurriculumGroup,
                       submit_date: datetime) -> str:
        """06_様式第14-1号 定額制サービスによる訓練に関する事業所確認票（R80302 新規）

        会社情報の事業所一覧（offices）を反映する：
          - 申請事業所(本社) = offices[0] または会社情報から構築 → 行16
          - 従たる事業所 = offices[1:] → 行21,23,25,...,39（最大10件）
        """
        offices = company.offices or []

        # 申請事業所（本社）
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

        # マージセル分析に基づくデータセル位置（O3/T3/V3 等はラベルセルなのでNG）
        values = {
            # 提出日 (R3:S3 merged=年, U3=月, W3=日)
            'R3': submit_date.year,
            'U3': submit_date.month,
            'W3': "",  # 日は空欄
            # 事業主名 (M8:X8 merged - J8は「事業主：」ラベル)
            'M8': company.company_name,
            # 所在地 (M10:X10 merged - J10は「所在地：」ラベル)
            'M10': company.address,
            # 訓練コース名 (F12:P12 merged - A12:E12は「訓練コースの名称」ラベル)
            'F12': group.curriculum_name,
            # 申請事業所（本社）行16-17 (B16:F17, G16:J17 などmerged)
            'A16': head_name,
            'G16': i1, 'L16': i2, 'S16': i3,
            'T16': head_emp if head_emp else "",
        }

        # 従たる事業所（最大10件、行21から2行ずつ merged）
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

        # 事業所数 (U12:V12 merged - R12は「事業所数：」ラベル)
        total_offices = max(1, len(offices)) if offices else 1
        values['U12'] = total_offices

        fname = f"06_様式14-1号_{company.company_name}_{group.curriculum_name}.xlsx"
        return self._patch(
            "計画申請/06_人材開発支援助成金（事業展開等リスキリング支援コース）定額制サービスによる訓練に関する事業所確認票(様式第14-1号)※定額制.xlsx",
            fname, values
        )

    # ==================================================================
    # 支給申請書類
    # ==================================================================

    def write_支給申請書(self, company: CompanyInfo, sr: SocialInsuranceLabor,
                         group: CurriculumGroup, submit_date: datetime) -> str:
        """01_支給申請書(様式第4-2号) - R80302版（セル位置は現行と同じ）"""
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
            'AG16': sp1, 'AL16': sp2,
            'AF17': sr.address if sr else "",
            'AF19': sr.office_name if sr else "",
            'AF20': sr.sr_name if sr else "",
            'AF21': spp1, 'AM21': spp2, 'AT21': spp3,
            'K26': company.main_business,
            'K27': company.employee_count,
            # R80302 (R8.3) では「設備投資加算」セクションが削除されたため
            # 雇用保険適用事業所名以降は -1 行（row 30→29、31→30、32→31、37→36）
            # 雇用保険適用事業所名（offices[0].name 優先）
            'K29': get_insurance_office_name(company),
            # 雇用保険番号 4桁-6桁-1桁
            'AN29': i1, 'AS29': i2, 'AZ29': i3,
            # 担当者氏名・所属
            'R30': company.contact_name,
            'AM30': company.contact_department,
            # 電話番号 + メール
            'R31': cp1, 'W31': cp2, 'AB31': cp3,
            'AM31': company.contact_email,
            # 男女別受講者数
            'M36': male,
            'AM36': female,
        }
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
        """02_経費助成の内訳(様式第6-2号) - R80302（R8.3）版

        新レイアウト（マージセル分析）:
          - B7:I7 = '職業訓練実施計画届の受付番号' ラベル
          - J7:W7 = 受付番号データ（外部付与のため空のまま）
          - Y7:AD7 = '雇用保険適用事業所の名称' ラベル
          - AE7:AT7 = 雇用保険適用事業所の名称データ
          ※ この様式に労働局は無い（様式5号と同じ）
        """
        values = {
            'AE7': get_insurance_office_name(company),
        }
        fname = f"02_経費助成の内訳_{group.curriculum_name}_{company.company_name}.xlsx"
        return self._patch(
            "支給申請/02_経費助成の内訳(様式第6-2号)_研修名_.xlsx",
            fname, values
        )

    def write_賃金助成内訳(self, company: CompanyInfo, group: CurriculumGroup) -> str:
        """03_賃金助成の内訳(様式第5号) - R80302版

        修正点:
          - BA7 = 雇用保険適用事業所名（offices[0].name 優先）
          - 行22以降の支給対象労働者一覧にデータを追加
          - ファイル名から (インタラクティブのみ) を削除

        ⚠️ 注意: 現在 R80302 の '03_賃金助成の内訳(様式第5号)_研修名_企業名.xlsx' の
           中身が様式6-2号（経費助成の内訳）になっている可能性があります。
           正しい様式5号テンプレートに差し替えてください。
        """
        values = {
            # 受付番号 (P7) は労働局から付与される番号なので明示的に空欄
            'P7': '',
            'BA7': get_insurance_office_name(company),
        }
        start_row = 22
        for idx, p in enumerate(group.participants):
            r = start_row + idx
            values[f'B{r}'] = p.name
            values[f'K{r}'] = getattr(p, 'furigana', '')
            values[f'T{r}'] = (p.insurance_number or "").replace('−', '-').replace('ー', '-')
        fname = f"03_賃金助成の内訳_{group.curriculum_name}_{company.company_name}.xlsx"
        return self._patch(
            "支給申請/03_賃金助成の内訳(様式第5号)_研修名_企業名.xlsx",
            fname, values
        )

    def write_経費助成内訳_定額制(self, company: CompanyInfo, group: CurriculumGroup) -> str:
        """03_様式第6-3号 ※定額制 定額制サービスによる訓練に関する経費助成の内訳（R80302）

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
        fname = f"03_様式6-3号_未生成_{group.curriculum_name}_{company.company_name}.xlsx"
        return self._patch(
            "支給申請/03_定額制サービスによる訓練に関する経費助成の内訳(様式第6-3号)※定額制.xlsx",
            fname, values,
        )

    def write_対象者一覧_支給_3_1(self, company: CompanyInfo, group: CurriculumGroup) -> str:
        """支給申請の 03_対象者一覧(様式第3-1号) - R80302版（セル位置は計画版と同じ）"""
        values = {
            'C8': company.company_name,
            'C9': group.curriculum_name,
        }
        start_row = 13
        rows_per_entry = 2
        for idx, p in enumerate(group.participants):
            r = start_row + idx * rows_per_entry
            values[f'B{r}'] = p.name
            parts = (p.insurance_number or "").replace('−', '-').replace('ー', '-').split('-')
            if len(parts) == 3:
                values[f'C{r}'] = parts[0]
                values[f'F{r}'] = parts[1]
                values[f'I{r}'] = parts[2]
        fname = f"03_対象者一覧(支給)_{group.curriculum_name}_{company.company_name}.xlsx"
        return self._patch(
            "支給申請/03_対象者一覧(様式第3-1号)_企業名.xlsx",
            fname, values
        )

    def write_事業所確認票(self, company: CompanyInfo, group: CurriculumGroup,
                          submit_date: datetime) -> str:
        """05_事業所確認票(様式第13号) - R80302版（セル位置は現行と同じ）"""
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

    def write_支給申請承諾書(self, company: CompanyInfo, group: CurriculumGroup,
                              submit_date: datetime, training_company=None) -> str:
        """04_支給申請承諾書(様式第12号) - R80302版（セル位置は現行と同じ）

        教育訓練機関欄 (training_company が指定された場合):
          F26: 所在地, F28: 名称, P30: 代表者氏名, F32: 法人番号
        """
        values = {
            # 確認日
            'R24': submit_date.year, 'V24': submit_date.month, 'Y24': "",
            # 対象訓練（1件目）
            'F40': group.curriculum_name,
            # 申請事業主
            'E51': company.address,
            'E53': company.company_name,
            'E55': company.representative_name,
        }
        if group.start_date:
            values['S40'] = group.start_date.year
            values['W40'] = group.start_date.month
            values['Y40'] = group.start_date.day
        if group.end_date:
            values['S42'] = group.end_date.year
            values['W42'] = group.end_date.month
            values['Y42'] = group.end_date.day
        # 教育訓練機関 (training_company が指定された場合のみ上書き)
        if training_company is not None:
            values['F26'] = training_company.address
            values['F28'] = training_company.name
            values['P30'] = training_company.representative_name
            values['F32'] = training_company.corporate_number

        fname = f"04_支給申請承諾書_{group.curriculum_name}_{company.company_name}.xlsx"
        return self._patch(
            "支給申請/04_人材開発支援助成金（事業展開等リスキリング支援コース）支給申請承諾書（訓練実施者）(様式第12号).xlsx",
            fname, values
        )

    # ==================================================================
    # 書類一式生成
    # ==================================================================

    def generate_plan_documents(self, company: CompanyInfo, sr: SocialInsuranceLabor,
                                 group: CurriculumGroup, submit_date: datetime,
                                 training_company=None) -> List[str]:
        """計画申請書類の一括生成（R80302）

        現行との差分:
          - Word書類（事業内職業能力開発計画）は無し
          - 定額制の場合は様式14-1号（事業所確認票）を追加
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
        """支給申請書類の一括生成（R80302）

        最新の支給申請フォルダ構成:
          - 01_支給申請書(様式4-2号)
          - 02_経費助成の内訳(様式6-2号) ← 通常時のみ
          - 03_様式6-3号 ※定額制 ← 定額制時のみ
          - 03_賃金助成の内訳(様式5号)
          - 04_支給申請承諾書(様式12号)
          - 05_事業所確認票(様式13号)
          ※ 様式3-1号(対象者一覧) は R80302 支給申請から削除済み
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
            ("賃金助成内訳", lambda: self.write_賃金助成内訳(company, group)),
            ("支給申請承諾書", lambda: self.write_支給申請承諾書(company, group, submit_date,
                                                              training_company=training_company)),
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
