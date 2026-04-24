"""Wordテンプレートへのデータ書き込みモジュール"""
from docx import Document
from docx.shared import Pt
from pathlib import Path
from datetime import datetime
from typing import Optional
import re

from .company_form import CompanyInfo, SocialInsuranceLabor
from .upload_handler import CurriculumGroup


class WordWriter:
    """Wordテンプレートへの書き込みを行うクラス"""

    def __init__(self, template_dir: str, output_dir: str):
        self.template_dir = Path(template_dir)
        self.output_dir = Path(output_dir)
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def _load_template(self, template_path: str) -> Document:
        """テンプレートファイルを読み込む"""
        full_path = self.template_dir / template_path
        return Document(str(full_path))

    def _save_document(self, doc: Document, filename: str) -> str:
        """ドキュメントを保存する"""
        # ファイル名をサニタイズ
        filename = self._sanitize_filename(filename)
        output_path = self.output_dir / filename
        doc.save(str(output_path))
        return str(output_path)

    def _sanitize_filename(self, filename: str) -> str:
        """ファイル名から不正な文字を除去"""
        # タブ、改行、制御文字を除去
        filename = re.sub(r'[\t\n\r\x00-\x1f\x7f]', '', filename)
        # Windowsで使用できない文字を置換
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        # 前後の空白を除去
        filename = filename.strip()
        return filename if filename else "unnamed.docx"

    def _replace_text_in_paragraph(self, paragraph, old_text: str, new_text: str):
        """段落内のテキストを置換する"""
        for run in paragraph.runs:
            if old_text in run.text:
                run.text = run.text.replace(old_text, new_text)

    def _replace_text_in_document(self, doc: Document, replacements: dict):
        """ドキュメント全体でテキストを置換する"""
        for paragraph in doc.paragraphs:
            for old_text, new_text in replacements.items():
                if old_text in paragraph.text:
                    self._replace_text_in_paragraph(paragraph, old_text, new_text)

        # テーブル内も置換
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        for old_text, new_text in replacements.items():
                            if old_text in paragraph.text:
                                self._replace_text_in_paragraph(paragraph, old_text, new_text)

    def write_事業内職業能力開発計画(self, company: CompanyInfo,
                                    create_date: datetime) -> str:
        """11_事業内職業能力開発計画.docxを作成"""
        doc = self._load_template("計画申請/11_事業内職業能力開発計画.docx")

        # 置換するテキスト
        replacements = {
            "株式会社○○○○": company.company_name,
            "〇〇〇〇": company.main_business if company.main_business else "（経営方針・理念を記載）",
            "年　　月　　日作成": f"{create_date.year}年{create_date.month}月{create_date.day}日作成",
        }

        self._replace_text_in_document(doc, replacements)

        filename = f"11_事業内職業能力開発計画_{company.company_name}.docx"
        return self._save_document(doc, filename)

    def generate_word_documents(self, company: CompanyInfo,
                                 group: CurriculumGroup,
                                 create_date: datetime) -> list:
        """Word書類一式を生成"""
        generated_files = []

        try:
            file_path = self.write_事業内職業能力開発計画(company, create_date)
            generated_files.append(file_path)
        except Exception as e:
            print(f"事業内職業能力開発計画の生成に失敗: {e}")

        return generated_files
