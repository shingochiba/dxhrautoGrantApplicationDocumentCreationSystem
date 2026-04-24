"""書類生成メインロジックモジュール"""
import os
import zipfile
import shutil
from pathlib import Path
from datetime import datetime
from typing import List, Dict, Optional
import tempfile

from .company_form import CompanyInfo, SocialInsuranceLabor
from .upload_handler import CurriculumGroup
from .excel_writer import ExcelWriter
from .word_writer import WordWriter


class DocumentGenerator:
    """書類生成を管理するメインクラス"""

    def __init__(self, base_dir: str):
        self.base_dir = Path(base_dir)
        self.template_dir = self.base_dir / "templates"
        self.output_dir = self.base_dir / "output"
        self.output_dir.mkdir(parents=True, exist_ok=True)

    def generate_documents(self,
                          company: CompanyInfo,
                          sr: SocialInsuranceLabor,
                          groups: Dict[str, CurriculumGroup],
                          generate_plan: bool = True,
                          generate_payment: bool = True,
                          selected_curricula: Optional[List[str]] = None,
                          submit_date: Optional[datetime] = None) -> Dict[str, str]:
        """
        書類を生成してZIPファイルを作成

        Args:
            company: 会社情報
            sr: 社労士情報
            groups: カリキュラムグループ辞書
            generate_plan: 計画申請書類を生成するか
            generate_payment: 支給申請書類を生成するか
            selected_curricula: 選択されたカリキュラムキーのリスト
            submit_date: 提出日

        Returns:
            カリキュラム名とZIPファイルパスの辞書
        """
        if submit_date is None:
            submit_date = datetime.now()

        if selected_curricula is None:
            selected_curricula = list(groups.keys())

        result_files = {}

        for curriculum_key in selected_curricula:
            if curriculum_key not in groups:
                continue

            group = groups[curriculum_key]

            # カリキュラム用の一時ディレクトリを作成
            safe_name = self._sanitize_filename(group.curriculum_name)
            temp_dir = self.output_dir / f"temp_{safe_name}_{datetime.now().strftime('%Y%m%d%H%M%S')}"
            temp_dir.mkdir(parents=True, exist_ok=True)

            try:
                generated_files = []

                # 計画申請書類の生成
                if generate_plan:
                    plan_dir = temp_dir / "計画申請"
                    plan_dir.mkdir(exist_ok=True)

                    excel_writer = ExcelWriter(str(self.template_dir), str(plan_dir))
                    plan_files = excel_writer.generate_plan_documents(
                        company, sr, group, submit_date
                    )
                    generated_files.extend(plan_files)

                    # Word書類
                    word_writer = WordWriter(str(self.template_dir), str(plan_dir))
                    word_files = word_writer.generate_word_documents(
                        company, group, submit_date
                    )
                    generated_files.extend(word_files)

                # 支給申請書類の生成
                if generate_payment:
                    payment_dir = temp_dir / "支給申請"
                    payment_dir.mkdir(exist_ok=True)

                    excel_writer = ExcelWriter(str(self.template_dir), str(payment_dir))
                    payment_files = excel_writer.generate_payment_documents(
                        company, sr, group, submit_date
                    )
                    generated_files.extend(payment_files)

                # ZIPファイルの作成
                safe_company_name = self._sanitize_filename(company.company_name)
                safe_curriculum_name = self._sanitize_filename(group.curriculum_name)
                zip_filename = f"{safe_company_name}_{safe_curriculum_name}_{submit_date.strftime('%Y%m%d')}.zip"
                zip_path = self.output_dir / zip_filename

                self._create_zip(temp_dir, zip_path)
                result_files[curriculum_key] = str(zip_path)

            finally:
                # 一時ディレクトリの削除（失敗しても続行）
                if temp_dir.exists():
                    try:
                        shutil.rmtree(temp_dir, ignore_errors=True)
                    except Exception:
                        pass  # 削除失敗は無視

        return result_files

    def _sanitize_filename(self, filename: str) -> str:
        """ファイル名から不正な文字を除去"""
        import re
        # タブ、改行、制御文字を除去
        filename = re.sub(r'[\t\n\r\x00-\x1f\x7f]', '', filename)
        # Windowsで使用できない文字を置換
        invalid_chars = '<>:"/\\|?*'
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        # 前後の空白を除去
        filename = filename.strip()
        # 空文字列の場合はデフォルト名を返す
        if not filename:
            filename = "unnamed"
        return filename

    def _create_zip(self, source_dir: Path, zip_path: Path):
        """ディレクトリをZIPファイルに圧縮"""
        with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(source_dir):
                for file in files:
                    file_path = Path(root) / file
                    arcname = file_path.relative_to(source_dir)
                    zipf.write(file_path, arcname)

    def get_output_dir(self) -> Path:
        """出力ディレクトリを取得"""
        return self.output_dir

    def cleanup_old_files(self, days: int = 7):
        """古い出力ファイルを削除"""
        import time
        cutoff = time.time() - (days * 24 * 60 * 60)

        for file_path in self.output_dir.glob("*.zip"):
            if file_path.stat().st_mtime < cutoff:
                file_path.unlink()


def generate_all_documents(company: CompanyInfo,
                           sr: SocialInsuranceLabor,
                           groups: Dict[str, CurriculumGroup],
                           base_dir: str,
                           generate_plan: bool = True,
                           generate_payment: bool = True,
                           selected_curricula: Optional[List[str]] = None,
                           submit_date: Optional[datetime] = None) -> Dict[str, str]:
    """
    便利関数：全書類を生成

    Args:
        company: 会社情報
        sr: 社労士情報
        groups: カリキュラムグループ辞書
        base_dir: ベースディレクトリ
        generate_plan: 計画申請書類を生成するか
        generate_payment: 支給申請書類を生成するか
        selected_curricula: 選択されたカリキュラムキーのリスト
        submit_date: 提出日

    Returns:
        カリキュラム名とZIPファイルパスの辞書
    """
    generator = DocumentGenerator(base_dir)
    return generator.generate_documents(
        company=company,
        sr=sr,
        groups=groups,
        generate_plan=generate_plan,
        generate_payment=generate_payment,
        selected_curricula=selected_curricula,
        submit_date=submit_date
    )
