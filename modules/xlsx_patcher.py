"""
xlsxテンプレートのセル値のみをZIPレベルで書き換えるモジュール。

openpyxlで保存するとチェックボックス（フォームコントロール）や一部の
描画要素が失われるため、テンプレートのZIPアーカイブを直接操作して
xl/worksheets/sheetN.xml のセル値だけを差し替える。
他のパーツ（ctrlProps、drawings、vmlDrawing、media、スタイル、マージセル等）
は一切変更しないため、チェックボックスも完全に保持される。
"""
from __future__ import annotations

import re
import shutil
import zipfile
from pathlib import Path
from typing import Dict, Any, Tuple


_CELL_REF_RE = re.compile(r'^([A-Z]+)(\d+)$')


def col_letter_to_num(letter: str) -> int:
    """'A'->1, 'B'->2, 'AA'->27"""
    num = 0
    for c in letter.upper():
        num = num * 26 + (ord(c) - ord('A') + 1)
    return num


def parse_cell_ref(ref: str) -> Tuple[str, int]:
    m = _CELL_REF_RE.match(ref.strip())
    if not m:
        raise ValueError(f"Invalid cell reference: {ref}")
    return m.group(1), int(m.group(2))


def _escape_xml(s: str) -> str:
    return (str(s)
            .replace('&', '&amp;')
            .replace('<', '&lt;')
            .replace('>', '&gt;')
            .replace('"', '&quot;')
            .replace("'", '&apos;'))


def _build_cell_xml(ref: str, value: Any, existing_attrs: str = "") -> str:
    """
    新しい <c> 要素のXMLを構築。
    existing_attrs: 既存セルの属性文字列（例: ' s="5"'）。スタイルを保持するため。
    """
    # 既存属性から t="..." を除去（新しい型を設定するため）
    attrs = re.sub(r'\s*t="[^"]*"', '', existing_attrs)
    attrs = attrs.strip()
    attr_str = f' {attrs}' if attrs else ''

    if isinstance(value, bool):
        v = '1' if value else '0'
        return f'<c r="{ref}"{attr_str} t="b"><v>{v}</v></c>'
    if isinstance(value, (int, float)):
        return f'<c r="{ref}"{attr_str}><v>{value}</v></c>'
    text = _escape_xml(str(value))
    return (f'<c r="{ref}"{attr_str} t="inlineStr">'
            f'<is><t xml:space="preserve">{text}</t></is></c>')


def _replace_or_insert_cell(sheet_xml: str, ref: str, value: Any) -> str:
    """
    sheet.xml 内で ref のセルを書き換え、または該当行に挿入する。
    """
    col_letter, row_num = parse_cell_ref(ref)
    col_num = col_letter_to_num(col_letter)

    # 既存セルを正規表現で検出: <c r="REF" ...>...</c> または <c r="REF" .../>
    cell_re = re.compile(
        rf'<c\s+r="{re.escape(ref)}"(?P<attrs>[^>]*?)(?P<close>/>|>(?P<body>.*?)</c>)',
        re.DOTALL
    )
    m = cell_re.search(sheet_xml)
    if m:
        new_cell = _build_cell_xml(ref, value, m.group('attrs'))
        return sheet_xml[:m.start()] + new_cell + sheet_xml[m.end():]

    # セルがなければ行を探す
    row_re = re.compile(
        rf'(?P<rowstart><row[^>]*\br="{row_num}"[^>]*>)(?P<rowbody>.*?)</row>',
        re.DOTALL
    )
    rm = row_re.search(sheet_xml)
    new_cell = _build_cell_xml(ref, value, "")
    if rm:
        # 行内のセルを列順で並べる必要がある
        body = rm.group('rowbody')
        # body内のセルを一覧化
        sub_cells = list(re.finditer(
            r'<c\s+r="([A-Z]+\d+)"[^>]*(?:/>|>.*?</c>)', body, re.DOTALL))
        # col_num より大きな最初のセルの前に挿入
        insert_pos = len(body)  # デフォルトは末尾
        for sc in sub_cells:
            sc_ref = sc.group(1)
            sc_col, _ = parse_cell_ref(sc_ref)
            if col_letter_to_num(sc_col) > col_num:
                insert_pos = sc.start()
                break
        new_body = body[:insert_pos] + new_cell + body[insert_pos:]
        return (sheet_xml[:rm.start('rowbody')]
                + new_body
                + sheet_xml[rm.end('rowbody'):])

    # 行もない場合: </sheetData> の直前に新規行を追加
    new_row = f'<row r="{row_num}">{new_cell}</row>'
    sd_close = sheet_xml.find('</sheetData>')
    if sd_close == -1:
        return sheet_xml  # sheetDataが見当たらない場合は諦める
    return sheet_xml[:sd_close] + new_row + sheet_xml[sd_close:]


def _resolve_merged_cell_target(sheet_xml: str, ref: str) -> str:
    """
    ref がマージセルの一部（左上以外）の場合、マージ範囲の左上セルに解決する。
    """
    # <mergeCells>...<mergeCell ref="A1:B2"/>...</mergeCells>
    merge_re = re.compile(r'<mergeCell\s+ref="([A-Z]+\d+):([A-Z]+\d+)"\s*/>')
    col_letter, row_num = parse_cell_ref(ref)
    col_num = col_letter_to_num(col_letter)
    for mm in merge_re.finditer(sheet_xml):
        top_left = mm.group(1)
        bottom_right = mm.group(2)
        tl_col_letter, tl_row = parse_cell_ref(top_left)
        br_col_letter, br_row = parse_cell_ref(bottom_right)
        tl_col = col_letter_to_num(tl_col_letter)
        br_col = col_letter_to_num(br_col_letter)
        if tl_col <= col_num <= br_col and tl_row <= row_num <= br_row:
            return top_left
    return ref


def patch_xlsx(template_path: str | Path,
               output_path: str | Path,
               cell_values: Dict[str, Any],
               sheet_file: str = 'xl/worksheets/sheet1.xml') -> str:
    """
    テンプレートxlsxをコピーし、指定セル値のみをXMLレベルで書き換える。

    Args:
        template_path: 元テンプレートへのパス
        output_path: 出力先パス
        cell_values: {'A1': 'text', 'B2': 123, ...} 形式の辞書
        sheet_file: 編集対象シートのXMLパス

    Returns:
        出力ファイルパス
    """
    template_path = Path(template_path)
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    # テンプレートをコピー
    shutil.copy(str(template_path), str(output_path))

    # ZIPから該当sheet.xmlを読み込み
    with zipfile.ZipFile(str(output_path), 'r') as zin:
        if sheet_file not in zin.namelist():
            # 大文字小文字違い等を救済
            candidates = [n for n in zin.namelist()
                          if n.lower() == sheet_file.lower()]
            if candidates:
                sheet_file = candidates[0]
            else:
                raise FileNotFoundError(
                    f"{sheet_file} not found in {template_path}")
        sheet_xml_bytes = zin.read(sheet_file)
        all_names = list(zin.infolist())
        other_data = {n.filename: zin.read(n.filename)
                      for n in all_names if n.filename != sheet_file}

    # デコード（UTF-8を前提。BOM保持）
    has_bom = sheet_xml_bytes[:3] == b'\xef\xbb\xbf'
    sheet_xml = sheet_xml_bytes.decode('utf-8-sig' if has_bom else 'utf-8')

    # 各セル値を書き換え
    # 注: 空文字列 "" はセルクリア用途で許容（テンプレートのプレースホルダ '●' 等を消す目的）
    # None のみ「書き込まない」スキップ対象とする
    for ref, value in cell_values.items():
        if value is None:
            continue
        target_ref = _resolve_merged_cell_target(sheet_xml, ref)
        sheet_xml = _replace_or_insert_cell(sheet_xml, target_ref, value)

    # 再エンコード
    new_bytes = sheet_xml.encode('utf-8')
    if has_bom:
        new_bytes = b'\xef\xbb\xbf' + new_bytes

    # ZIPを再構築（全ファイルを順番に再書き込み）
    temp_path = str(output_path) + '.tmp'
    with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info in all_names:
            if info.filename == sheet_file:
                zout.writestr(info.filename, new_bytes)
            else:
                zout.writestr(info.filename, other_data[info.filename])

    # 置き換え
    Path(temp_path).replace(output_path)
    return str(output_path)
