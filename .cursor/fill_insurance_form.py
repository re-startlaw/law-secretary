"""Fill page 2 (資格取得届) of 01.xlsx by editing sheet1.xml directly.

Why direct XML: openpyxl drops EMF images (the barcode and form artwork) on save.
Direct edit preserves every byte of every other part.
"""
import os
import re
import shutil
import zipfile
from xml.sax.saxutils import escape

SRC = '/Users/kometaninaoki/Downloads/01.xlsx'
DST = '/Users/kometaninaoki/Downloads/01_2.xlsx'

# (cell, value, kind). kind: 'n' = number, 's' = inline string.
FILLS = [
    # 提出日（令和 _年 _月 _日提出）
    ('J230', '8', 's'), ('P230', '4', 's'), ('W230', '24', 's'),
    # 事業所整理記号 "84-101"
    ('M232', '8', 's'), ('P232', '4', 's'),
    ('AB232', '1', 's'), ('AE232', '0', 's'), ('AH232', '1', 's'),
    # 事業所番号 "11640"
    ('AT232', '1', 's'), ('AX232', '1', 's'), ('BB232', '6', 's'),
    ('BF232', '4', 's'), ('BJ232', '0', 's'),
    # 事業所所在地
    ('R241', '170', 's'), ('AA241', '6012', 's'),
    ('M244', '東京都豊島区東池袋三丁目1-1 サンシャイン60 12階', 's'),
    # 事業所名称 / 事業主氏名 / 電話番号
    ('M255', 'Re-Start法律事務所', 's'),
    ('M264', '米谷 尚起', 's'),
    ('V273', '03', 's'), ('AH273', '6820', 's'), ('AX273', '3815', 's'),
    # 被保険者1 フリガナ・氏名
    ('AB279', 'コメタニ', 's'), ('AZ279', 'ナオキ', 's'),
    ('AB282', '米谷', 's'), ('AZ282', '尚起', 's'),
    # 被保険者1 生年月日 平成06年07月20日
    ('CJ279', '0', 's'), ('CM279', '6', 's'),
    ('CP279', '0', 's'), ('CS279', '7', 's'),
    ('CV279', '2', 's'), ('CY279', '0', 's'),
    # 被保険者1 個人番号 125623113784 — numeric format "0_"
    ('AB289', 1, 'n'), ('AF289', 2, 'n'), ('AJ289', 5, 'n'), ('AN289', 6, 'n'),
    ('AR289', 2, 'n'), ('AV289', 3, 'n'), ('AZ289', 1, 'n'), ('BD289', 1, 'n'),
    ('BH289', 3, 'n'), ('BL289', 7, 'n'), ('BP289', 8, 'n'), ('BT289', 4, 'n'),
    # 被保険者1 取得年月日 令和08年04月23日
    ('CJ289', '0', 's'), ('CM289', '8', 's'),
    ('CP289', '0', 's'), ('CS289', '4', 's'),
    ('CV289', '2', 's'), ('CY289', '3', 's'),
    # 被保険者1 報酬月額 ㋐通貨 (#,##0_ format → numeric)
    ('T295', 400000, 'n'),
    # 被保険者1 報酬月額 ㋒合計 9桁右詰め（400000）
    ('BE299', '4', 's'), ('BH299', '0', 's'), ('BK299', '0', 's'),
    ('BN299', '0', 's'), ('BQ299', '0', 's'), ('BT299', '0', 's'),
    # 被保険者1 住所
    ('O305', '171', 's'), ('U305', '0022', 's'),
    ('AF307', '東京都豊島区南池袋二丁目18-9 マ・シャンブル南池袋902', 's'),
]


def patch_cell(xml_text: str, coord: str, value, kind: str) -> str:
    """Replace an empty <c r="..." s="..."/> with a populated version."""
    pattern = re.compile(rf'<c\s+r="{re.escape(coord)}"([^/>]*)/>')
    m = pattern.search(xml_text)
    if not m:
        raise RuntimeError(f'cell {coord} not found as empty self-closing <c/>')
    attrs = m.group(1)  # e.g. ' s="432"'
    if kind == 'n':
        replacement = f'<c r="{coord}"{attrs}><v>{value}</v></c>'
    elif kind == 's':
        replacement = f'<c r="{coord}"{attrs} t="inlineStr"><is><t xml:space="preserve">{escape(str(value))}</t></is></c>'
    else:
        raise ValueError(kind)
    return xml_text[:m.start()] + replacement + xml_text[m.end():]


def main():
    with zipfile.ZipFile(SRC, 'r') as zin:
        sheet_xml = zin.read('xl/worksheets/sheet1.xml').decode('utf-8')

    for coord, value, kind in FILLS:
        sheet_xml = patch_cell(sheet_xml, coord, value, kind)

    if os.path.exists(DST):
        os.remove(DST)
    # Copy every part from SRC to DST, swapping in the patched sheet1.xml.
    with zipfile.ZipFile(SRC, 'r') as zin, zipfile.ZipFile(DST, 'w', zipfile.ZIP_DEFLATED) as zout:
        for info in zin.infolist():
            if info.filename == 'xl/worksheets/sheet1.xml':
                zout.writestr(info, sheet_xml.encode('utf-8'))
            else:
                zout.writestr(info, zin.read(info.filename))
    print(f'Saved: {DST}')


if __name__ == '__main__':
    main()
