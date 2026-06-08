# -*- coding: utf-8 -*-
import os
from pypdf import PdfReader, PdfWriter

FOLDER = '/Users/kometaninaoki/Library/CloudStorage/GoogleDrive-n.kometani@re-startlaw.com/.shortcut-targets-by-id/1pEJgbHuaj4L4ovFEISOuq-0LeHGInxYl/大中 忠生 : 詐欺未遂/資料/★大中さんと共有/証拠アップロード用フォルダ/京阪事務機器/提出正式版'
DVD = os.path.join(FOLDER, 'DVDフォルダ')

def p(name):
    return os.path.join(FOLDER, name)

# 1) 資料２の１（サンプル）を全体PDFの1枚目から抽出
src_full = p('資料２【PDF】京阪事務機器_販売管理台帳.pdf')
sample2 = p('資料２の１_京阪事務機器_販売管理台帳（サンプル）.pdf')
reader = PdfReader(src_full)
writer = PdfWriter()
writer.add_page(reader.pages[0])
with open(sample2, 'wb') as f:
    writer.write(f)
print('CREATED:', os.path.basename(sample2), os.path.getsize(sample2), 'bytes')

# 2) リネーム表（現名 -> 新名）。Trueなら最後にDVDフォルダへ移動
renames = [
    ('資料１の３【PDF】京阪事務機器_備品管理台帳_ラベル表示.pdf', '資料１の３_京阪事務機器_備品管理台帳_ラベル表示（１２２０枚）.pdf', True),
    ('資料１の４【PDF】京阪事務機器_備品管理台帳_写真表示.pdf', '資料１の４_京阪事務機器_備品管理台帳_写真表示（１２２０枚）.pdf', True),
    ('資料２【PDF】京阪事務機器_販売管理台帳.pdf', '資料２の２_京阪事務機器_販売管理台帳（１３７１枚）.pdf', True),
    ('資料３の２【PDF】京阪事務見積納品請求書.pdf', '資料３の２_京阪事務機器_見積納品請求書（２９１８枚）.pdf', True),
    ('資料４の２【PDF】京阪事務機器ラベル.pdf', '資料４の２_京阪事務機器ラベル（１４１枚）.pdf', True),
    ('資料５の２【PDF】ゼタの備品管理台帳（京阪へ譲渡したもののうちラベルがあるもの）.pdf', '資料５の２_ゼタ備品管理台帳（京阪へ譲渡したもののうちラベルがあるもの）（１３０枚）.pdf', True),
    ('資料１の１京阪事務機器_備品管理台帳_ラベル表示.pdf', '資料１の１_京阪事務機器_備品管理台帳_ラベル表示（サンプル）.pdf', False),
    ('資料１の2京阪事務機器_備品管理台帳_写真表示.pdf', '資料１の２_京阪事務機器_備品管理台帳_写真表示（サンプル）.pdf', False),
    ('資料３の１京阪事務見積納品請求書.pdf', '資料３の１_京阪事務機器_見積納品請求書（サンプル）.pdf', False),
    ('資料４の１京阪事務機器ラベル.pdf', '資料４の１_京阪事務機器ラベル（サンプル）.pdf', False),
    ('資料５の１ゼタの備品管理台帳（京阪へ譲渡したもののうちラベルがあるもの）.pdf', '資料５の１_ゼタ備品管理台帳（京阪へ譲渡したもののうちラベルがあるもの）（サンプル）.pdf', False),
    ('資料７の１.pdf', '資料７の１_業務管理システムスクリーンショット（令和４年９月２１日〜２４日）.pdf', False),
    ('資料７の２.pdf', '資料７の２_業務管理システムスクリーンショット（令和４年１０月４日）.pdf', False),
    ('資料７の３.pdf', '資料７の３_業務管理システムスクリーンショット（令和４年１２月６日）.pdf', False),
    ('資料７の４.pdf', '資料７の４_業務管理システムスクリーンショット（令和５年２月８日〜９日）.pdf', False),
]

# 検証：全ての元ファイルが存在するか
missing = [o for o, n, d in renames if not os.path.isfile(p(o))]
if missing:
    print('MISSING SOURCE FILES:', missing)
    raise SystemExit('元ファイルが見つかりません。中止します。')

# リネーム実行
for old, new, _ in renames:
    os.rename(p(old), p(new))
    print('RENAMED:', old, '->', new)

# DVDフォルダ作成
os.makedirs(DVD, exist_ok=True)
print('DVD FOLDER:', os.path.isdir(DVD))

# DVD移動（リネーム済みの新名で移動）
for old, new, dvd in renames:
    if dvd:
        os.rename(p(new), os.path.join(DVD, new))
        print('MOVED TO DVD:', new)

print('\n=== 最終状態 ===')
for root, dirs, files in os.walk(FOLDER):
    rel = os.path.relpath(root, FOLDER)
    print('[DIR]', rel)
    for f in sorted(files):
        if f == '.DS_Store':
            continue
        print('   ', f)
