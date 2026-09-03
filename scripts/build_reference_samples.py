# -*- coding: utf-8 -*-
"""
参考データ（照合パターン）ファイル生成スクリプト
参考数据（匹配模式）文件生成脚本

実データ画面（サクラエディタ / O:\\work\\NDLC\\...\\*002.txt）で確認された
照合パターンのテストデータを再現する。
复现在实际数据画面中确认到的匹配模式测试数据。

レコードIDの命名規則 / 记录ID命名规则:
    1<f1>2<f2>3<f3>[h]
      f1 : 名称（フリガナ） A=一致 / B=不一致          名称（片假名） A=匹配 / B=不匹配
      f2 : 名称（漢字）     A=一致 / B=不一致 / C=未設定  名称（汉字） A=匹配 / B=不匹配 / C=未设置
      f3 : 設立/生年月日    A=一致 / B=不一致 / C=未設定  出生年月日   A=匹配 / B=不匹配 / C=未设置
      h  : ヒット想定（フリガナ一致 かつ 漢字が不一致でない かつ 生年月日が不一致でない）
           预期命中（片假名匹配 且 汉字非不匹配 且 出生日期非不匹配）

実行 / 执行:  python scripts/build_reference_samples.py
"""
import os
import sys

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

# ----------------------------------------------------------------------
# 照合パターンの値 / 匹配模式取值
# ----------------------------------------------------------------------
FURIGANA = {'A': 'アイダマナ', 'B': 'アイタナナ'}          # No.3  名称（フリガナ）
KANJI    = {'A': '相田　マナ', 'B': '相他ナナ', 'C': ''}    # No.4  名称（漢字）
BIRTHDAY = {'A': '20000101', 'B': '20010101', 'C': ''}      # No.12 設立/生年月日
BUTEN    = '0011'                                           # No.13 部店コード（実データ準拠）

# 設計書 No.14 以降は原本未入手。実データ画面から推定した末尾列。
# 规格书 No.14 之后未获取，以下为根据实际数据画面推定的尾部列。
EXTRA_COLUMNS = [''] * 8 + ['A100000001-01', '0']


def build_records():
    """18レコード(2x3x3)を実データと同じ並び順で生成する。"""
    records = []
    for f1 in ('A', 'B'):
        for f3 in ('A', 'B', 'C'):
            for f2 in ('A', 'B', 'C'):
                hit = (f1 == 'A') and (f2 != 'B') and (f3 != 'B')
                rec_id = f'1{f1}2{f2}3{f3}' + ('h' if hit else '')

                row = [''] * 13
                row[0] = rec_id            # No.1  レコードID
                row[2] = FURIGANA[f1]      # No.3  名称（フリガナ）
                row[3] = KANJI[f2]         # No.4  名称（漢字）
                row[11] = BIRTHDAY[f3]     # No.12 設立/生年月日
                row[12] = BUTEN            # No.13 部店コード
                records.append((f1, f2, f3, hit, row))
    return records


def write_tsv(path, rows, encoding):
    text = ''.join('\t'.join(r) + '\r\n' for r in rows)
    with open(path, 'wb') as fp:
        fp.write(text.encode(encoding))
    return len(text.encode(encoding))


def build_mapping_tsv(path, records):
    """レコードIDと期待結果の対応表（Excel / エディタで確認用）。"""
    header = ['No.', 'レコードID / 记录ID', 'f1 フリガナ / 片假名', 'f2 漢字 / 汉字',
              'f3 生年月日 / 出生日期', 'ヒット想定 / 预期命中',
              '名称（フリガナ）', '名称（漢字）', '設立/生年月日', '部店コード']
    label1 = {'A': '一致 / 匹配', 'B': '不一致 / 不匹配'}
    label = {'A': '一致 / 匹配', 'B': '不一致 / 不匹配', 'C': '未設定 / 未设置'}
    lines = ['\t'.join(header)]
    for i, (f1, f2, f3, hit, row) in enumerate(records, start=1):
        lines.append('\t'.join([
            str(i), row[0], f'{f1}：{label1[f1]}', f'{f2}：{label[f2]}', f'{f3}：{label[f3]}',
            'ヒット / 命中' if hit else '−',
            row[2], row[3] or '(未設定)', row[11] or '(未設定)', row[12],
        ]))
    # 日中併記のため UTF-8 BOM 付きで出力（Excel でそのまま開ける）
    # 因中日双语并记，输出为带 BOM 的 UTF-8（Excel 可直接打开）
    with open(path, 'wb') as fp:
        fp.write(('\r\n'.join(lines) + '\r\n').encode('utf-8-sig'))
    return len(lines) - 1


def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out = os.path.join(root, 'samples')
    os.makedirs(out, exist_ok=True)

    records = build_records()
    rows13 = [r[4] for r in records]
    rows23 = [r[4] + EXTRA_COLUMNS for r in records]

    files = [
        ('参考データ_照合パターン_13項目_SJIS.txt', rows13, 'cp932',
         '設計書 No.1〜13 のみ / 仅规格书 No.1〜13'),
        ('参考データ_照合パターン_13項目_UTF8.txt', rows13, 'utf-8',
         '同上（UTF-8 版・確認用）/ 同上（UTF-8 版・供查看）'),
        ('参考データ_照合パターン_実データ再現_23項目_SJIS.txt', rows23, 'cp932',
         '実データ画面の再現（No.14以降は推定）/ 复现实际数据画面（No.14 之后为推定）'),
    ]
    for name, rows, enc, note in files:
        size = write_tsv(os.path.join(out, name), rows, enc)
        print(f'[OK] {name}  ({len(rows)}件 / {size} bytes / {enc})  {note}')

    n = build_mapping_tsv(os.path.join(out, '参考データ_パターン対応表.tsv'), records)
    print(f'[OK] 参考データ_パターン対応表.tsv  ({n}件 / UTF-8 BOM)  Excelで開ける対応表 / 可用Excel打开的对照表')

    hits = sum(1 for r in records if r[3])
    print(f'\n     合計 {len(records)} パターン（ヒット想定 {hits} / 非ヒット {len(records) - hits}）')


if __name__ == '__main__':
    main()
