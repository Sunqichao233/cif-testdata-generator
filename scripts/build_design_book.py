# -*- coding: utf-8 -*-
"""
CIF情報ファイル テストデータ設計書 (Excel) 生成スクリプト  ※日中対照版
CIF信息文件 测试数据设计书 (Excel) 生成脚本  ※中日对照版

設計書「CIF情報ファイルの項目」および実データ画面を元に、6シート構成のブックを出力する。
根据规格书《CIF信息文件的项目》与实际数据画面，输出6个 sheet 的工作簿。

    0. 表紙              / 封面
    1. 項目定義          / 项目定义      : No.1〜13 の仕様 + 検索対象者用 / CIF 列
    2. 生成条件          / 生成条件      : 対象ファイル・件数・部店コード・出力仕様
    3. テストケース一覧  / 测试用例一览  : 正常/境界/異常/照合系ケース
    4. 実行手順          / 执行步骤      : Amazon WorkSpaces 上の作業チェックリスト
    5. 参考データ        / 参考数据      : 照合パターン(2x3x3=18)の実データ再現

実行 / 执行:  python scripts/build_design_book.py
必要 / 依赖:  openpyxl (pip install openpyxl)
"""
import os
import sys
import datetime

try:
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
except Exception:
    pass

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ----------------------------------------------------------------------
# 共通スタイル / 通用样式
# ----------------------------------------------------------------------
FONT_BASE = Font(name='ＭＳ Ｐゴシック', size=9)
FONT_HEAD = Font(name='ＭＳ Ｐゴシック', size=9, bold=True)
FONT_TITLE = Font(name='ＭＳ Ｐゴシック', size=14, bold=True)
FONT_NOTE = Font(name='ＭＳ Ｐゴシック', size=9, color='C00000')
FONT_CN = Font(name='ＭＳ Ｐゴシック', size=9, color='1F4E79')

FILL_HEAD = PatternFill('solid', fgColor='DCE6F1')
FILL_HEAD_CN = PatternFill('solid', fgColor='EAF1F8')
FILL_SUB = PatternFill('solid', fgColor='F2F2F2')
FILL_REQ = PatternFill('solid', fgColor='FFF2CC')
FILL_HIT = PatternFill('solid', fgColor='E2EFDA')
FILL_EST = PatternFill('solid', fgColor='FCE4EC')

THIN = Side(style='thin', color='808080')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

AL_TL = Alignment(horizontal='left', vertical='top', wrap_text=True)
AL_CT = Alignment(horizontal='center', vertical='center', wrap_text=True)

DOC_VERSION = '2.0'
DOC_DATE = datetime.date.today().strftime('%Y/%m/%d')


def write_row(ws, row, values, font=FONT_BASE, fill=None, align=AL_TL, height=None, cn_cols=()):
    for i, v in enumerate(values, start=1):
        c = ws.cell(row=row, column=i, value=v)
        c.font = FONT_CN if (i in cn_cols and font is FONT_BASE) else font
        c.alignment = align
        c.border = BORDER
        if fill:
            c.fill = FILL_HEAD_CN if (i in cn_cols and fill is FILL_HEAD) else fill
    if height:
        ws.row_dimensions[row].height = height
    return row + 1


def set_widths(ws, widths):
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def sheet_title(ws, row, ja, cn, span):
    c = ws.cell(row=row, column=1, value=ja)
    c.font = FONT_TITLE
    ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=span)
    ws.row_dimensions[row].height = 22
    c2 = ws.cell(row=row + 1, column=1, value=cn)
    c2.font = FONT_CN
    ws.merge_cells(start_row=row + 1, start_column=1, end_row=row + 1, end_column=span)
    return row + 3


def note_lines(ws, row, lines):
    for text, font in lines:
        c = ws.cell(row=row, column=1, value=text)
        c.font = font
        c.alignment = AL_TL
        row += 1
    return row


# ======================================================================
# シート1 : 項目定義 / 项目定义
# ======================================================================
# (No, 項目, 项目, 最大桁数, 可変/固定, 文字種, 必須, 備考, 备注, 検索対象者用, CIF)
FIELDS = [
    (1, 'レコードID', '记录ID', 30, '可変', 'c', '○',
     'ファイル作成システム側のユニークなID（CIF番号など）、後続スペースはカットして有効桁のみセットする\n'
     '※同一CIF情報ファイル内でレコードIDが重複している場合、エラーとなり取込不可\n'
     '※検索対象者の場合はレコードIDはシステムで自動採番するため設定不要',
     '文件创建系统侧的唯一ID（CIF号等）。尾部空格需去除，仅设置有效位。\n'
     '※同一CIF信息文件内记录ID重复时报错，无法导入。\n'
     '※检索对象者的记录ID由系统自动采番，无需设置。',
     '連番 / 连号', '連番 / 连号'),
    (2, 'タイプ', '类型', 1, '可変', 'c', '',
     '個人：「I」（Individualの頭文字）、 それ以外：「」',
     '个人："I"（Individual 首字母）；其他：""', '-', '-'),
    (3, '名称（フリガナ）', '名称（片假名）', 200, '可変', 'k', '○',
     '名称（フリガナ）、名称（漢字）、名称（英字）のうち、少なくとも一項目は必須',
     '片假名、汉字、英文名称中至少必须设置一项',
     '重複しない / 不重复', '重複しない / 不重复'),
    (4, '名称（漢字）', '名称（汉字）', 200, '可変', 'j', '',
     'フリガナ、英字については、半角スペースの入力可',
     '片假名与英文可输入半角空格（汉字请用全角空格）', '-', '-'),
    (5, '名称（英字）', '名称（英文）', 200, '可変', 'c', '', '', '', '-', '-'),
    (6, '電話番号', '电话号码', 30, '可変', 'c', '', '', '', '-', '-'),
    (7, 'ID情報', 'ID信息', 30, '可変', 'c', '', '', '', '-', '-'),
    (8, 'ID情報種類', 'ID信息种类', 2, '固定', 'n', '',
     '運転免許証、旅券などに対応するコード値（導入時にヒアリングシートで確認いたします）',
     '对应驾照、护照等的代码值（导入时通过访谈表确认）', '-', '-'),
    (9, '国内郵便番号', '国内邮编', 7, '固定', 'n', '',
     '数値7桁（ハイフンなし）', '数值7位（无连字符）', '-', '-'),
    (10, '国内住所（市区町村以降）', '国内住址（市区町村以后）', 200, '可変', 'j', '', '', '', '-', '-'),
    (11, '英文住所', '英文地址', 256, '可変', 'c', '',
     '半角スペースの入力可\nリストの都市・住所と照合するため国名は記入しない',
     '可输入半角空格。\n为与列表的城市・地址进行匹配，不填写国名。', '-', '-'),
    (12, '設立/生年月日', '成立日 / 出生年月日', 8, '固定', 'n', '',
     '数値8桁（YYYYMMDD）\n生年月日チェックを行う場合は、リストの生年月日と照合対象',
     '数值8位（YYYYMMDD）。\n进行出生日期检查时，与列表的出生日期进行匹配。', '-', '-'),
    (13, '部店コード', '部店代码', 32, '可変', 'c', '○',
     '部店コードは、導入時にヒアリングシートで確認いたします',
     '部店代码在导入时通过访谈表确认',
     '横浜銀行用：0003、L&F用：LF002\n横滨银行用：0003、L&F用：LF002',
     '*001.txt, *002.txt, *003.txt\n001：横浜銀行用 → 0005\n002：L&F①用 → LF003\n003：L&F②用 → LF004'),
]

CHAR_TYPES = [
    ('c', '半角英数記号', '半角英数符号', '半角スペースの入力可。ASCII 0x20〜0x7E。/ 可输入半角空格。'),
    ('k', '全角カタカナ', '全角片假名', '半角スペースの入力可（フリガナ）。/ 可输入半角空格。'),
    ('j', '全角文字（漢字・かな）', '全角字符（汉字・假名）', 'タブ・改行は区切り文字のため使用不可。/ 制表符与换行不可使用。'),
    ('n', '半角数値', '半角数值', '0〜9のみ。ハイフン・記号は不可。/ 仅 0〜9，不可用连字符与符号。'),
]

FILE_FORMAT = [
    ('区切り文字 / 分隔符', 'タブ（TAB / 0x09）'),
    ('1レコードの項目数 / 每条记录的列数',
     '13項目（設計書 No.1〜13）※実データでは No.14 以降が存在（シート5参照）/ ※实际数据存在 No.14 之后的列（见 sheet5）'),
    ('レコード区切り / 记录分隔', 'CRLF'),
    ('文字コード / 字符编码', 'Shift_JIS（CP932）※ config.json の output.encoding で変更可'),
    ('拡張子・ファイル名 / 扩展名・文件名', '.txt　*001.txt, *002.txt, *003.txt'),
]


def build_sheet_fields(wb):
    ws = wb.create_sheet('1.項目定義')
    set_widths(ws, [5, 24, 24, 9, 9, 7, 6, 52, 52, 24, 28])
    r = sheet_title(ws, 1, 'CIF情報ファイルの項目', 'CIF信息文件的项目', 11)

    r = note_lines(ws, r, [
        ('次の項目セットをタブ区切りで1レコードに作成してください。 / 请将下列项目集以制表符分隔生成为1条记录。', FONT_BASE),
        ('※必須欄の○は設定必須の項目です。 / ※必须栏的○表示必须设置的项目。', FONT_BASE),
        ('※本シートは設計書原本の転記です。生成ツール（scripts\\Generate-CifTestData.ps1）の '
         '$Script:FieldDefs と1対1で対応します。 / 本 sheet 为规格书原件的转录，与生成工具的校验定义一一对应。', FONT_NOTE),
    ]) + 1

    head_row = r
    r = write_row(ws, r, ['No.', '項目', '项目', '最大桁数\n最大位数', '可変/固定\n可变/固定',
                          '文字種\n字符种', '必須\n必填', '備考', '备注',
                          '検索対象者用 / 检索对象者用', 'CIF'],
                  font=FONT_HEAD, fill=FILL_HEAD, align=AL_CT, height=30, cn_cols=(3, 9))

    for f in FIELDS:
        lines = max(len(str(f[7]).split('\n')), len(str(f[8]).split('\n')), len(str(f[10]).split('\n')), 1)
        r = write_row(ws, r, list(f), height=max(16, 12.5 * lines), cn_cols=(3, 9))
        for col in (1, 4, 5, 6, 7):
            ws.cell(row=r - 1, column=col).alignment = AL_CT
        if f[6] == '○':
            ws.cell(row=r - 1, column=7).fill = FILL_REQ
            ws.cell(row=r - 1, column=2).font = FONT_HEAD

    ws.freeze_panes = ws.cell(row=head_row + 1, column=1)
    ws.auto_filter.ref = f'A{head_row}:K{r - 1}'

    r += 2
    ws.cell(row=r, column=1, value='【文字種の定義 / 字符种定义】').font = FONT_HEAD
    r += 1
    r = write_row(ws, r, ['記号', '文字種', '字符种', '補足 / 补充'] + [''] * 7,
                  font=FONT_HEAD, fill=FILL_HEAD, align=AL_CT, cn_cols=(3,))
    for t in CHAR_TYPES:
        r = write_row(ws, r, list(t) + [''] * 7, cn_cols=(3,))
        ws.cell(row=r - 1, column=1).alignment = AL_CT

    r += 2
    ws.cell(row=r, column=1, value='【ファイル形式 / 文件格式】').font = FONT_HEAD
    r += 1
    for k, v in FILE_FORMAT:
        r = write_row(ws, r, [k, '', v] + [''] * 8)
        ws.merge_cells(start_row=r - 1, start_column=1, end_row=r - 1, end_column=2)
        ws.merge_cells(start_row=r - 1, start_column=3, end_row=r - 1, end_column=11)
    return ws


# ======================================================================
# シート2 : 生成条件 / 生成条件
# ======================================================================
TARGETS = [
    ('001', 'CIF', 'CIF 横浜銀行用', 'CIF 横滨银行用', '0005', 'CIF_yyyymmdd_001.txt', 100,
     '設計書 CIF列 001：横浜銀行用→0005'),
    ('002', 'CIF', 'CIF L&F①用', 'CIF L&F①用', 'LF003', 'CIF_yyyymmdd_002.txt', 100,
     '設計書 CIF列 002：L&F①用→LF003'),
    ('003', 'CIF', 'CIF L&F②用', 'CIF L&F②用', 'LF004', 'CIF_yyyymmdd_003.txt', 100,
     '設計書 CIF列 003：L&F②用→LF004'),
    ('S01', '検索対象者', '検索対象者 横浜銀行用', '检索对象者 横滨银行用', '0003', 'SEARCH_yyyymmdd_S01.txt', 50,
     '設計書 検索対象者用列：横浜銀行用→0003'),
    ('S02', '検索対象者', '検索対象者 L&F用', '检索对象者 L&F用', 'LF002', 'SEARCH_yyyymmdd_S02.txt', 50,
     '設計書 検索対象者用列：L&F用→LF002'),
]

PROFILES = [
    ('minimal', '必須項目のみ（No.1 / No.3 / No.13）',
     '仅必填项目（No.1 记录ID / No.3 片假名名称 / No.13 部店代码）。用于导入的基本连通确认。'),
    ('full', '全13項目を設定',
     '设置全部13个项目：类型、汉字・英文姓名、电话、ID信息、ID种类、邮编、国内住址、英文地址、出生日期。'),
    ('matching', '照合パターン 2×3×3＝18レコード',
     '匹配（名寄せ）测试用。记录ID＝1<f1>2<f2>3<f3>[h]，详见 sheet5「参考数据」。'),
    ('-IncludeBoundary（追加スイッチ）', '各ファイル末尾に境界値レコード2件を追加',
     '在各文件末尾追加"恰好等于最大位数"的边界值记录2条（记录ID 30位・片假名 200字・英文地址 256字・部店代码 32位）。'),
]

RULES = [
    ('No.1 レコードID', '連番（既定10桁ゼロ埋め）。ファイル間で通し番号を継続し、同一ファイル内で重複させない。',
     '连号（默认10位补零）。跨文件连续编号，同一文件内不重复。'),
    ('No.1（検索対象者）', 'システム自動採番のため設定不要。setRecordIdForSearch = false で未設定にできる。',
     '系统自动采番，无需设置。设为 false 即可留空。'),
    ('No.3 名称（フリガナ）', '姓20種 × 名20種＝400通り。400件超はカタカナ連番を付加し重複を回避。',
     '姓20种 × 名20种＝400组合。超过400条时附加片假名连号以避免重复。'),
    ('No.4 名称（漢字）', 'フリガナと同一人物の漢字表記。全角スペース区切り。',
     '与片假名同一人物的汉字写法，用全角空格分隔。'),
    ('No.5 名称（英字）', 'フリガナと同一人物のローマ字表記。半角スペース区切り。',
     '与片假名同一人物的罗马字写法，用半角空格分隔。'),
    ('No.8 ID情報種類', '01〜05を巡回設定（暫定値）。正式値は導入時のヒアリングシートで確定。',
     '循环设置 01〜05（暂定值）。正式取值在导入时通过访谈表确定。'),
    ('No.9 / 10 / 11', '実在の8住所パターンを巡回。英文住所には国名を含めない。',
     '循环使用8个真实住址组合。英文地址不含国名。'),
    ('No.12 設立/生年月日', '1950〜2005年の実在する日付をYYYYMMDD形式で生成。',
     '生成 1950〜2005 年之间真实存在的日期（YYYYMMDD）。'),
    ('No.13 部店コード', '対象ファイルごとに固定値（上表のとおり）。',
     '按目标文件使用固定值（见上表）。'),
    ('No.14 以降 / No.14 之后', '設計書原本が未入手のため既定では出力しない。config.json の extraColumns で有効化できる。',
     '因未获取规格书原件，默认不输出。可通过 config.json 的 extraColumns 启用。'),
]

OUTPUT_SPEC = [
    ('文字コード / 字符编码', 'Shift_JIS（CP932）／ utf-8 ／ utf-8-bom を config.json で選択'),
    ('改行コード / 换行符', 'CRLF（既定）／ LF'),
    ('区切り文字 / 分隔符', 'タブ（TAB）、1レコード13項目（extraColumns 有効時は 13＋n 項目）'),
    ('出力先 / 输出目录', 'output\\ 配下'),
    ('参考データ / 参考数据', 'samples\\ 配下（シート5参照 / 见 sheet5）'),
    ('ログ / 日志', 'logs\\generate_yyyyMMdd_HHmmss.log（UTF-8 BOM付）'),
    ('終了コード / 退出码', '0＝検証OK ／ 1＝検証NG（1件以上のエラー検出）'),
]


def build_sheet_conditions(wb):
    ws = wb.create_sheet('2.生成条件')
    set_widths(ws, [8, 13, 22, 22, 11, 26, 9, 46])
    r = sheet_title(ws, 1, 'テストデータ生成条件', '测试数据生成条件', 8)

    ws.cell(row=r, column=1, value='【生成対象ファイル / 生成目标文件】').font = FONT_HEAD
    r += 1
    r = write_row(ws, r, ['ID', '区分', '名称', '名称(中文)', '部店コード\n部店代码',
                          'ファイル名 / 文件名', '既定件数\n默认条数', '根拠（設計書）/ 依据（规格书）'],
                  font=FONT_HEAD, fill=FILL_HEAD, align=AL_CT, height=26, cn_cols=(4,))
    for t in TARGETS:
        r = write_row(ws, r, list(t), cn_cols=(4,))
        for col in (1, 2, 5, 7):
            ws.cell(row=r - 1, column=col).alignment = AL_CT

    r += 2
    ws.cell(row=r, column=1, value='【データプロファイル / 数据配置】').font = FONT_HEAD
    r += 1
    r = write_row(ws, r, ['プロファイル\n配置名', '設定内容', '', '说明(中文)', '', '', '', ''],
                  font=FONT_HEAD, fill=FILL_HEAD, align=AL_CT, height=26, cn_cols=(4,))
    ws.merge_cells(start_row=r - 1, start_column=2, end_row=r - 1, end_column=3)
    ws.merge_cells(start_row=r - 1, start_column=4, end_row=r - 1, end_column=8)
    for p in PROFILES:
        r = write_row(ws, r, [p[0], p[1], '', p[2], '', '', '', ''], height=26, cn_cols=(4,))
        ws.merge_cells(start_row=r - 1, start_column=2, end_row=r - 1, end_column=3)
        ws.merge_cells(start_row=r - 1, start_column=4, end_row=r - 1, end_column=8)

    r += 2
    ws.cell(row=r, column=1, value='【項目別 生成ルール / 各项目生成规则】').font = FONT_HEAD
    r += 1
    r = write_row(ws, r, ['項目 / 项目', '', '生成ルール', '', '生成规则(中文)', '', '', ''],
                  font=FONT_HEAD, fill=FILL_HEAD, align=AL_CT, cn_cols=(5,))
    ws.merge_cells(start_row=r - 1, start_column=1, end_row=r - 1, end_column=2)
    ws.merge_cells(start_row=r - 1, start_column=3, end_row=r - 1, end_column=4)
    ws.merge_cells(start_row=r - 1, start_column=5, end_row=r - 1, end_column=8)
    for rule in RULES:
        r = write_row(ws, r, [rule[0], '', rule[1], '', rule[2], '', '', ''], height=26, cn_cols=(5,))
        ws.merge_cells(start_row=r - 1, start_column=1, end_row=r - 1, end_column=2)
        ws.merge_cells(start_row=r - 1, start_column=3, end_row=r - 1, end_column=4)
        ws.merge_cells(start_row=r - 1, start_column=5, end_row=r - 1, end_column=8)

    r += 2
    ws.cell(row=r, column=1, value='【出力仕様 / 输出规格】').font = FONT_HEAD
    r += 1
    for k, v in OUTPUT_SPEC:
        r = write_row(ws, r, [k, '', v, '', '', '', '', ''])
        ws.merge_cells(start_row=r - 1, start_column=1, end_row=r - 1, end_column=2)
        ws.merge_cells(start_row=r - 1, start_column=3, end_row=r - 1, end_column=8)
    return ws


# ======================================================================
# シート3 : テストケース一覧 / 测试用例一览
# ======================================================================
# (No, 分類, 分类, 対象項目, 観点(日), 观点(中), テストデータ, 期待結果(日), 期待结果(中), 生成方法)
CASES = [
    ('TC-001', '正常', '正常', '全体 / 整体', '必須3項目のみのレコードを取り込む', '仅导入含3个必填项目的记录',
     'No.1/No.3/No.13のみ設定', '正常に取り込まれる', '正常导入', '-DataProfile minimal'),
    ('TC-002', '正常', '正常', '全体 / 整体', '全13項目を設定したレコードを取り込む', '导入设置了全部13项的记录',
     '13項目すべてに有効値', '正常に取り込まれる', '正常导入', '-DataProfile full'),
    ('TC-003', '正常', '正常', '全体 / 整体', 'タブ区切り13項目／1レコードの形式', '制表符分隔13列／1条记录的格式',
     '各行がTABで13分割される', '正常に取り込まれる', '正常导入', '自動 / 自动'),
    ('TC-004', '異常', '异常', '全体 / 整体', '項目数が13以外', '列数不是13',
     '12項目／14項目の行', '形式エラーとなり取込不可', '格式错误，无法导入', '手動編集 / 手动编辑'),
    ('TC-005', '正常', '正常', 'No.1 レコードID', '最大桁数30桁ちょうど', '恰好30位（最大位数）',
     '30桁の英数字', '正常に取り込まれる', '正常导入', '-IncludeBoundary'),
    ('TC-006', '境界', '边界', 'No.1 レコードID', '最大桁数超過（31桁）', '超过最大位数（31位）',
     '31桁の英数字', '桁数エラーとなり取込不可', '位数错误，无法导入', '手動編集 / 手动编辑'),
    ('TC-007', '異常', '异常', 'No.1 レコードID', '必須項目の未設定', '必填项目未设置',
     'レコードIDが空', '必須エラーとなり取込不可', '必填错误，无法导入', '手動編集 / 手动编辑'),
    ('TC-008', '異常', '异常', 'No.1 レコードID', '同一ファイル内での重複', '同一文件内重复',
     '同一レコードIDの行を2件', 'エラーとなり取込不可（設計書 備考）', '报错，无法导入（规格书备注）', '手動編集 / 手动编辑'),
    ('TC-009', '正常', '正常', 'No.1 レコードID', '後続スペースのカット', '尾部空格被去除',
     '「0000000001␣␣␣」', '後続スペースがカットされ有効桁のみ取込', '去除尾部空格后仅导入有效位', '手動編集 / 手动编辑'),
    ('TC-010', '正常', '正常', 'No.1 レコードID', '検索対象者は自動採番のため未設定可', '检索对象者可不设置（自动采番）',
     '検索対象者ファイルでレコードID空', '自動採番され正常に取り込まれる', '自动采番并正常导入',
     'setRecordIdForSearch=false'),
    ('TC-011', '正常', '正常', 'No.2 タイプ', '個人区分「I」', '个人区分 "I"',
     '"I"', '個人として取り込まれる', '作为个人导入', '-DataProfile full'),
    ('TC-012', '正常', '正常', 'No.2 タイプ', '個人以外（空）', '非个人（空）',
     '""', '個人以外として取り込まれる', '作为非个人导入', '-DataProfile minimal'),
    ('TC-013', '異常', '异常', 'No.2 タイプ', '最大桁数超過（2桁）', '超过最大位数（2位）',
     '"IX"', '桁数エラーとなり取込不可', '位数错误，无法导入', '手動編集 / 手动编辑'),
    ('TC-014', '異常', '异常', 'No.3/4/5 名称', 'フリガナ・漢字・英字がすべて未設定', '片假名・汉字・英文全部未设置',
     '3項目とも空', '必須エラーとなり取込不可（設計書 備考）', '必填错误，无法导入（规格书备注）', '手動編集 / 手动编辑'),
    ('TC-015', '正常', '正常', 'No.3 名称（フリガナ）', '最大桁数200文字ちょうど', '恰好200字（最大位数）',
     '全角カタカナ200文字', '正常に取り込まれる', '正常导入', '-IncludeBoundary'),
    ('TC-016', '境界', '边界', 'No.3 名称（フリガナ）', '最大桁数超過（201文字）', '超过最大位数（201字）',
     '全角カタカナ201文字', '桁数エラーとなり取込不可', '位数错误，无法导入', '手動編集 / 手动编辑'),
    ('TC-017', '正常', '正常', 'No.3 名称（フリガナ）', '半角スペースの入力', '输入半角空格',
     '「ヤマダ␣タロウ」', '正常に取り込まれる', '正常导入', '自動 / 自动'),
    ('TC-018', '正常', '正常', 'No.3 名称（フリガナ）', 'ファイル内で重複しない', '文件内不重复',
     '全レコードでフリガナが一意', '正常に取り込まれる', '正常导入', '自動 / 自动'),
    ('TC-019', '異常', '异常', 'No.3 名称（フリガナ）', '文字種違反（ひらがな・漢字混在）', '字符种违规（混入平假名・汉字）',
     '「やまだ　太郎」', '文字種エラーとなり取込不可', '字符种错误，无法导入', '手動編集 / 手动编辑'),
    ('TC-020', '正常', '正常', 'No.4 名称（漢字）', '最大桁数200文字ちょうど', '恰好200字（最大位数）',
     '全角漢字200文字', '正常に取り込まれる', '正常导入', '-IncludeBoundary'),
    ('TC-021', '正常', '正常', 'No.5 名称（英字）', '最大桁数200文字＋半角スペース', '恰好200字＋半角空格',
     '半角英字200文字', '正常に取り込まれる', '正常导入', '-IncludeBoundary'),
    ('TC-022', '正常', '正常', 'No.6 電話番号', '最大桁数30桁ちょうど', '恰好30位（最大位数）',
     '半角数字30桁', '正常に取り込まれる', '正常导入', '-IncludeBoundary'),
    ('TC-023', '正常', '正常', 'No.7 ID情報', '最大桁数30桁ちょうど', '恰好30位（最大位数）',
     '半角英数30桁', '正常に取り込まれる', '正常导入', '-IncludeBoundary'),
    ('TC-024', '正常', '正常', 'No.8 ID情報種類', '固定2桁の有効コード値', '固定2位的有效代码值',
     '"01"（運転免許証）等', '正常に取り込まれる', '正常导入', '-DataProfile full'),
    ('TC-025', '異常', '异常', 'No.8 ID情報種類', '固定桁数違反（1桁／3桁）', '固定位数违规（1位／3位）',
     '"1" ／ "011"', '桁数エラーとなり取込不可', '位数错误，无法导入', '手動編集 / 手动编辑'),
    ('TC-026', '正常', '正常', 'No.9 国内郵便番号', '数値7桁（ハイフンなし）', '数值7位（无连字符）',
     '"2200011"', '正常に取り込まれる', '正常导入', '-DataProfile full'),
    ('TC-027', '異常', '异常', 'No.9 国内郵便番号', 'ハイフンあり', '含连字符',
     '"220-0011"', '文字種・桁数エラーとなり取込不可', '字符种・位数错误，无法导入', '手動編集 / 手动编辑'),
    ('TC-028', '異常', '异常', 'No.9 国内郵便番号', '固定桁数違反（6桁）', '固定位数违规（6位）',
     '"220001"', '桁数エラーとなり取込不可', '位数错误，无法导入', '手動編集 / 手动编辑'),
    ('TC-029', '正常', '正常', 'No.10 国内住所', '市区町村以降200文字ちょうど', '市区町村以后恰好200字',
     '全角住所200文字', '正常に取り込まれる', '正常导入', '-IncludeBoundary'),
    ('TC-030', '正常', '正常', 'No.11 英文住所', '最大桁数256文字＋半角スペース', '恰好256字＋半角空格',
     '半角英数256文字', '正常に取り込まれる', '正常导入', '-IncludeBoundary'),
    ('TC-031', '異常', '异常', 'No.11 英文住所', '国名を記入した場合', '填写了国名的情况',
     '"..., Yokohama, JAPAN"', 'リストの都市・住所と照合できない', '无法与列表的城市・地址匹配', '手動編集 / 手动编辑'),
    ('TC-032', '正常', '正常', 'No.12 設立/生年月日', '数値8桁 YYYYMMDD', '数值8位 YYYYMMDD',
     '"19800101"', 'リストの生年月日と照合される', '与列表的出生日期进行匹配', '-DataProfile full'),
    ('TC-033', '異常', '异常', 'No.12 設立/生年月日', '存在しない日付', '不存在的日期',
     '"20260230"', '日付エラーとなり取込不可', '日期错误，无法导入', '手動編集 / 手动编辑'),
    ('TC-034', '異常', '异常', 'No.12 設立/生年月日', '固定桁数違反（7桁／9桁）', '固定位数违规（7位／9位）',
     '"1980010" ／ "198001011"', '桁数エラーとなり取込不可', '位数错误，无法导入', '手動編集 / 手动编辑'),
    ('TC-035', '異常', '异常', 'No.13 部店コード', '必須項目の未設定', '必填项目未设置',
     '部店コードが空', '必須エラーとなり取込不可', '必填错误，无法导入', '手動編集 / 手动编辑'),
    ('TC-036', '正常', '正常', 'No.13 部店コード', '検索対象者の部店コード', '检索对象者的部店代码',
     '横浜銀行用=0003 ／ L&F用=LF002', '正常に取り込まれる', '正常导入', '対象 S01 / S02'),
    ('TC-037', '正常', '正常', 'No.13 部店コード', 'CIFのファイル別部店コード', 'CIF 按文件区分的部店代码',
     '001=0005 ／ 002=LF003 ／ 003=LF004', '正常に取り込まれる', '正常导入', '対象 001 / 002 / 003'),
    ('TC-038', '正常', '正常', 'No.13 部店コード', '最大桁数32桁ちょうど', '恰好32位（最大位数）',
     '半角英数32桁', '正常に取り込まれる', '正常导入', '-IncludeBoundary'),
    ('TC-039', '正常', '正常', 'ファイル名 / 文件名', '*001.txt / *002.txt / *003.txt の命名規則', '文件命名规则',
     'CIF_yyyymmdd_001.txt 等', '対応する部店コードで正常に取り込まれる', '按对应的部店代码正常导入', '自動 / 自动'),
    ('TC-040', '正常', '正常', '文字コード / 字符编码', 'Shift_JIS（CP932）／ CRLF', 'Shift_JIS（CP932）／ CRLF',
     '全角漢字・カタカナを含むレコード', '文字化けせず正常に取り込まれる', '不乱码，正常导入', 'encoding=shift_jis'),
    ('TC-041', '正常', '正常', '照合 / 匹配', 'フリガナ・漢字・生年月日すべて一致', '片假名・汉字・出生日期全部匹配',
     'レコードID 1A2A3Ah', '検索対象者とヒットする', '与检索对象者命中', '-DataProfile matching'),
    ('TC-042', '正常', '正常', '照合 / 匹配', '漢字・生年月日が未設定でも一致', '汉字・出生日期未设置时也匹配',
     'レコードID 1A2C3Ch', '検索対象者とヒットする', '与检索对象者命中', '-DataProfile matching'),
    ('TC-043', '正常', '正常', '照合 / 匹配', '漢字が不一致の場合', '汉字不匹配的情况',
     'レコードID 1A2B3A', 'ヒットしない', '不命中', '-DataProfile matching'),
    ('TC-044', '正常', '正常', '照合 / 匹配', '生年月日が不一致の場合', '出生日期不匹配的情况',
     'レコードID 1A2A3B', 'ヒットしない', '不命中', '-DataProfile matching'),
    ('TC-045', '正常', '正常', '照合 / 匹配', 'フリガナが不一致の場合', '片假名不匹配的情况',
     'レコードID 1B2A3A', 'ヒットしない', '不命中', '-DataProfile matching'),
]

CAT_FILL = {
    '正常': PatternFill('solid', fgColor='E2EFDA'),
    '境界': PatternFill('solid', fgColor='FFF2CC'),
    '異常': PatternFill('solid', fgColor='FCE4EC'),
}


def build_sheet_cases(wb):
    ws = wb.create_sheet('3.テストケース一覧')
    set_widths(ws, [9, 7, 7, 20, 30, 30, 26, 30, 30, 22, 8, 10, 18])
    r = sheet_title(ws, 1, 'テストケース一覧', '测试用例一览', 13)

    c = ws.cell(row=r, column=1,
                value='※「生成方法」が 自動 / -DataProfile / -IncludeBoundary のケースは本ツールで生成可能。'
                      '「手動編集」は生成後のファイルを編集して異常系データを作成する。\n'
                      '※"生成方法"为 自动 / -DataProfile / -IncludeBoundary 的用例可由本工具生成；'
                      '"手动编辑"需在生成后编辑文件以制作异常数据。')
    c.font = FONT_NOTE
    c.alignment = AL_TL
    ws.row_dimensions[r].height = 28
    r += 2

    head_row = r
    r = write_row(ws, r, ['No.', '分類', '分类', '対象項目 / 对象项目', 'テスト観点', '测试观点',
                          'テストデータ / 测试数据', '期待結果', '预期结果', '生成方法 / 生成方式',
                          '結果\n结果', '実施日\n实施日', '備考 / 备注'],
                  font=FONT_HEAD, fill=FILL_HEAD, align=AL_CT, height=30, cn_cols=(3, 6, 9))
    for case in CASES:
        r = write_row(ws, r, list(case) + ['', '', ''], height=24, cn_cols=(3, 6, 9))
        ws.cell(row=r - 1, column=1).alignment = AL_CT
        for col in (2, 3):
            cell = ws.cell(row=r - 1, column=col)
            cell.alignment = AL_CT
            cell.fill = CAT_FILL.get(case[1], FILL_SUB)
        for col in (11, 12):
            ws.cell(row=r - 1, column=col).alignment = AL_CT

    ws.freeze_panes = ws.cell(row=head_row + 1, column=4)
    ws.auto_filter.ref = f'A{head_row}:M{r - 1}'

    r += 1
    counts = {k: sum(1 for x in CASES if x[1] == k) for k in ('正常', '境界', '異常')}
    c = ws.cell(row=r, column=1,
                value=f'合計 {len(CASES)} ケース（正常 {counts["正常"]} ／ 境界 {counts["境界"]} ／ 異常 {counts["異常"]}）'
                      f'　合计 {len(CASES)} 条（正常 {counts["正常"]} ／ 边界 {counts["境界"]} ／ 异常 {counts["異常"]}）')
    c.font = FONT_HEAD
    return ws


# ======================================================================
# シート4 : 実行手順 / 执行步骤
# ======================================================================
STEPS = [
    ('1', '事前準備 / 事前准备', 'Amazon WorkSpaces にログインする', '登录 Amazon WorkSpaces',
     'WorkSpaces クライアントまたは Web Access から接続。'),
    ('2', '事前準備 / 事前准备', 'ツール一式を作業フォルダへ配置する', '将工具整套放到工作目录',
     '例）D:\\work\\cif-testdata\\ に run_generate.bat / scripts / config / samples を配置。'
     'D ドライブ（ユーザーボリューム）は再起動後も保持される。'),
    ('3', '事前準備 / 事前准备', 'ブロック解除を確認する', '确认已解除文件阻止（Mark of the Web）',
     'run_generate.bat が起動時に Unblock-File を自動実行する。'),
    ('4', '事前準備 / 事前准备', 'PowerShell 実行ポリシーを確認する', '确认 PowerShell 执行策略',
     'run_generate.bat は -ExecutionPolicy Bypass で起動するため端末側の変更は不要。'),
    ('5', '設定 / 设置', 'config\\config.json を編集する', '编辑 config\\config.json',
     '件数（count）、部店コード（butenCode）、文字コード（encoding）、プロファイル（profile）を設定。'),
    ('6', '設定 / 设置', '部店コードをヒアリングシートと突き合わせる', '将部店代码与访谈表核对',
     '001→0005 / 002→LF003 / 003→LF004 / 検索対象者→0003・LF002。'
     '※実データ画面では 0011 が使われていたため要確認。'),
    ('7', '生成 / 生成', 'run_generate.bat をダブルクリックする', '双击 run_generate.bat',
     'メニューから [1]標準／[2]全項目／[3]全項目＋境界値／[4]照合パターン／[5]件数指定 を選択。'),
    ('8', '生成 / 生成', '「結果 : OK」を確認する', '确认显示 "結果 : OK"',
     '1件でも項目定義違反があると「結果 : NG」となり終了コード1で終了する。'),
    ('9', '確認 / 确认', 'output\\ の生成ファイルを確認する', '确认 output\\ 下的生成文件',
     'CIF_yyyymmdd_001/002/003.txt、SEARCH_yyyymmdd_S01/S02.txt。'),
    ('10', '確認 / 确认', 'logs\\ のログを保管する', '保存 logs\\ 下的日志',
     'generate_yyyyMMdd_HHmmss.log をエビデンスとして課題管理へ添付する。'),
    ('11', '確認 / 确认', '文字コード・改行コードを確認する', '确认字符编码与换行符',
     'サクラエディタ等で Shift_JIS・CRLF・タブ区切りであることを確認。'),
    ('12', '確認 / 确认', '参考データと見比べる', '与参考数据比对',
     'samples\\ の参考データ（照合パターン18件）と項目位置・値を突き合わせる。シート5参照。'),
    ('13', '試験 / 测试', '異常系データを手動作成する', '手动制作异常数据',
     'シート3で「手動編集」となっているケースについて、生成済みファイルをコピーして書き換える。'),
    ('14', '試験 / 测试', '手動編集後のファイルを再検証する', '重新校验手动编辑后的文件',
     'run_generate.bat -ValidateOnly で想定どおりのエラーが検出されることを確認。'),
    ('15', '試験 / 测试', '対象システムへ取り込み結果を突き合わせる', '导入目标系统并核对结果',
     'シート3の「結果」「実施日」欄に記入する。'),
    ('16', '後片付け / 收尾', 'テストデータを削除する', '删除测试数据',
     '試験完了後、output\\ 配下のファイルを削除する。'),
]

TROUBLES = [
    ('コンソールの日本語が化ける / 控制台日文乱码',
     'コンソールフォントを「ＭＳ ゴシック」等に変更する。/ 将控制台字体改为「ＭＳ ゴシック」等。'),
    ('生成ファイルが化ける / 生成文件乱码',
     'config.json の output.encoding を取込システムの仕様に合わせる。/ 使其与导入系统的规格一致。'),
    ('「このスクリプトは実行できません」/ 提示无法执行脚本',
     'run_generate.bat 経由で起動する（-ExecutionPolicy Bypass 付き）。/ 请通过 run_generate.bat 启动。'),
    ('全角文字で桁数エラー / 全角字符导致位数错误',
     'generation.lengthCheckMode を "byte" に変更する。/ 将其改为 "byte"（按 Shift_JIS 字节数判定）。'),
    ('項目数エラーになる / 提示列数错误',
     '実データが13項目でない場合は extraColumns を有効化する。/ 若实际数据不是13列，请启用 extraColumns。'),
    ('照合パターンでフリガナ重複エラー / 匹配模式下片假名重复报错',
     'matching プロファイルでは自動的に重複チェックを無効化する。/ matching 配置下会自动关闭重复校验。'),
    ('WorkSpace 再起動でファイルが消えた / 重启后文件消失',
     'C ドライブは再構築で初期化される。D ドライブ配下に配置する。/ C 盘会被重建初始化，请放在 D 盘。'),
]


def build_sheet_steps(wb):
    ws = wb.create_sheet('4.実行手順')
    set_widths(ws, [6, 20, 38, 34, 62, 12])
    r = sheet_title(ws, 1, '実行手順チェックリスト（Amazon WorkSpaces）', '执行步骤检查表（Amazon WorkSpaces）', 6)

    head_row = r
    r = write_row(ws, r, ['No.', '区分 / 区分', '作業内容', '作业内容(中文)', '補足 / 补充', 'チェック\n检查'],
                  font=FONT_HEAD, fill=FILL_HEAD, align=AL_CT, height=26, cn_cols=(4,))
    for s in STEPS:
        r = write_row(ws, r, list(s) + [''], height=28, cn_cols=(4,))
        for col in (1, 2, 6):
            ws.cell(row=r - 1, column=col).alignment = AL_CT
    ws.freeze_panes = ws.cell(row=head_row + 1, column=1)

    r += 2
    ws.cell(row=r, column=1, value='【トラブルシューティング / 故障排查】').font = FONT_HEAD
    r += 1
    r = write_row(ws, r, ['事象 / 现象', '', '', '対処 / 处理', '', ''],
                  font=FONT_HEAD, fill=FILL_HEAD, align=AL_CT)
    ws.merge_cells(start_row=r - 1, start_column=1, end_row=r - 1, end_column=3)
    ws.merge_cells(start_row=r - 1, start_column=4, end_row=r - 1, end_column=6)
    for t in TROUBLES:
        r = write_row(ws, r, [t[0], '', '', t[1], '', ''], height=26)
        ws.merge_cells(start_row=r - 1, start_column=1, end_row=r - 1, end_column=3)
        ws.merge_cells(start_row=r - 1, start_column=4, end_row=r - 1, end_column=6)
    return ws


# ======================================================================
# シート5 : 参考データ / 参考数据
# ======================================================================
FURIGANA = {'A': 'アイダマナ', 'B': 'アイタナナ'}
KANJI = {'A': '相田　マナ', 'B': '相他ナナ', 'C': ''}
BIRTHDAY = {'A': '20000101', 'B': '20010101', 'C': ''}
BUTEN_REF = '0011'

LABEL1 = {'A': '一致 / 匹配', 'B': '不一致 / 不匹配'}
LABEL = {'A': '一致 / 匹配', 'B': '不一致 / 不匹配', 'C': '未設定 / 未设置'}

REF_FILES = [
    ('参考データ_照合パターン_13項目_SJIS.txt', 'Shift_JIS', '13',
     '設計書 No.1〜13 のみ。生成ツールの出力とバイト単位で一致。/ 仅规格书 No.1〜13，与生成工具输出逐字节一致。'),
    ('参考データ_照合パターン_13項目_UTF8.txt', 'UTF-8', '13',
     '内容は同上。エディタ確認用。/ 内容同上，供编辑器查看。'),
    ('参考データ_照合パターン_実データ再現_23項目_SJIS.txt', 'Shift_JIS', '23',
     '実データ画面の再現。No.14以降は推定値。/ 复现实际数据画面，No.14 之后为推定值。'),
    ('参考データ_パターン対応表.tsv', 'UTF-8 (BOM)', '10',
     'レコードIDと期待結果の対応表。Excelで開ける。/ 记录ID与预期结果对照表，可用 Excel 打开。'),
]

REF_LAYOUT = [
    ('No.1', 'レコードID / 记录ID', '1A2A3Ah', '設計書 / 规格书'),
    ('No.2', 'タイプ / 类型', '（空）', '設計書 / 规格书'),
    ('No.3', '名称（フリガナ）/ 名称（片假名）', 'アイダマナ', '設計書 / 规格书'),
    ('No.4', '名称（漢字）/ 名称（汉字）', '相田　マナ', '設計書 / 规格书'),
    ('No.5〜11', '英字名・電話番号・ID情報・ID種類・郵便番号・国内住所・英文住所', '（空）', '設計書 / 规格书'),
    ('No.12', '設立/生年月日 / 出生年月日', '20000101', '設計書 / 规格书'),
    ('No.13', '部店コード / 部店代码', '0011', '設計書 / 规格书'),
    ('No.14〜21', '未確定 / 未确定', '（空）', '★推定 / 推定'),
    ('No.22', '未確定（実データでは契約番号形式）/ 未确定（实际数据为合同号格式）', 'A100000001-01', '★推定 / 推定'),
    ('No.23', '未確定（実データでは1桁フラグ）/ 未确定（实际数据为1位标志）', '0', '★推定 / 推定'),
]


def build_matching_rows():
    rows = []
    for f1 in ('A', 'B'):
        for f3 in ('A', 'B', 'C'):
            for f2 in ('A', 'B', 'C'):
                hit = (f1 == 'A') and (f2 != 'B') and (f3 != 'B')
                rec_id = f'1{f1}2{f2}3{f3}' + ('h' if hit else '')
                rows.append((f1, f2, f3, hit, rec_id))
    return rows


def build_sheet_reference(wb):
    ws = wb.create_sheet('5.参考データ')
    set_widths(ws, [6, 16, 20, 20, 20, 15, 18, 16, 14, 14])
    r = sheet_title(ws, 1, '参考データ（照合パターン）', '参考数据（匹配模式）', 10)

    r = note_lines(ws, r, [
        ('出典：実データ画面（サクラエディタ / O:\\work\\NDLC\\yyyymmdd\\...*002.txt）', FONT_BASE),
        ('出处：实际数据画面（サクラエディタ / O:\\work\\NDLC\\yyyymmdd\\...*002.txt）', FONT_CN),
        ('', FONT_BASE),
        ('【レコードIDの命名規則 / 记录ID命名规则】  1<f1>2<f2>3<f3>[h]', FONT_HEAD),
        ('  f1 = 名称（フリガナ）  A：一致 / B：不一致　　　　　　　片假名名称  A：匹配 / B：不匹配', FONT_BASE),
        ('  f2 = 名称（漢字）　　  A：一致 / B：不一致 / C：未設定　汉字名称　  A：匹配 / B：不匹配 / C：未设置', FONT_BASE),
        ('  f3 = 設立/生年月日　　 A：一致 / B：不一致 / C：未設定　出生年月日  A：匹配 / B：不匹配 / C：未设置', FONT_BASE),
        ('  h  = ヒット想定（フリガナ一致 かつ 漢字が不一致でない かつ 生年月日が不一致でない）', FONT_BASE),
        ('       预期命中（片假名匹配 且 汉字非不匹配 且 出生日期非不匹配）', FONT_CN),
        ('  並び順 / 排列顺序： f1 → f3 → f2 の順にループ（実データと同一）/ 按 f1 → f3 → f2 循环（与实际数据一致）', FONT_BASE),
        ('', FONT_BASE),
    ])

    ws.cell(row=r, column=1, value='【照合パターン一覧 2×3×3＝18件 / 匹配模式一览 2×3×3＝18条】').font = FONT_HEAD
    r += 1
    head_row = r
    r = write_row(ws, r, ['No.', 'レコードID\n记录ID', 'f1 フリガナ / 片假名', 'f2 漢字 / 汉字',
                          'f3 生年月日 / 出生日期', 'ヒット想定\n预期命中',
                          'No.3 名称（フリガナ）', 'No.4 名称（漢字）', 'No.12 生年月日', 'No.13 部店コード'],
                  font=FONT_HEAD, fill=FILL_HEAD, align=AL_CT, height=30)

    for i, (f1, f2, f3, hit, rec_id) in enumerate(build_matching_rows(), start=1):
        vals = [i, rec_id, f'{f1}：{LABEL1[f1]}', f'{f2}：{LABEL[f2]}', f'{f3}：{LABEL[f3]}',
                'ヒット / 命中' if hit else '−',
                FURIGANA[f1], KANJI[f2] or '（未設定）', BIRTHDAY[f3] or '（未設定）', BUTEN_REF]
        r = write_row(ws, r, vals)
        for col in (1, 2, 6, 9, 10):
            ws.cell(row=r - 1, column=col).alignment = AL_CT
        if hit:
            for col in range(1, 11):
                ws.cell(row=r - 1, column=col).fill = FILL_HIT
    ws.freeze_panes = ws.cell(row=head_row + 1, column=1)

    r += 2
    ws.cell(row=r, column=1, value='【実データの項目レイアウト / 实际数据的列布局】').font = FONT_HEAD
    r += 1
    r = write_row(ws, r, ['列 / 列', '項目 / 项目', '', '', '実データ例 / 实际数据示例', '', '出典 / 出处', '', '', ''],
                  font=FONT_HEAD, fill=FILL_HEAD, align=AL_CT)
    ws.merge_cells(start_row=r - 1, start_column=2, end_row=r - 1, end_column=4)
    ws.merge_cells(start_row=r - 1, start_column=5, end_row=r - 1, end_column=6)
    ws.merge_cells(start_row=r - 1, start_column=7, end_row=r - 1, end_column=10)
    for col, name, sample, src in REF_LAYOUT:
        r = write_row(ws, r, [col, name, '', '', sample, '', src, '', '', ''], height=22)
        ws.merge_cells(start_row=r - 1, start_column=2, end_row=r - 1, end_column=4)
        ws.merge_cells(start_row=r - 1, start_column=5, end_row=r - 1, end_column=6)
        ws.merge_cells(start_row=r - 1, start_column=7, end_row=r - 1, end_column=10)
        ws.cell(row=r - 1, column=1).alignment = AL_CT
        if src.startswith('★'):
            for c2 in range(1, 11):
                ws.cell(row=r - 1, column=c2).fill = FILL_EST

    r += 1
    c = ws.cell(row=r, column=1,
                value='★ 設計書原本の No.14 以降が未入手のため、実データ画面から推定した値です。'
                      '正式な項目定義を入手したら config.json の extraColumns と本シートを更新してください。\n'
                      '★ 因未获取规格书原件的 No.14 之后部分，此为根据实际数据画面推定的值。'
                      '取得正式项目定义后，请更新 config.json 的 extraColumns 与本 sheet。')
    c.font = FONT_NOTE
    c.alignment = AL_TL
    ws.row_dimensions[r].height = 30
    r += 2

    ws.cell(row=r, column=1, value='【参考データファイル / 参考数据文件】 samples\\ 配下').font = FONT_HEAD
    r += 1
    r = write_row(ws, r, ['ファイル名 / 文件名', '', '', '文字コード\n字符编码', '項目数\n列数',
                          '内容 / 内容', '', '', '', ''],
                  font=FONT_HEAD, fill=FILL_HEAD, align=AL_CT, height=26)
    ws.merge_cells(start_row=r - 1, start_column=1, end_row=r - 1, end_column=3)
    ws.merge_cells(start_row=r - 1, start_column=6, end_row=r - 1, end_column=10)
    for name, enc, cols, desc in REF_FILES:
        r = write_row(ws, r, [name, '', '', enc, cols, desc, '', '', '', ''], height=24)
        ws.merge_cells(start_row=r - 1, start_column=1, end_row=r - 1, end_column=3)
        ws.merge_cells(start_row=r - 1, start_column=6, end_row=r - 1, end_column=10)
        for col in (4, 5):
            ws.cell(row=r - 1, column=col).alignment = AL_CT

    r += 2
    note_lines(ws, r, [
        ('【再生成コマンド / 重新生成命令】', FONT_HEAD),
        ('  python scripts\\build_reference_samples.py            … samples\\ の参考データを再生成 / 重新生成参考数据', FONT_BASE),
        ('  run_generate.bat -DataProfile matching -Target 002   … 同じ18件を output\\ に生成 / 在 output\\ 生成相同的18条', FONT_BASE),
        ('  ※ 部店コードは config.json の matching.butenCodeOverride（既定 0011）で指定', FONT_BASE),
        ('  ※ 部店代码由 config.json 的 matching.butenCodeOverride（默认 0011）指定', FONT_CN),
    ])
    return ws


# ======================================================================
# シート0 : 表紙 / 封面
# ======================================================================
def build_sheet_cover(wb):
    ws = wb.create_sheet('0.表紙', 0)
    set_widths(ws, [4, 24, 22, 46, 16])
    ws.cell(row=2, column=2, value='CIF情報ファイル　テストデータ設計書').font = FONT_TITLE
    ws.cell(row=3, column=2, value='CIF信息文件　测试数据设计书').font = FONT_CN

    r = 6
    for k, kc, v in [
        ('文書名', '文档名', 'CIF情報ファイル テストデータ設計書 / CIF信息文件 测试数据设计书'),
        ('版数', '版本', DOC_VERSION),
        ('作成日', '创建日', DOC_DATE),
        ('作成', '编制', '（記入してください / 请填写）'),
        ('承認', '批准', '（記入してください / 请填写）'),
        ('対象システム', '目标系统', 'CIF情報ファイル取込機能 / CIF信息文件导入功能'),
        ('実行環境', '执行环境', 'Amazon WorkSpaces（Windows / PowerShell 5.1以上）'),
        ('関連資材', '相关资料', 'run_generate.bat / scripts\\ / config\\config.json / samples\\ / README.md'),
    ]:
        for i, (val, font) in enumerate([(k, FONT_HEAD), (kc, FONT_CN), (v, FONT_BASE)]):
            c = ws.cell(row=r, column=2 + i, value=val)
            c.font = font
            c.border = BORDER
            c.alignment = AL_TL
            if i < 2:
                c.fill = FILL_SUB
        r += 1

    r += 2
    ws.cell(row=r, column=2, value='【シート構成 / sheet 构成】').font = FONT_HEAD
    r += 1
    for k, v in [
        ('1.項目定義 / 项目定义', '設計書「CIF情報ファイルの項目」の転記（No.1〜13）＋文字種定義・ファイル形式'),
        ('2.生成条件 / 生成条件', '生成対象ファイル、データプロファイル、項目別生成ルール、出力仕様'),
        ('3.テストケース一覧 / 测试用例一览', '正常／境界／異常／照合系のケースと期待結果、結果記入欄'),
        ('4.実行手順 / 执行步骤', 'Amazon WorkSpaces 上の作業チェックリスト＋トラブルシューティング'),
        ('5.参考データ / 参考数据', '実データ画面から再現した照合パターン18件と項目レイアウト'),
    ]:
        c1 = ws.cell(row=r, column=2, value=k)
        c1.font = FONT_BASE
        c1.border = BORDER
        c1.alignment = AL_TL
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        c2 = ws.cell(row=r, column=4, value=v)
        c2.font = FONT_BASE
        c2.border = BORDER
        c2.alignment = AL_TL
        r += 1

    r += 2
    ws.cell(row=r, column=2, value='【改訂履歴 / 修订履历】').font = FONT_HEAD
    r += 1
    for i, v in enumerate(['版数 / 版本', '改訂日 / 修订日', '改訂内容 / 修订内容', '改訂者 / 修订者']):
        c = ws.cell(row=r, column=2 + i, value=v)
        c.font = FONT_HEAD
        c.fill = FILL_HEAD
        c.border = BORDER
        c.alignment = AL_CT
    r += 1
    for ver, date, desc in [
        ('1.0', DOC_DATE, '初版作成（設計書「CIF情報ファイルの項目」に基づく）/ 初版（基于规格书）'),
        ('2.0', DOC_DATE, '日中対照化、シート5「参考データ」追加、照合系ケース TC-041〜045 追加 / '
                          '改为中日对照，新增 sheet5「参考数据」与匹配用例'),
    ]:
        for i, v in enumerate([ver, date, desc, '']):
            c = ws.cell(row=r, column=2 + i, value=v)
            c.font = FONT_BASE
            c.border = BORDER
            c.alignment = AL_TL
        r += 1
    return ws


# ======================================================================
def main():
    root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    out_dir = os.path.join(root, 'docs')
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, 'CIF情報ファイル_テストデータ設計書.xlsx')

    wb = Workbook()
    wb.remove(wb.active)
    build_sheet_fields(wb)
    build_sheet_conditions(wb)
    build_sheet_cases(wb)
    build_sheet_steps(wb)
    build_sheet_reference(wb)
    build_sheet_cover(wb)

    for ws in wb.worksheets:
        ws.sheet_view.showGridLines = False
    wb.active = 0

    try:
        wb.save(out_path)
    except PermissionError:
        # Excel で開かれている場合は別名で保存する / 若文件正被 Excel 打开，则另存为
        alt = out_path.replace('.xlsx', '_new.xlsx')
        wb.save(alt)
        print('[WARN] 対象ファイルが Excel で開かれているため別名で保存しました。')
        print('[WARN] 目标文件正被 Excel 打开，已另存为其他文件名。')
        print(f'       -> {alt}')
        print('       Excel を閉じてから再実行すると正規のファイル名で上書きされます。')
        print('       关闭 Excel 后重新执行，即可覆盖为正式文件名。')
        out_path = alt

    counts = {k: sum(1 for x in CASES if x[1] == k) for k in ('正常', '境界', '異常')}
    print(f'[OK] 生成しました / 已生成: {out_path}')
    print(f'     シート / sheet: {[ws.title for ws in wb.worksheets]}')
    print(f'     項目定義 {len(FIELDS)} 項目 / テストケース {len(CASES)} 件 '
          f'(正常{counts["正常"]} 境界{counts["境界"]} 異常{counts["異常"]}) / 照合パターン 18 件')


if __name__ == '__main__':
    main()
