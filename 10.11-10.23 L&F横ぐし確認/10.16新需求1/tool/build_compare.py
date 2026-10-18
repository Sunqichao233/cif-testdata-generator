# -*- coding: utf-8 -*-
"""
insert.sql と マスタ定義書（横浜銀行様用出荷設定）を横並びにした対比表を作る。

  シート構成 : insert.sql のコメント「-- N_xxx」単位で1シート
  各シートの並び
      原文SQL
      カンマ区切り（サンプル様式）
      ① insert.sql の値          ← 列名は定義書から取る
      ② 定義書 横浜銀行様用出荷設定   ← 同じ列名
      突合（判定／相違列／定義書の行）
      定義書のSQL（固定・固定・固定・本文・固定）
      再構築・検証（insert側）
      確認結果（判定・備考の記入欄）

使い方 : python build_compare.py insert.sql マスタ定義書.xlsx 出力.xlsx
"""
import os, re, sys, datetime, warnings
warnings.filterwarnings('ignore')
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as CL

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from sql2excel import load, to_value, Q, FONT, SZ, sheet_name

C_RAW     = 'FFD9D9D9'
C_SPLIT   = 'FFDDEBF7'
C_INS     = 'FFFFFF99'   # ① insert の値（サンプルと同じ色）
C_DEF     = 'FFDDF0D8'   # ② 定義書の値
C_JDG     = 'FFFFE699'   # 突合
C_DSQL    = 'FFEDEDED'   # 定義書のSQL
C_REBUILD = 'FFE2EFDA'
C_NOTE    = 'FFFCE4D6'
C_OTHER   = 'FFFFF2CC'
C_NG      = 'FFFFC7CE'   # 相違
C_NONE    = 'FFF8CBAD'   # 定義書に該当なし
C_BLUE    = 'FF99CCFF'   # 定義書と同じ：ブロック見出し
C_GREEN   = 'FFCCFFCC'   # 定義書と同じ：列名の行
C_YELLOW  = 'FFFFFF99'   # 定義書と同じ：データ
C_PINK    = 'FFFF99CC'   # 定義書と同じ：SQLブロック
SEPW      = 2.0

LABEL = '横浜銀行様用出荷設定'
thin = Side(style='thin', color='FFBFBFBF')
BOX  = Border(left=thin, right=thin, top=thin, bottom=thin)


# ================= 共通 =================
def put(ws, r, c, v, fill=None, bold=False, wrap=False, align=None):
    cell = ws.cell(row=r, column=c)
    cell.value = v
    cell.font = Font(name=FONT, size=SZ, bold=bold)
    cell.border = BOX
    if fill:
        cell.fill = PatternFill('solid', fgColor=fill)
    cell.alignment = Alignment(horizontal=align, vertical='center', wrap_text=wrap)
    if isinstance(v, str) and not v.startswith('='):
        cell.number_format = '@'
    return cell


def block_label(ws, r, c1, c2, text, fill):
    for c in range(c1, c2 + 1):
        ws.cell(row=r, column=c).fill = PatternFill('solid', fgColor=fill)
        ws.cell(row=r, column=c).border = BOX
    if c2 > c1:
        ws.merge_cells(start_row=r, start_column=c1, end_row=r, end_column=c2)
    put(ws, r, c1, text, fill=fill, bold=True, align='center')


def nv(v):
    """定義書／insert の値を比べるための文字列化。"""
    if v is None:
        return ''
    if isinstance(v, datetime.datetime):
        return v.strftime('%Y/%m/%d %H:%M:%S')
    if isinstance(v, datetime.date):
        return v.strftime('%Y/%m/%d')
    if isinstance(v, bool):
        return 'TRUE' if v else 'FALSE'
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v).strip()


def nv2(v):
    """SQL本文用。末尾の空白も意味を持つので strip しない。"""
    if v is None:
        return ''
    if isinstance(v, datetime.datetime):
        return v.strftime('%Y/%m/%d %H:%M:%S')
    if isinstance(v, datetime.date):
        return v.strftime('%Y/%m/%d')
    if isinstance(v, bool):
        return 'TRUE' if v else 'FALSE'
    if isinstance(v, float) and v == int(v):
        return str(int(v))
    return str(v)


def nb(s):
    return re.sub(r'\s+', '', nv(s))


# ================= 定義書の読み取り =================
def row1_labels(ws):
    """1行目の見出しを [(列, 見出し)] で返す（結合は先頭列のみ）。"""
    spans = {mr.min_col: mr.max_col for mr in ws.merged_cells.ranges if mr.min_row == 1}
    out = []
    taken = set()
    for c in sorted(spans):
        taken.update(range(c + 1, spans[c] + 1))
    for c in range(1, ws.max_column + 1):
        if c in taken:
            continue
        v = ws.cell(row=1, column=c).value
        if v not in (None, ''):
            out.append((c, nv(v)))
    return out


def block_cols(ws, start, labels):
    """start 列から、2行目が埋まっていて次の見出しに当たるまでを1ブロックとする。"""
    nxt = min([c for c, _ in labels if c > start], default=ws.max_column + 1)
    cols = []
    c = start
    while c < nxt and c <= ws.max_column:
        v = ws.cell(row=2, column=c).value
        if v in (None, ''):
            break
        cols.append(c)
        c += 1
    return cols


def is_red(ws, r, cols):
    """定義書では横浜で追加した行が赤字になっている。その印を拾う。"""
    for c in cols:
        f = ws.cell(row=r, column=c).font
        if f and f.color is not None and f.color.type == 'rgb' and f.color.rgb == 'FFFF0000':
            return True
    return False


def read_def(ws):
    """定義書1シートから「横浜銀行様用出荷設定」とSQLブロックを取り出す。"""
    labels = row1_labels(ws)
    ycol = [c for c, v in labels if v == LABEL]
    scol = [c for c, v in labels if v == 'SQL']
    if not ycol or not scol:
        return None
    ycols = block_cols(ws, ycol[0], labels)
    scols = block_cols(ws, scol[0], labels)
    names = [nv(ws.cell(row=2, column=c).value) for c in ycols]
    fcol = None
    for c in scols:
        if nv(ws.cell(row=2, column=c).value) == '数式':
            fcol = c
    pkg = [c for c, v in labels if v == 'パッケージ標準']
    pkgcols = block_cols(ws, pkg[0], labels) if pkg else []

    rows = []
    for r in range(3, ws.max_row + 1):
        vals = [ws.cell(row=r, column=c).value for c in ycols]
        if all(v in (None, '') for v in vals):
            continue
        rows.append(dict(
            row=r,
            vals=[nv(v) for v in vals],
            sql=[nv2(ws.cell(row=r, column=c).value) for c in scols],
            body=nv2(ws.cell(row=r, column=fcol).value) if fcol else '',
            pkg=any(ws.cell(row=r, column=c).value not in (None, '') for c in pkgcols),
            red=is_red(ws, r, ycols),
        ))
    return dict(names=names, rows=rows, scols=len(scols),
                snames=[nv(ws.cell(row=2, column=c).value) for c in scols])


# ================= 突合 =================
def match(ins, drows):
    """insert の各行に定義書の行を割り当てる。本文一致→先頭列からのキー一致。"""
    used = set()
    link = [None] * len(ins)

    bodyidx = {}
    for i, d in enumerate(drows):
        if d['body']:
            bodyidx.setdefault(nb(d['body']), []).append(i)

    for k, s in enumerate(ins):
        key = nb(s['_body']) if s['_body'] else None
        if not key:
            continue
        for i in bodyidx.get(key, []):
            if i not in used:
                used.add(i); link[k] = i; break

    # 本文で取れなかったものは、先頭 n 列が一致して候補が1つだけなら対応づける
    for k, s in enumerate(ins):
        if link[k] is not None:
            continue
        v = s['_vals']
        for n in range(len(v), 0, -1):
            pre = tuple(v[:n])
            cand = [i for i, d in enumerate(drows)
                    if i not in used and tuple(d['vals'][:n]) == pre]
            if len(cand) == 1:
                used.add(cand[0]); link[k] = cand[0]; break
    return link


# ================= シート作成 =================
def put2(ws, r, c, v, fill=None, bold=False, wrap=False, align=None, red=False):
    cell = ws.cell(row=r, column=c)
    cell.value = v
    cell.font = Font(name=FONT, size=SZ, bold=bold,
                     color='FFFF0000' if red else None)
    cell.border = BOX
    if fill:
        cell.fill = PatternFill('solid', fgColor=fill)
    cell.alignment = Alignment(horizontal=align, vertical='center', wrap_text=wrap)
    if isinstance(v, str) and not v.startswith('='):
        cell.number_format = '@'
    return cell


def head2(ws, c1, hd, c_blk, c_col, label):
    """定義書と同じ2段見出し（1行目＝ブロック名／2行目＝列名）。"""
    block_label(ws, 1, c1, c1 + len(hd) - 1, label, c_blk)
    for i, h in enumerate(hd):
        put2(ws, 2, c1 + i, h, fill=c_col, bold=True, align='center', wrap=True)


def naive_index(args):
    """定義書の列 j が、カンマ区切りの何列目から始まるか。
    to_timestamp は中にカンマを1つ持つので2セルを使う。"""
    out, k = [], 0
    for tok in args:
        out.append(k)
        k += tok.count(',') + 1
    return out


def value_cmp_formula(types, ni, cS0, cD0, r):
    """カンマ区切り（①）と定義書の値（②）を、式1本で突き合わせる。

    ②の値は SQL の書き方に合わせてから比べる（引用符を足す／to_timestamp で包む）。
    to_timestamp の書式文字列は値ではないので、①側も②側も比較に含めない。
    """
    lhs = '&","&'.join('TRIM(%s%d)' % (CL(cS0 + ni[j]), r) for j in range(len(types)))
    parts = []
    for j, t in enumerate(types):
        d = '%s%d' % (CL(cD0 + j), r)
        if t == '日時':
            parts.append('"to_timestamp(%s"&%s&"%s"' % (Q, d, Q))
        elif t in ('文字列', '空'):
            parts.append('"%s"&%s&"%s"' % (Q, d, Q))
        else:
            parts.append(d)
    return '=EXACT(%s,%s)' % (lhs, '&","&'.join(parts))


def build_sheet(wb, sec, dinfo, extra):
    ws = wb.create_sheet(sec['sheet'])
    sts = sec['stmts']
    ins = [s for s in sts if s['kind'] in ('VALUES', 'SELECT')]
    link = match(ins, dinfo['rows']) if dinfo else [None] * len(ins)
    dmap = {id(s): (dinfo['rows'][i] if i is not None else None)
            for s, i in zip(ins, link)} if dinfo else {}

    names  = dinfo['names'] if dinfo else []
    snames = dinfo['snames'] if dinfo else ['固定', '固定', '固定', '数式', '固定']
    m, ns = len(names), len(snames)
    n_split = max([len(s['naive']) for s in ins if s['kind'] == 'VALUES'], default=1)
    is_sel = any(s['kind'] == 'SELECT' for s in ins)
    split_hd = (['挿入先テーブル', '挿入列1', '挿入列2の値', '参照テーブル', '条件列', '条件値']
                if is_sel else ['区切%d' % (i + 1) for i in range(n_split)])
    xtra_hd = (['部店コード', '部店名（定義書より）', 'UNIT_ID（定義書の5_部店マスタより）']
               if is_sel else [])

    HDR, TOP = 2, 3
    cA  = 1
    cS0 = 3
    cX0 = cS0 + len(split_hd) + 1
    cD0 = (cX0 + len(xtra_hd) + 1) if xtra_hd else cX0      # ② 定義書の値
    cJ0 = cD0 + m + 1                     # ①②の突合（1列）
    cQ0 = cJ0 + 2                         # 定義書のSQL
    cK0 = cQ0 + ns + 1                    # ①③の突合（1列）
    cN0 = cK0 + 2
    cEnd = cN0 + 1

    head2(ws, cA, ['原文SQL'], C_RAW, C_RAW, '原文')
    head2(ws, cS0, split_hd, C_SPLIT, C_SPLIT, '構文の切り分け' if is_sel else 'カンマ区切り')
    if xtra_hd:
        head2(ws, cX0, xtra_hd, C_SPLIT, C_SPLIT, '参考')
    head2(ws, cD0, names, C_BLUE, C_GREEN, '横浜銀行様用出荷設定')
    head2(ws, cJ0, ['一致'], C_JDG, C_JDG, '①②の突合')
    head2(ws, cQ0, snames, C_PINK, C_PINK, 'SQL')
    head2(ws, cK0, ['一致'], C_JDG, C_JDG, '①③の突合')
    head2(ws, cN0, ['判定', '備考'], C_NOTE, C_NOTE, '確認結果')

    diffs = []
    r = TOP
    for s in ins:          # INSERT 文のみ。BEGIN / COMMIT / DELETE / UPDATE は載せない
        d = dmap.get(id(s))
        red = d['red'] if d else True
        put2(ws, r, cA, s['raw'])

        if s['kind'] == 'VALUES':
            for i, tok in enumerate(s['naive']):
                put2(ws, r, cS0 + i, tok)
        else:
            for i, v in enumerate([s['tbl'], s['col1'], s['col2'],
                                   s['src'], s['wcol'], s['wval']]):
                put2(ws, r, cS0 + i, v)
        for i, v in enumerate(s.get('_extra', [])):
            put2(ws, r, cX0 + i, v)

        iv = s['_vals']
        if d:
            bad = [i for i in range(m)
                   if nv(iv[i] if i < len(iv) else '') != nv(d['vals'][i] if i < len(d['vals']) else '')]
            for i in range(m):
                put2(ws, r, cD0 + i, d['vals'][i] if i < len(d['vals']) else '',
                     fill=C_NG if i in bad else C_YELLOW, red=red)
            for i in range(ns):
                put2(ws, r, cQ0 + i, d['sql'][i] if i < len(d['sql']) else '')
            bodydiff = bool(d['body'] and s['_body'] and nb(d['body']) != nb(s['_body']))
            if bad:
                diffs.append(dict(sheet=sec['sheet'], iline=s['line'], drow=d['row'],
                                  jdg='相違', cols='／'.join(names[i] for i in bad),
                                  iv='／'.join(nv(iv[i]) for i in bad),
                                  dv='／'.join(nv(d['vals'][i]) for i in bad)))
            elif bodydiff:
                diffs.append(dict(sheet=sec['sheet'], iline=s['line'], drow=d['row'],
                                  jdg='SQL本文のみ相違', cols='（SQL本文）',
                                  iv=s['_body'], dv=d['body']))
        else:
            for i in range(m):
                put2(ws, r, cD0 + i, '', fill=C_NONE)
            for i in range(ns):
                put2(ws, r, cQ0 + i, '', fill=C_NONE)
            diffs.append(dict(sheet=sec['sheet'], iline=s['line'], drow='',
                              jdg='定義書に該当なし', cols='',
                              iv='／'.join(nv(x) for x in iv), dv=''))

        # ---- 突合（式1本ずつ）----
        if s['kind'] == 'VALUES':
            ni = naive_index(s['args'])
            types = [to_value(t)[1] for t in s['args']][:m]
            val_f = value_cmp_formula(types, ni, cS0, cD0, r)
            cat = '&'.join('%s%d' % (CL(cQ0 + k), r) for k in range(ns))
            sql_f = ('=EXACT(SUBSTITUTE(%s%d," ",""),SUBSTITUTE(%s," ",""))'
                     % (CL(cA), r, cat))
        else:
            # ①は UNIT_CODE、②は UNIT_ID なので、定義書の 5_部店マスタ で引いた値と比べる
            val_f = '=EXACT(TRIM(%s%d)&","&TRIM(%s%d),%s%d&","&%s%d)' % (
                CL(cX0 + 2), r, CL(cS0 + 2), r, CL(cD0), r, CL(cD0 + 1), r)
            # ①は INSERT-SELECT、②は VALUES なので本文どうしで比べる
            sql_f = ('=EXACT(SUBSTITUTE(%s%d," ",""),SUBSTITUTE(%s%d&" , "&%s%d," ",""))'
                     % (CL(cQ0 + 3), r, CL(cX0 + 2), r, CL(cS0 + 2), r))
        put2(ws, r, cJ0, val_f, align='center')
        put2(ws, r, cK0, sql_f, align='center')
        ws.cell(row=r, column=2).border = BOX
        r += 1

    last = r - 1
    ws.freeze_panes = '%s%d' % (CL(cS0), TOP)
    ws.auto_filter.ref = '%s%d:%s%d' % (CL(cA), HDR, CL(cEnd), last)
    ws.column_dimensions['A'].width = 72
    for c in (2, cX0 - 1, cD0 - 1, cJ0 - 1, cQ0 - 1, cK0 - 1, cN0 - 1):
        if c >= 1:
            ws.column_dimensions[CL(c)].width = SEPW
    for i in range(len(split_hd)):
        ws.column_dimensions[CL(cS0 + i)].width = 22
    for i in range(len(xtra_hd)):
        ws.column_dimensions[CL(cX0 + i)].width = 22
    for i in range(m):
        ws.column_dimensions[CL(cD0 + i)].width = 15
    ws.column_dimensions[CL(cJ0)].width = 11
    ws.column_dimensions[CL(cK0)].width = 11
    for i in range(ns):
        ws.column_dimensions[CL(cQ0 + i)].width = 62 if snames[i] == '数式' else 11
    ws.column_dimensions[CL(cN0)].width = 8
    ws.column_dimensions[CL(cEnd)].width = 30
    ws.row_dimensions[1].height = 20
    ws.row_dimensions[HDR].height = 30
    ws.sheet_view.zoomScale = 85

    sec['diffs'] = diffs
    sec['nmatch'] = sum(1 for s in ins if dmap.get(id(s)))
    sec['ndiff'] = sum(1 for d in diffs if d['jdg'] in ('相違', 'SQL本文のみ相違'))
    sec['nnone'] = sum(1 for d in diffs if d['jdg'] == '定義書に該当なし')
    sec['nins'] = len(ins)
    sec['ndef'] = len(dinfo['rows']) if dinfo else 0
    sec['ndef_y'] = sum(1 for d in dinfo['rows'] if not d['pkg']) if dinfo else 0
    sec['names'] = names
    return ws


# ================= 値の準備 =================
def prep(secs, unit_id_of):
    """各 INSERT 文に、突合に使う値（_vals）と表示用（_disp）を持たせる。"""
    for sec in secs:
        for s in sec['stmts']:
            if s['kind'] == 'VALUES':
                vals = [to_value(t)[0] for t in s['args']]
                s['_vals'] = [nv(v) for v in vals]
                s['_disp'] = vals
                s['_body'] = s['body']
                s['_extra'] = []
            elif s['kind'] == 'SELECT':
                code = s['wval']
                uid = unit_id_of.get(code, '')
                fid = to_value(s['col2'])[0]
                s['_vals'] = [nv(uid), nv(fid)]
                s['_disp'] = [uid, fid]
                s['_body'] = ''
                s['_extra'] = []


# ================= 目次・サマリー =================
def build_index(wb, secs, srcsql, srcdef):
    ws = wb.create_sheet('00_目次', 0)
    c = put(ws, 1, 1, 'insert.sql ⇔ マスタ定義書（横浜銀行様用出荷設定）　対比表', bold=True)
    c.font = Font(name=FONT, size=14, bold=True)
    c.border = Border()
    info = [
        ('insert.sql', srcsql),
        ('マスタ定義書', srcdef),
        ('作成日', datetime.date.today().strftime('%Y/%m/%d')),
        ('シート構成', 'insert.sql のコメント「-- N_xxx」単位で1シート。並びはSQLの記載順。'),
        ('各シートの並び', '原文SQL ／ カンマ区切り ／ ① insert.sql の値 ／ '
                           '② 定義書 横浜銀行様用出荷設定 ／ ①②の突合 ／ 定義書のSQL ／ '
                           '再構築・検証 ／ 確認結果'),
        ('突合のしかた', '定義書のSQL本文と insert.sql の本文が一致する行を対応づけ、'
                         '一致しない場合は先頭の列から順にキーが合う行を対応づけている。'
                         '相違したセルは赤、定義書に該当行が無いものは橙。'),
        ('行の対応', '定義書は既存行も含むため行数が多い。①②は insert.sql の件数に合わせて並べている。'),
    ]
    r = 3
    for k, v in info:
        put(ws, r, 1, k, fill=C_RAW, bold=True)
        for cc in range(2, 11):
            ws.cell(row=r, column=cc).border = BOX
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=10)
        put(ws, r, 2, v, wrap=True)
        ws.row_dimensions[r].height = 30
        r += 1

    r += 1
    hdr = ['No', 'シート名', 'テーブル名', 'insert件数', '定義書行数',
           'うち横浜追加', '対応づいた', '一致', '相違', '定義書に無し', '項目数']
    for i, h in enumerate(hdr, 1):
        put(ws, r, i, h, fill=C_SPLIT, bold=True, align='center', wrap=True)
    hr = r
    r += 1
    tot = [0] * 7
    for i, s in enumerate(secs, 1):
        ok = s['nmatch'] - s['ndiff']
        put(ws, r, 1, i, align='center')
        cc = put(ws, r, 2, s['sheet'])
        cc.hyperlink = "#'%s'!A1" % s['sheet']
        cc.font = Font(name=FONT, size=SZ, color='FF0563C1', underline='single')
        put(ws, r, 3, ','.join(sorted({x['tbl'] for x in s['stmts'] if x['tbl']})))
        vals = [s['nins'], s['ndef'], s['ndef_y'], s['nmatch'], ok, s['ndiff'], s['nnone']]
        for j, v in enumerate(vals):
            cel = put(ws, r, 4 + j, v, align='center')
            if j == 5 and v:
                cel.fill = PatternFill('solid', fgColor=C_NG)
            if j == 6 and v:
                cel.fill = PatternFill('solid', fgColor=C_NONE)
            tot[j] += v
        put(ws, r, 11, len(s['names']), align='center')
        r += 1
    put(ws, r, 1, '', fill=C_RAW)
    put(ws, r, 2, '合計', fill=C_RAW, bold=True)
    put(ws, r, 3, '', fill=C_RAW)
    for j, v in enumerate(tot):
        put(ws, r, 4 + j, v, fill=C_RAW, bold=True, align='center')
    put(ws, r, 11, '', fill=C_RAW)
    ws.auto_filter.ref = 'A%d:K%d' % (hr, r - 1)

    # ---- 所見（突合結果の読み取り）----
    r += 2
    block_label(ws, r, 1, 11, '所見（「00_突合結果」を読んだ結果）', C_NOTE)
    r += 1
    for i, h in enumerate(['No', '対象シート', '件数', '内容'], 1):
        put(ws, r, i, h, fill=C_SPLIT, bold=True, align='center')
    ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=7)
    put(ws, r, 4, '内容', fill=C_SPLIT, bold=True, align='center')
    ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=11)
    put(ws, r, 8, '確認したいこと', fill=C_SPLIT, bold=True, align='center')
    for cc in (5, 6, 7, 9, 10, 11):
        ws.cell(row=r, column=cc).fill = PatternFill('solid', fgColor=C_SPLIT)
        ws.cell(row=r, column=cc).border = BOX
    r += 1
    for i, (sh, cnt, what, ask) in enumerate(NOTES, 1):
        for cc in range(1, 12):
            ws.cell(row=r, column=cc).border = BOX
        put(ws, r, 1, i, align='center')
        put(ws, r, 2, sh, wrap=True)
        put(ws, r, 3, cnt, align='center')
        ws.merge_cells(start_row=r, start_column=4, end_row=r, end_column=7)
        put(ws, r, 4, what, wrap=True)
        ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=11)
        put(ws, r, 8, ask, wrap=True)
        ws.row_dimensions[r].height = 58
        r += 1

    for col, w in zip('ABCDEFGHIJK', [5, 30, 26, 10, 10, 11, 11, 8, 8, 12, 8]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A%d' % (hr + 1)
    ws.sheet_view.showGridLines = False
    return ws


NOTES = [
    ('30_取引リスト表示管理', 32,
     'insert.sql には PRF_ID=11（L&F担当者）の DEAL_ID=10（未判定照会）・11（判定済照会）'
     'が入っているが、定義書には無い。他のプロファイル（2・3・4・5・9・10）には両方ある。'
     'SCL#5（バッチフィルタリング）ケース07・11・16・20・26 で、L&F が未判定照会・判定済照会を'
     '使うことは確認済み。',
     '資材（insert.sql）が正しく、定義書の記載漏れという理解でよろしいでしょうか。'
     '定義書側を是正する認識で進めます。'),
    ('17_コード設定マスタ', 2,
     'DSP_ORDER が insert=103／106、定義書=102／105 と1つずれている。'
     'insert.sql 側には別に UPDATE 文（CODE_CNF_ID=34 の DSP_ORDER を 103→105）もある。'
     '対象は「該当あり（C情報）」と「取消し」。',
     'プルダウン（補助コメント）の並び順に直結します。ヒアリングシートの指定はどちらでしょうか。'
     'SCL#5 ケース58 の並び順とも相違しているため、合わせて確認させてください。'),
    ('11_プロファイル設定', 1,
     'insert.sql に PRF_ID=11（L&F担当者）／EXEC_AUTH_ID=1003 の行があるが、定義書に無い。',
     'EXEC_AUTH_ID=1003 は何の権限でしょうか。定義書に追記が必要か確認させてください。'),
    ('34_リスト処理属性管理', 1,
     'RSV_FLG1 が insert=0／定義書=1 で相違（LST_ID=16 BファイルC情報）。',
     '予備フラグのため影響は小さいと思われますが、どちらが正でしょうか。'),
    ('15_リスト発行元マスタ', 1,
     '値は一致。ただし定義書のSQL生成数式が NULL を \'NULL\'（文字列）として出力している'
     '（IS_DEMANDED・FILE_MAKE_STEP）。',
     '定義書のSQL列をそのまま流すと文字列 \'NULL\' が入ります。'
     '数式側の誤りという理解でよろしいでしょうか。'),
    ('25_部店フィルタリング設定', 0,
     '25部店すべてが同じ12個の設定ID（1,4,5,7,8,9,11,12,13,14,15,16）。'
     '2・3・6・10 は未設定。定義書と insert.sql は一致している。',
     '抜けている4つは意図的な除外でしょうか。ヒアリングシートの設定内容と突合したいです。'),
    ('30_取引リスト表示管理（参考）', 0,
     'PRF 9・10（L&F管理者）には DEAL_ID 14・15（連動）の行があるが、'
     'PRF 11（L&F担当者）には無い。定義書と insert.sql は一致しているため相違ではない。',
     '担当者だけ連動の行が無いのは意図どおりでしょうか。'),
]


def build_summary(wb, secs):
    ws = wb.create_sheet('00_突合結果', 1)
    c = put(ws, 1, 1, '① insert.sql と ② 定義書 が合わなかった行', bold=True)
    c.font = Font(name=FONT, size=14, bold=True)
    c.border = Border()
    hdr = ['No', 'シート', 'SQLの行', '定義書の行', '判定', '相違した列',
           '① insert.sql の値', '② 定義書の値']
    r = 3
    for i, h in enumerate(hdr, 1):
        put(ws, r, i, h, fill=C_SPLIT, bold=True, align='center')
    hr = r
    r += 1
    n = 0
    for s in secs:
        for d in s['diffs']:
            n += 1
            put(ws, r, 1, n, align='center')
            cc = put(ws, r, 2, d['sheet'])
            cc.hyperlink = "#'%s'!A1" % d['sheet']
            cc.font = Font(name=FONT, size=SZ, color='FF0563C1', underline='single')
            put(ws, r, 3, d['iline'], align='center')
            put(ws, r, 4, d['drow'], align='center')
            put(ws, r, 5, d['jdg'], align='center',
                fill=C_NG if d['jdg'] == '相違' else C_NONE)
            put(ws, r, 6, d['cols'])
            put(ws, r, 7, d['iv'], wrap=True)
            put(ws, r, 8, d['dv'], wrap=True)
            r += 1
    if n == 0:
        put(ws, r, 1, '相違なし')
    ws.auto_filter.ref = 'A%d:H%d' % (hr, max(hr, r - 1))
    for col, w in zip('ABCDEFGH', [5, 28, 9, 10, 15, 28, 46, 46]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = 'A%d' % (hr + 1)
    ws.sheet_view.showGridLines = False
    print('  突合で合わなかった行 : %d 件' % n)
    return ws


# ================= main =================
def main():
    srcsql, srcdef, out = sys.argv[1], sys.argv[2], sys.argv[3]

    head, secs = load(srcsql)
    wbd = openpyxl.load_workbook(srcdef, data_only=True)

    # 部店コード -> UNIT_ID / 部店名（定義書の 5_部店マスタ から）
    unit_id_of, unit_name_of = {}, {}
    if '5_部店マスタ' in wbd.sheetnames:
        di = read_def(wbd['5_部店マスタ'])
        for d in di['rows']:
            if len(d['vals']) >= 3:
                unit_id_of[d['vals'][1]] = d['vals'][0]
                unit_name_of[d['vals'][1]] = d['vals'][2]
    prep(secs, unit_id_of)
    for sec in secs:
        for s in sec['stmts']:
            if s['kind'] == 'SELECT':
                s['_extra'] = [s['wval'],
                               unit_name_of.get(s['wval'], '（定義書に無し）'),
                               unit_id_of.get(s['wval'], '')]

    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    used = set()
    for s in secs:
        s['sheet'] = sheet_name(s['name'], used)
    for s in secs:
        dinfo = read_def(wbd[s['name']]) if s['name'] in wbd.sheetnames else None
        extra = None
        build_sheet(wb, s, dinfo, extra)
        print('%-30s insert %4d / 定義書 %5d / 対応 %4d / 相違 %3d / 無し %4d'
              % (s['sheet'], s['nins'], s['ndef'], s['nmatch'], s['ndiff'], s['nnone']))
    wb.active = 0
    wb.save(out)
    show_diffs(secs)
    print()
    print('出力 :', out)
    print('シート数 :', len(wb.sheetnames))



def show_diffs(secs):
    """合わなかった行を画面に出す（シートには載せない）。"""
    n = 0
    for s in secs:
        for d in s['diffs']:
            n += 1
            print('%3d  %-26s SQL%-6s 定義書%-6s %-14s %s' %
                  (n, d['sheet'], d['iline'], d['drow'], d['jdg'], d['cols']))
            if d['jdg'] != '定義書に該当なし':
                print('        (1) %s' % d['iv'][:110])
                print('        (2) %s' % d['dv'][:110])
            else:
                print('        (1) %s' % d['iv'][:110])
    print()
    print('合わなかった行 : %d 件' % n)


if __name__ == '__main__':
    main()
