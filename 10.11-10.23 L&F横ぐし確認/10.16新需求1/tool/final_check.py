# -*- coding: utf-8 -*-
"""
最終確認。
  (1) insert.sql の INSERT 件数  と  Excel の行数 が一致するか
  (2) Excel に載せた原文SQL が insert.sql の該当行と1文字も違わないか
  (3) Excel の②の値・②のSQL が マスタ定義書の該当セルと一致するか（取りこぼし確認）
"""
import warnings
warnings.filterwarnings('ignore')
import sys, collections
import openpyxl
sys.path.insert(0, '.')
from build_compare import read_def, nv, nv2, LABEL
from sql2excel import load

SQL = r'C:\Users\Sun\Desktop\major\10.11-10.23 L&F横ぐし確認\10.16新需求1\insert.sql'
DEF = r'C:\Users\Sun\Desktop\major\10.11-10.23 L&F横ぐし確認\10.16新需求1\Oculus-filterV5.0_マスタ定義書_AML01_横浜 1.xlsx'
OUT = sys.argv[1] if len(sys.argv) > 1 else 'test4.xlsx'

head, secs = load(SQL)
wb = openpyxl.load_workbook(OUT)
wbd = openpyxl.load_workbook(DEF, data_only=True)

print('%-28s %7s %7s %7s  %s' % ('シート', 'SQL', 'Excel', '定義書', '判定'))
print('-' * 72)
ng = []
tot_sql = tot_xls = 0
for sec in secs:
    name = sec['name']
    ins = [x for x in sec['stmts'] if x['kind'] in ('VALUES', 'SELECT')]
    other = [x for x in sec['stmts'] if x['kind'] not in ('VALUES', 'SELECT')]
    ws = wb[name]
    rows = [r for r in range(3, ws.max_row + 1) if ws.cell(row=r, column=1).value]
    di = read_def(wbd[name])
    ok = (len(ins) == len(rows))
    tot_sql += len(ins); tot_xls += len(rows)

    # 原文SQL が1文字も違わないか
    mism = 0
    for st, r in zip(ins, rows):
        if ws.cell(row=r, column=1).value != st['raw']:
            mism += 1
    # ②の値・②のSQL が定義書のセルと一致するか
    hdr = {}
    for c in range(1, ws.max_column + 1):
        hdr.setdefault(str(ws.cell(row=2, column=c).value or ''), []).append(c)
    dcols = []
    for c in range(1, ws.max_column + 1):
        if str(ws.cell(row=1, column=c).value or '') == '横浜銀行様用出荷設定':
            dcols = list(range(c, c + len(di['names'])))
    dmis = 0
    # Excel に出ている②の値が、定義書のどこかの行にそのまま存在するか
    defset = collections.Counter(tuple(d['vals']) for d in di['rows'])
    for r in rows:
        v = tuple(nv(ws.cell(row=r, column=c).value) for c in dcols)
        if all(x == '' for x in v):
            continue                       # 定義書に該当なしの行
        if defset[v] <= 0:
            dmis += 1

    mark = 'OK' if (ok and mism == 0 and dmis == 0) else 'NG'
    if mark == 'NG':
        ng.append((name, len(ins), len(rows), mism, dmis))
    extra = ('  ※%s' % '／'.join('%s %d件' % (k, sum(1 for x in other if x['kind'] == k))
                                 for k in sorted({x['kind'] for x in other}))) if other else ''
    print('%-28s %7d %7d %7d  %s%s' % (name, len(ins), len(rows), len(di['rows']), mark, extra))

print('-' * 72)
print('%-28s %7d %7d' % ('合計', tot_sql, tot_xls))
print()
if ng:
    print('■ 不一致')
    for x in ng:
        print('   ', x)
else:
    print('■ すべて一致')
    print('   ・insert.sql の INSERT 件数 と Excel の行数 が一致')
    print('   ・Excel の原文SQL が insert.sql と1文字も相違なし')
    print('   ・Excel に載せた②の値が、いずれも定義書に実在する行と一致')
print()
allother = [(x['line'], x['raw']) for sec in secs for x in sec['stmts']
            if x['kind'] not in ('VALUES', 'SELECT')]
print('■ Excel に載せていない文（INSERT以外）: %d 件' % len(allother))
for ln, t in allother:
    print('   %5d行  %s' % (ln, t))
