# libInfo.txt 变换处理（txt → txt，不用 Excel）

## 怎么用

1. 把 `libInfo.txt` 拖到 **`run.bat`** 上（或放进同一文件夹，双击 `run.bat`）
2. 先出调查结果。带 ★ 的项目要确认，尤其「含 folder 的行」不是 0 时
3. 确认没问题按 `Y`，生成 **`libInfo_converted.txt`**（Tab 分隔 / Shift_JIS）

- 只被 `old` 命中的行**默认删除**，同时留档到 `_old_candidates.txt`；想先不删就加 `-KeepOld`
- 改输出编码、删除关键词 → 改 `convert.ps1` 开头的参数和 `$Patterns`
- 行末空格：默认 `-TailSpace manual`（只删1个，**与手工正解逐字节一致**）；`all` = 全删，数据更干净但和正解每行都有差
- 命中删除模式但不该删的行 → 写进 `convert.ps1` 的 `$KeepPatterns`
  已登录：`AfterRcv_org.bat`（`_org.` 命中但手工正解保留的正规批处理）
- 和手工做的 `_変換後.txt` 对账 → 双击 **`compare.bat`**（对应「步骤」⑦ 結果検証）
  报告做成一屏能截完的长度（`compare_report.txt`）；要全部明细加 `-Full`
- 试跑用 `test\libInfo.txt`；`LibInfoConv.bas` 是 Excel 宏版，用不到可以删

## 脚本之间的关系

```
run.bat
  │
  ├─① convert.ps1 -Mode check     只调查，不写任何文件 → 显示报告
  │
  │   ……按 Y 才继续（按别的键直接退出，什么都没动）
  │
  └─② convert.ps1 -Mode convert   真正变换，写出 txt

compare.bat ─> compare.ps1         run.bat 不调用它，需要时自己双击
```

- **`run.bat`** 只是个壳：找 PowerShell、传文件路径、中间插一个 Y/N 确认。没有任何处理逻辑
- **`convert.ps1`** 是本体。`-Mode check` 和 `-Mode convert` 跑的是同一套判定，区别只在「最后写不写文件」
- 所以 ② 会把 ① 的报告**再打印一遍** —— 不是 bug，是让最终执行自带证据
- **`compare.ps1`** 和变换无关，只做「脚本输出 vs 手工正解」的差分
- `LibInfoConv.bas` 是 Excel 宏版，和上面三个没有关系，是另一条路

## convert.ps1 内部的流程

```
读入（Shift_JIS）
  → 丢掉第 1 行
  → 逐行判定：命中 11 个删除模式中任意一个 → 丢弃（按模式分别计数）
  → 活下来的行：3 个空格 → Tab      …手順書(1)
                删掉行末空格         …手順書(8)
  → 写出（Tab 分隔）
```

11 个删除模式：
`\log\` `\save\` `\common\tools\JudgedCIF\data\output` `コピー`
`bk.` `_bk` `bkup` `bak` `org.` `old` `origin.txt`

全部是**不区分大小写的字面匹配，不用正则** ——
正则会把 `bk.` `org.` 里的 `.` 变成通配符，把手順書写明「対象外」的 `ORGXX` 一起删掉。
