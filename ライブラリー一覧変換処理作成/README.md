# libInfo.txt 变换处理（txt → txt，不用 Excel）

1. 把 `libInfo.txt` 拖到 **`run.bat`** 上（或放进同一文件夹，双击 `run.bat`）
2. 先出调查结果。带 ★ 的项目要确认，尤其「含 folder 的行」不是 0 时
3. 确认没问题按 `Y`，生成 **`libInfo_converted.txt`**（Tab 分隔 / Shift_JIS）

- 只被 `old` 命中的行默认**不删**，另存到 `_old_candidates.txt`；确认该删后加 `-NoSafeOld` 重跑
- 改输出编码、删除关键词 → 改 `convert.ps1` 开头的参数和 `$Patterns`
- 试跑用 `test\libInfo.txt`；`LibInfoConv.bas` 是 Excel 宏版，用不到可以删
