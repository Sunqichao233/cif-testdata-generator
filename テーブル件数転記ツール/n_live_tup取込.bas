Attribute VB_Name = "modTableCount"
Option Explicit

'================================================================
' n_live_tup 取込マクロ
'
'   統計情報ファイルから切り出した CSV を読み込み、
'   @テーブル件数確認_filter の対象シートの右端に 1 か月分（3 列）を追加する。
'
'   前任者のマクロ（n_live_tup取込）の構成をそのまま踏襲し、
'   出力だけを提示フォーマット（取得日／件数／想定件数比率）に合わせている。
'
'   ・出力は表示中の 1 シートだけ。別シートは作らない
'   ・CSV は読み取るだけで、保存せずに閉じる
'================================================================


'========================
'設定
'========================

'--- 転記先シート（@テーブル件数確認_filter）---
Private Const YM_ROW          As Long = 2     ' 年月を書く行
Private Const HEADER_ROW      As Long = 3     ' 見出し行
Private Const FIRST_DATA_ROW  As Long = 4     ' データ開始行
Private Const COL_SCHEMA      As Long = 1     ' A列 スキーマ名
Private Const COL_TABLE       As Long = 2     ' B列 テーブル名
Private Const BLOCK_COLS      As Long = 3     ' 1 か月あたりの列数
Private Const YM_OFFSET_MONTH As Long = -1    ' 年月＝取得日の何か月ずらしか（-1＝前月）
Private Const YM_FORMAT       As String = "yyyy/m"

Private Const HDR_DATE  As String = "取得日"
Private Const HDR_COUNT As String = "件数"
Private Const HDR_RATIO As String = "想定件数比率"

'--- CSV 側 ---
' 見出し行（schemaname / relname / n_live_tup）を自動で探す。
' 見つからない場合だけ、下の既定値を使う（前任者のマクロと同じ位置）。
Private Const CSV_COL_SCHEMA As Long = 2      ' B列 schemaname
Private Const CSV_COL_TABLE  As Long = 3      ' C列 relname
Private Const CSV_COL_LIVE   As Long = 15     ' O列 n_live_tup
Private Const CSV_FIRST_ROW  As Long = 42     ' データ開始行
Private Const CSV_SEARCH_ROWS As Long = 300   ' 見出し行を探す範囲

'--- その他 ---
Private Const NG_COLOR       As Long = 255    ' CSV に無いテーブルの塗り色（RGB(255,0,0)）
Private Const MSG_LIST_MAX   As Long = 10     ' 完了メッセージに並べる最大件数
Private Const MSG_LIST_CHARS As Long = 400    ' 1 つの一覧に使う最大文字数


'================================================================
' 本体
'================================================================
Sub n_live_tup取込()

    '========================
    '変数定義
    '========================

    'CSVファイル用
    Dim wbCSV As Workbook                   '開いたCSVブック
    Dim wsCSV As Worksheet                  'CSVシート

    '転記先Excel用
    Dim wsXLS As Worksheet                  '表示中のシート（横東 / 陸道七）

    'schema|table名 → n_live_tup を保持
    Dim dict As Object

    'CSV選択ダイアログの戻り値
    Dim csvPath As Variant

    '最終行
    Dim lastRowCSV As Long
    Dim lastRowXLS As Long

    Dim i As Long

    'CSV側の情報
    Dim schemaName As String                'B列(schemaname)
    Dim tblName As String                   'C列(relname)

    'Excel側の情報
    Dim currentSchema As String             'A列のschema(aml/aml01/aml02...)
    Dim excelTbl As String                  'B列から抽出したテーブル物理名

    'Dictionary検索用キー
    Dim dictKey As String

    'CSVの列位置と開始行（見出し行から決める）
    Dim colSchema As Long, colTable As Long, colLive As Long, firstRow As Long

    '追加する3列
    Dim lastCol As Long, newCol As Long
    Dim getDate As Date
    Dim fileName As String

    '集計用
    Dim usedKey As Object                   'CSVのうち転記に使ったキー
    Dim sheetSchema As Object               'シートに出てくるスキーマ
    Dim missList As Object                  'シートにあるがCSVに無い
    Dim extraList As Object                 'CSVにあるがシートに無い
    Dim nWrite As Long


    '========================
    '初期設定
    '========================

    Set wsXLS = ActiveSheet                 '表示中のシートに出力する（別シートは作らない）
    Set dict = CreateObject("Scripting.Dictionary")

    '転記先のシートかどうか確認する
    lastCol = wsXLS.Cells(HEADER_ROW, wsXLS.Columns.Count).End(xlToLeft).Column
    If Not IsTargetSheet(wsXLS, lastCol) Then
        MsgBox "このシートは転記の対象ではないようです。" & vbCrLf & vbCrLf & _
               HEADER_ROW & " 行目が見出し行で、右端が「" & _
               HDR_DATE & "／" & HDR_COUNT & "／" & HDR_RATIO & "」の順に並んでいる" & _
               "シート（横東 / 陸道七）を表示してから実行してください。", _
               vbExclamation, "n_live_tup取込"
        Exit Sub
    End If

    lastRowXLS = wsXLS.Cells(wsXLS.Rows.Count, COL_TABLE).End(xlUp).Row   'B列最終行
    If lastRowXLS < FIRST_DATA_ROW Then
        MsgBox "データ行が見つかりません。", vbExclamation, "n_live_tup取込"
        Exit Sub
    End If


    '========================
    'CSV選択
    '========================

    csvPath = Application.GetOpenFilename("CSV (*.csv),*.csv")   'ファイル選択ダイアログ表示

    If VarType(csvPath) = vbBoolean Then Exit Sub                'キャンセルしたら終了

    fileName = Mid$(CStr(csvPath), InStrRev(CStr(csvPath), "\") + 1)

    'ファイル名末尾12桁(yyyymmddhhmm)を取得日にする
    If Not StampToDate(fileName, getDate) Then
        MsgBox "ファイル名から取得日を読み取れませんでした。" & vbCrLf & vbCrLf & _
               "選択したファイル : " & fileName & vbCrLf & _
               "期待する形式     : ～_yyyymmddhhmm.csv" & vbCrLf & _
               "例               : exp_Statics_202607010100.csv", _
               vbExclamation, "n_live_tup取込"
        Exit Sub
    End If


    '========================
    'CSVを開く
    '========================

    Application.ScreenUpdating = False

    Set wbCSV = Workbooks.Open(csvPath, ReadOnly:=True)   'CSVを開く（読み取り専用）
    Set wsCSV = wbCSV.Sheets(1)                           '1シート目取得


    '========================
    'CSV読込
    '
    'CSV
    ' B列 = schema
    ' C列 = table名(relname)
    ' O列 = n_live_tup
    '
    ' Dictionaryキー :
    ' schema|table名
    '
    ' 例
    ' aml|cnf_sec_tbl
    ' aml01|cnf_sec_tbl
    ' aml02|cnf_sec_tbl
    '
    ' 列位置は見出し行（schemaname / relname / n_live_tup）から決める。
    ' 統計情報ファイル全体を CSV にすると前に別のセクションが入り、
    ' データ開始行がファイルによって変わるため。
    ' 見出しが見つからない場合だけ B / C / O 列・42 行目を使う。
    '========================

    If Not FindCsvHeader(wsCSV, colSchema, colTable, colLive, firstRow) Then
        colSchema = CSV_COL_SCHEMA
        colTable = CSV_COL_TABLE
        colLive = CSV_COL_LIVE
        firstRow = CSV_FIRST_ROW
    End If

    lastRowCSV = wsCSV.Cells(wsCSV.Rows.Count, colTable).End(xlUp).Row   'テーブル名列の最終行取得

    Dim started As Boolean
    Dim liveVal As Variant

    For i = firstRow To lastRowCSV          '見出しの次の行から最終行まで。

        schemaName = LCase$(Trim$(CStr(wsCSV.Cells(i, colSchema).Value)))   'schema取得
        tblName = LCase$(Trim$(CStr(wsCSV.Cells(i, colTable).Value)))       'テーブル名取得
        liveVal = wsCSV.Cells(i, colLive).Value

        'データが始まったあとで空行が来たら、そこがセクションの終わり
        If started And tblName = "" Then Exit For

        If schemaName <> "" And tblName <> "" Then                         '空チェック
            If IsNumeric(liveVal) Then                                     '罫線行(-----)を除くため

                started = True
                dictKey = schemaName & "|" & tblName                       'キー作成(ex:aml01|cnf_sec_tbl)

                '値は n_live_tup(O列)
                dict(dictKey) = CDbl(liveVal)                              'n_live_tup格納

            End If
        End If

    Next i

    wbCSV.Close False                                                      '保存せず閉じる

    If dict.Count = 0 Then
        Application.ScreenUpdating = True
        MsgBox "CSV から件数を読み取れませんでした。" & vbCrLf & vbCrLf & _
               "・「テーブル統計情報」の部分が入っているか" & vbCrLf & _
               "・見出しに schemaname / relname / n_live_tup があるか" & vbCrLf & _
               "を確認してください。", vbExclamation, "n_live_tup取込"
        Exit Sub
    End If


    '========================
    '3列追加
    '
    ' 見出し行の右端の次から
    '   1列目 = 取得日
    '   2列目 = 件数(n_live_tup)
    '   3列目 = 想定件数比率(空欄。書式のみ)
    ' 2行目に年月(取得日の前月)
    '========================

    '右端に同じ取得日が入っていないか確認する
    Dim prevDate As Variant
    prevDate = FirstDateInColumn(wsXLS, FIRST_DATA_ROW, lastRowXLS, lastCol - BLOCK_COLS + 1)
    If IsDate(prevDate) Then
        If CDate(prevDate) = getDate Then
            Application.ScreenUpdating = True
            If MsgBox("右端に同じ取得日（" & Format$(getDate, "yyyy/m/d h:mm") & "）の" & _
                      "データが既にあります。" & vbCrLf & vbCrLf & _
                      "それでも 3 列追加しますか？", _
                      vbYesNo + vbExclamation, "n_live_tup取込") <> vbYes Then Exit Sub
            Application.ScreenUpdating = False
        End If
    End If

    On Error GoTo Fail

    newCol = lastCol + 1

    '直前の3列から書式・列幅をコピーする（結合セル・罫線・塗りも付いてくる）
    wsXLS.Range(wsXLS.Cells(YM_ROW, lastCol - BLOCK_COLS + 1), _
                wsXLS.Cells(lastRowXLS, lastCol)).Copy
    wsXLS.Cells(YM_ROW, newCol).PasteSpecial Paste:=xlPasteFormats
    Application.CutCopyMode = False

    For i = 0 To BLOCK_COLS - 1
        wsXLS.Columns(newCol + i).ColumnWidth = _
            wsXLS.Columns(lastCol - BLOCK_COLS + 1 + i).ColumnWidth
    Next i

    '直前の件数列に赤セルが残っていると書式ごとコピーされるため消す
    wsXLS.Range(wsXLS.Cells(FIRST_DATA_ROW, newCol + 1), _
                wsXLS.Cells(lastRowXLS, newCol + 1)).Interior.Pattern = xlNone

    '年月と見出し
    wsXLS.Cells(YM_ROW, newCol).Value = _
        Format$(DateAdd("m", YM_OFFSET_MONTH, getDate), YM_FORMAT)
    wsXLS.Cells(HEADER_ROW, newCol).Value = HDR_DATE
    wsXLS.Cells(HEADER_ROW, newCol + 1).Value = HDR_COUNT
    wsXLS.Cells(HEADER_ROW, newCol + 2).Value = HDR_RATIO


    '========================
    'Excelへ転記
    '
    'A列 = schema
    'B列 = テーブル名
    '追加した3列 = 転記先
    '========================

    Set usedKey = CreateObject("Scripting.Dictionary")
    Set sheetSchema = CreateObject("Scripting.Dictionary")
    Set missList = CreateObject("Scripting.Dictionary")

    currentSchema = ""                      '初期化

    For i = FIRST_DATA_ROW To lastRowXLS    '4行目から最終行まで処理。

        '----------------------------------
        'A列にschemaがある行だけ更新
        '
        'aml
        ' amlのテーブル群
        '
        'aml01
        ' aml01のテーブル群
        '----------------------------------
        If Trim$(CStr(wsXLS.Cells(i, COL_SCHEMA).Value)) <> "" Then
            currentSchema = LCase$(Trim$(CStr(wsXLS.Cells(i, COL_SCHEMA).Value)))   'currentSchemaにスキーマを保存し値があるまで更新しない
        End If

        excelTbl = CStr(wsXLS.Cells(i, COL_TABLE).Value)

        '----------------------------------
        '物理テーブル名抽出
        '
        'セキュリティ設定(CNF_SEC_TBL)
        ' ↓
        'cnf_sec_tbl
        '
        ' 最後の括弧の中身を取る。
        ' リストファイルアップロード管理（バッチ）(MNG_..._TBL) のように
        ' 括弧が2組ある行があるため。
        '----------------------------------
        excelTbl = LCase$(ExtractTableName(excelTbl))

        If excelTbl <> "" And currentSchema <> "" Then

            If Not sheetSchema.Exists(currentSchema) Then sheetSchema.Add currentSchema, True

            '1列目 取得日
            wsXLS.Cells(i, newCol).Value = getDate

            '----------------------------------
            '検索キー作成
            '
            '例
            ' aml01|cnf_sec_tbl
            '----------------------------------
            dictKey = currentSchema & "|" & excelTbl

            '----------------------------------
            '一致したら2列目へ転記
            '無ければ空欄のまま赤く塗る
            '----------------------------------
            If dict.Exists(dictKey) Then

                wsXLS.Cells(i, newCol + 1).Value = dict(dictKey)
                If Not usedKey.Exists(dictKey) Then usedKey.Add dictKey, True
                nWrite = nWrite + 1

            Else

                wsXLS.Cells(i, newCol + 1).ClearContents
                wsXLS.Cells(i, newCol + 1).Interior.Color = NG_COLOR
                If Not missList.Exists(dictKey) Then missList.Add dictKey, currentSchema & "." & excelTbl

            End If

        End If

    Next i


    '========================
    'CSVにあってシートに無いテーブル
    '
    ' シートのA列に出てくるスキーマだけを対象にする。
    ' 無関係なスキーマまで並べると読めないため。
    '========================

    Set extraList = CreateObject("Scripting.Dictionary")

    Dim k As Variant, sch As String
    For Each k In dict.Keys
        If Not usedKey.Exists(k) Then
            sch = Left$(CStr(k), InStr(CStr(k), "|") - 1)
            If sheetSchema.Exists(sch) Then
                extraList.Add k, Replace(CStr(k), "|", ".")
            End If
        End If
    Next k

    Application.ScreenUpdating = True


    '========================
    '完了メッセージ
    '========================

    Dim m As String
    m = "n_live_tup転記完了" & vbCrLf & vbCrLf
    m = m & "  対象シート             : " & wsXLS.Name & vbCrLf
    m = m & "  CSV                    : " & fileName & vbCrLf
    m = m & "  取得日                 : " & Format$(getDate, "yyyy/m/d h:mm") & vbCrLf
    m = m & "  追加した列             : " & ColLetter(newCol) & "～" & _
                                           ColLetter(newCol + BLOCK_COLS - 1) & "列" & _
                                           "（年月 " & wsXLS.Cells(YM_ROW, newCol).Value & "）" & vbCrLf & vbCrLf
    m = m & "  転記                   : " & nWrite & " 行" & vbCrLf
    m = m & "  CSVに無いテーブル      : " & missList.Count & " 件（赤セル）" & vbCrLf
    m = m & "  シートに無いテーブル   : " & extraList.Count & " 件" & vbCrLf

    If missList.Count > 0 Then
        m = m & vbCrLf & "【CSVに無いテーブル】件数は空欄で赤く塗りました" & vbCrLf
        m = m & ListToText(missList)
    End If
    If extraList.Count > 0 Then
        m = m & vbCrLf & "【シートに無いテーブル】テーブル追加の可能性があります" & vbCrLf
        m = m & ListToText(extraList)
    End If

    m = m & vbCrLf & "※ 自動保存はしません。内容を確認してから保存してください。"
    MsgBox m, vbInformation, "n_live_tup取込"
    Exit Sub

Fail:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    MsgBox "処理中にエラーが発生しました。" & vbCrLf & vbCrLf & _
           Err.Number & " : " & Err.Description & vbCrLf & vbCrLf & _
           "保存せずにブックを閉じると、実行前の状態に戻せます。", _
           vbCritical, "n_live_tup取込"
End Sub


'================================================================
' CSV の見出し行を探す
'   schemaname / relname / n_live_tup がそろっている行を見出しとみなす。
'   「テーブルサイズ」のように schemaname と relname はあるが
'   n_live_tup が無いセクションは見出しにしない。
'================================================================
Private Function FindCsvHeader(ByVal ws As Worksheet, ByRef colSchema As Long, _
                               ByRef colTable As Long, ByRef colLive As Long, _
                               ByRef firstRow As Long) As Boolean
    Dim r As Long, c As Long, lastC As Long
    Dim s As String
    Dim cs As Long, ct As Long, cl As Long

    For r = 1 To CSV_SEARCH_ROWS
        cs = 0: ct = 0: cl = 0
        lastC = ws.Cells(r, ws.Columns.Count).End(xlToLeft).Column
        If lastC > 60 Then lastC = 60
        For c = 1 To lastC
            s = LCase$(Trim$(CStr(ws.Cells(r, c).Value)))
            Select Case s
                Case "schemaname", "schema": cs = c
                Case "relname", "tablename", "table_name": ct = c
                Case "n_live_tup":           cl = c
            End Select
        Next c
        If cs > 0 And ct > 0 And cl > 0 Then
            colSchema = cs
            colTable = ct
            colLive = cl
            firstRow = r + 1
            FindCsvHeader = True
            Exit Function
        End If
    Next r
End Function


'================================================================
' 転記先のシートかどうか。右端3列が 取得日／件数／想定件数比率 か見る
'================================================================
Private Function IsTargetSheet(ByVal ws As Worksheet, ByVal lastCol As Long) As Boolean
    If lastCol < COL_TABLE + BLOCK_COLS Then Exit Function
    If Trim$(CStr(ws.Cells(HEADER_ROW, lastCol - 2).Value)) <> HDR_DATE Then Exit Function
    If Trim$(CStr(ws.Cells(HEADER_ROW, lastCol - 1).Value)) <> HDR_COUNT Then Exit Function
    If Trim$(CStr(ws.Cells(HEADER_ROW, lastCol).Value)) <> HDR_RATIO Then Exit Function
    IsTargetSheet = True
End Function


'================================================================
' ファイル名末尾12桁(yyyymmddhhmm) → 日付
'================================================================
Private Function StampToDate(ByVal fileName As String, ByRef dt As Date) As Boolean
    Dim base As String, s As String, i As Long

    base = fileName
    If InStrRev(base, ".") > 0 Then base = Left$(base, InStrRev(base, ".") - 1)
    If Len(base) < 12 Then Exit Function
    s = Right$(base, 12)

    For i = 1 To 12
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i

    On Error GoTo Bad
    dt = DateSerial(CLng(Left$(s, 4)), CLng(Mid$(s, 5, 2)), CLng(Mid$(s, 7, 2))) _
       + TimeSerial(CLng(Mid$(s, 9, 2)), CLng(Mid$(s, 11, 2)), 0)
    StampToDate = True
    Exit Function
Bad:
End Function


'================================================================
' B列から「最後の括弧の中身」を取り出す
'   リストファイルアップロード管理（バッチ）(MNG_LIST_FILE_UPLOAD_BATCH_TBL)
'     → MNG_LIST_FILE_UPLOAD_BATCH_TBL
'================================================================
Private Function ExtractTableName(ByVal s As String) As String
    Dim p1 As Long, p2 As Long

    s = Trim$(s)
    If s = "" Then Exit Function

    p2 = InStrRev(s, ")")
    If p2 > 0 Then
        p1 = InStrRev(s, "(", p2)
        If p1 > 0 And p2 - p1 > 1 Then
            ExtractTableName = Trim$(Mid$(s, p1 + 1, p2 - p1 - 1))
            Exit Function
        End If
    End If

    '半角括弧が無い場合のみ全角で試す
    p2 = InStrRev(s, "）")
    If p2 > 0 Then
        p1 = InStrRev(s, "（", p2)
        If p1 > 0 And p2 - p1 > 1 Then
            ExtractTableName = Trim$(Mid$(s, p1 + 1, p2 - p1 - 1))
        End If
    End If
End Function


'================================================================
' 小物
'================================================================
Private Function FirstDateInColumn(ByVal ws As Worksheet, ByVal r1 As Long, _
                                   ByVal r2 As Long, ByVal col As Long) As Variant
    Dim r As Long
    For r = r1 To r2
        If IsDate(ws.Cells(r, col).Value) Then
            FirstDateInColumn = ws.Cells(r, col).Value
            Exit Function
        End If
    Next r
    FirstDateInColumn = Empty
End Function

Private Function ColLetter(ByVal col As Long) As String
    Dim a As String
    a = Cells(1, col).Address(True, False)      '列だけ相対にすると "A$1" になる
    ColLetter = Left$(a, InStr(a, "$") - 1)
End Function

Private Function ListToText(ByVal d As Object) As String
    'MsgBox は約 1024 文字で切れるため、件数と文字数の両方で打ち切る
    Dim k As Variant, n As Long, s As String
    For Each k In d.Keys
        If n >= MSG_LIST_MAX Or Len(s) >= MSG_LIST_CHARS Then
            s = s & "    ほか " & (d.Count - n) & " 件" & vbCrLf
            Exit For
        End If
        n = n + 1
        s = s & "    " & d(k) & vbCrLf
    Next k
    ListToText = s
End Function
