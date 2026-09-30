Attribute VB_Name = "LibInfoConv"
Option Explicit

'============================================================
' libInfo.txt 変換処理（手順書 ④ (1)～(9)）
'
'   CheckLibInfo   … 変換せずに危険箇所を調査する（★まずこれを実行）
'   ConvertLibInfo … 変換を実行し、タブ区切りテキストを出力する
'
'   入力 : libInfo.txt（Shift_JIS）
'   出力 : タブ区切りのテキストファイル
'          ※Excelシートには貼らない。テキストファイルとして出力する
'
'   詳細は README.md を参照
'============================================================

'==== 設定ここから ==========================================
Private Const SKIP_FIRST_LINE As Boolean = True        ' 1行目を削除する
Private Const IN_CHARSET      As String = "Shift_JIS"  ' 入力の文字コード
Private Const OUT_CHARSET     As String = "Shift_JIS"  ' 出力の文字コード ※未確認。UTF-8にするならここ
Private Const OUT_UTF8_BOM    As Boolean = False       ' UTF-8出力時にBOMを付けるか
Private Const SAFE_MODE_OLD   As Boolean = True        ' True = (7)"old" は削除せず、別ファイルに書き出すだけ

' 行末スペースを先に落とすか
'   False … 手順書どおりの順序（(1)タブ化 → (8)RTrim）。
'           行末に3個以上スペースがある行は、行末がタブで終わる。
'   True  … 先に行末スペースを落としてからタブ化する。
'           行末にタブが残らない。(8)の「★末尾の半角スペースを削除する★」
'           という意図には近いが、手作業の結果とは差分が出る。
'   ※CheckLibInfo の「変換後に行末がタブになる行」が 0 なら、どちらでも同じ
Private Const TRIM_SPACES_FIRST As Boolean = False
'==== 設定ここまで ==========================================

Private Const SPACE3 As String = "   "   ' 半角スペース3つ

'------------------------------------------------------------
' 削除パターン（手順書 (2)～(7)(9)）
'   ・すべて「部分一致」「大文字小文字を区別しない」
'   ・正規表現は使わない
'     (6)の bk. / org. の「.」を文字そのものとして扱う必要があるため。
'     正規表現にすると「.」が任意1文字になり、対象外のはずの
'     ORGXX まで削除してしまう。
'------------------------------------------------------------
Private Function DeletePatterns() As Variant
    Dim kopi As String
    ' 「コピー」。ChrW で組み立てているのは、この .bas を別の言語環境の
    ' Excel に取り込んだときに文字化けして一致しなくなるのを防ぐため。
    kopi = ChrW(&H30B3) & ChrW(&H30D4) & ChrW(&H30FC)

    DeletePatterns = Array( _
        "\log\", _
        "\save\", _
        "\common\tools\JudgedCIF\data\output", _
        kopi, _
        "bk.", _
        "_bk", _
        "bkup", _
        "bak", _
        "org.", _
        "old", _
        "origin.txt")
End Function

' パターンの表示名（メッセージ用。コピーだけ文字化け回避で別扱い）
Private Function PatternLabel(ByVal idx As Long) As String
    Dim p As Variant
    p = DeletePatterns()
    If idx = 3 Then
        PatternLabel = "kopi(" & p(idx) & ")"
    Else
        PatternLabel = p(idx)
    End If
End Function


'============================================================
' ファイル入出力（ADODB.Stream を使う）
'   VBA標準の Open / Line Input は文字コードを指定できず、
'   日本語が文字化けするので使わない
'============================================================
Private Function ReadTextFile(ByVal path As String, ByVal charset As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2                 ' adTypeText
    st.charset = charset
    st.Open
    st.LoadFromFile path
    ReadTextFile = st.ReadText(-1)   ' adReadAll
    st.Close
End Function

Private Sub WriteTextFile(ByVal path As String, ByVal text As String, _
                          ByVal charset As String, ByVal withBom As Boolean)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2
    st.charset = charset
    st.Open
    st.WriteText text

    If (UCase$(charset) = "UTF-8") And (Not withBom) Then
        ' ADODB.Stream はUTF-8で必ずBOMを付けるので、
        ' バイナリで読み直して先頭3バイトを捨てる
        Dim bin As Object
        st.Position = 0
        st.Type = 1             ' adTypeBinary
        st.Position = 3         ' BOM(EF BB BF)を飛ばす
        Set bin = CreateObject("ADODB.Stream")
        bin.Type = 1
        bin.Open
        st.CopyTo bin
        bin.SaveToFile path, 2  ' adSaveCreateOverWrite
        bin.Close
    Else
        st.SaveToFile path, 2
    End If
    st.Close
End Sub

' 改行コードを LF に揃えてから行に分割する（CRLF / CR / LF のどれでも可）
Private Function SplitLines(ByVal all As String) As Variant
    all = Replace(all, vbCrLf, vbLf)
    all = Replace(all, vbCr, vbLf)
    SplitLines = Split(all, vbLf)
End Function

' 入力ファイルを選ぶ
Private Function PickInputFile() As String
    Dim f As Variant
    f = Application.GetOpenFilename( _
            "テキストファイル (*.txt),*.txt,すべてのファイル (*.*),*.*", _
            1, "libInfo.txt を選択してください")
    If VarType(f) = vbBoolean Then
        PickInputFile = ""
    Else
        PickInputFile = CStr(f)
    End If
End Function

' 1行に対する変換 … 手順書(1)と(8)
'   (1) 半角スペース3つ -> タブ
'   (8) 行末の半角スペースを削除（RTrimで全部消す）
'   ※順序に注意。先にRTrimすると、行末スペースがタブに化けた後は
'     RTrimでは消せなくなる（RTrimはスペースしか消さない）
Private Function ConvertLine(ByVal s As String) As String
    If TRIM_SPACES_FIRST Then s = RTrim$(s)
    s = Replace(s, SPACE3, vbTab)
    ConvertLine = RTrim$(s)
End Function

' DeletePatterns の中から "old" の位置を探す（添字の直書きを避ける）
Private Function OldPatternIndex() As Long
    Dim p As Variant, j As Long
    p = DeletePatterns()
    OldPatternIndex = -1
    For j = LBound(p) To UBound(p)
        If CStr(p(j)) = "old" Then
            OldPatternIndex = j
            Exit Function
        End If
    Next j
End Function

' 拡張子を除いたパス（拡張子が無い場合もそのまま使えるように）
Private Function BaseNameOf(ByVal path As String) As String
    Dim dot As Long, sep As Long
    dot = InStrRev(path, ".")
    sep = InStrRev(path, "\")
    If dot > sep And dot > 0 Then
        BaseNameOf = Left$(path, dot - 1)
    Else
        BaseNameOf = path
    End If
End Function


'============================================================
' CheckLibInfo … 変換せずに調査するだけ
'   未確認事項を実データで確かめるための事前チェック
'============================================================
Public Sub CheckLibInfo()
    Dim inPath As String
    inPath = PickInputFile()
    If inPath = "" Then Exit Sub

    Dim lines As Variant
    lines = SplitLines(ReadTextFile(inPath, IN_CHARSET))

    Dim pat As Variant
    pat = DeletePatterns()

    Dim hit() As Long
    ReDim hit(LBound(pat) To UBound(pat))

    Dim i As Long, j As Long
    Dim s As String, conv As String
    Dim total As Long, blankLine As Long
    Dim sp4 As Long                 ' 4個以上連続するスペースがある行
    Dim tailSpace As Long           ' 行末にスペースがある行
    Dim maxTail As Long             ' 行末スペースの最大個数
    Dim folderCnt As Long           ' "folder" を含む行（(7)の誤爆候補）
    Dim tailTab As Long             ' 変換後に行末がタブになる行
    Dim colCnt As Object
    Set colCnt = CreateObject("Scripting.Dictionary")

    Dim startIdx As Long
    startIdx = LBound(lines)
    If SKIP_FIRST_LINE Then startIdx = startIdx + 1

    For i = startIdx To UBound(lines)
        s = lines(i)
        If Len(s) = 0 Then
            blankLine = blankLine + 1
        Else
            total = total + 1

            ' 4個以上連続するスペース
            If InStr(1, s, "    ", vbBinaryCompare) > 0 Then sp4 = sp4 + 1

            ' 行末スペース
            Dim n As Long
            n = Len(s) - Len(RTrim$(s))
            If n > 0 Then
                tailSpace = tailSpace + 1
                If n > maxTail Then maxTail = n
            End If

            ' (7)の誤爆候補
            If InStr(1, s, "folder", vbTextCompare) > 0 Then folderCnt = folderCnt + 1

            ' 各削除パターンの該当数
            For j = LBound(pat) To UBound(pat)
                If InStr(1, s, CStr(pat(j)), vbTextCompare) > 0 Then
                    hit(j) = hit(j) + 1
                End If
            Next j

            ' 変換後の列数と行末タブ
            conv = ConvertLine(s)
            Dim c As Long
            c = UBound(Split(conv, vbTab)) + 1
            If colCnt.Exists(c) Then
                colCnt(c) = colCnt(c) + 1
            Else
                colCnt.Add c, 1
            End If
            If Right$(conv, 1) = vbTab Then tailTab = tailTab + 1
        End If
    Next i

    ' ---- 結果を組み立てる ----
    Dim msg As String
    msg = "【調査結果】" & vbCrLf & inPath & vbCrLf & vbCrLf
    msg = msg & "全行数(1行目除く) : " & total & vbCrLf
    msg = msg & "空行              : " & blankLine & vbCrLf & vbCrLf

    msg = msg & "--- 未確認事項の実データ確認 ---" & vbCrLf
    msg = msg & "4個以上連続するスペースを含む行 : " & sp4
    If sp4 > 0 Then msg = msg & "  ★要確認（タブが増えて列がずれる）"
    msg = msg & vbCrLf
    msg = msg & "行末にスペースがある行          : " & tailSpace & _
                "（最大 " & maxTail & " 個）" & vbCrLf
    msg = msg & "変換後に行末がタブになる行      : " & tailTab
    If tailTab > 0 Then msg = msg & "  ★要確認"
    msg = msg & vbCrLf
    msg = msg & "'folder' を含む行               : " & folderCnt
    If folderCnt > 0 Then msg = msg & "  ★★(7)の'old'で巻き添え削除される"
    msg = msg & vbCrLf & vbCrLf

    msg = msg & "--- 変換後のタブ区切り列数 ---" & vbCrLf
    Dim k As Variant
    For Each k In colCnt.Keys
        msg = msg & "  " & k & "列 : " & colCnt(k) & " 行" & vbCrLf
    Next k
    msg = msg & vbCrLf

    msg = msg & "--- 削除パターン別の該当行数 ---" & vbCrLf
    Dim sumHit As Long
    For j = LBound(pat) To UBound(pat)
        msg = msg & "  " & PatternLabel(j) & " : " & hit(j) & vbCrLf
        sumHit = sumHit + hit(j)
    Next j
    msg = msg & "  （重複あり。合計 " & sumHit & " 件）" & vbCrLf

    MsgBox msg, vbInformation, "CheckLibInfo"
    Debug.Print msg
End Sub


'============================================================
' ConvertLibInfo … 変換を実行する
'============================================================
Public Sub ConvertLibInfo()
    Dim inPath As String
    inPath = PickInputFile()
    If inPath = "" Then Exit Sub

    Dim lines As Variant
    lines = SplitLines(ReadTextFile(inPath, IN_CHARSET))

    Dim pat As Variant
    pat = DeletePatterns()
    Dim oldIdx As Long
    oldIdx = OldPatternIndex()

    Dim hit() As Long
    ReDim hit(LBound(pat) To UBound(pat))

    Dim outBuf() As String
    ReDim outBuf(0 To UBound(lines))
    Dim oldBuf() As String
    ReDim oldBuf(0 To UBound(lines))

    Dim nOut As Long, nOld As Long
    Dim total As Long, deleted As Long, blankLine As Long

    Dim i As Long, j As Long
    Dim s As String
    Dim killed As Boolean
    Dim killedByOldOnly As Boolean

    Dim startIdx As Long
    startIdx = LBound(lines)
    If SKIP_FIRST_LINE Then startIdx = startIdx + 1

    For i = startIdx To UBound(lines)
        s = lines(i)
        If Len(s) = 0 Then
            blankLine = blankLine + 1
        Else
            total = total + 1
            killed = False
            killedByOldOnly = False

            For j = LBound(pat) To UBound(pat)
                If InStr(1, s, CStr(pat(j)), vbTextCompare) > 0 Then
                    hit(j) = hit(j) + 1
                    If j = oldIdx Then
                        killedByOldOnly = True
                    Else
                        killed = True
                    End If
                End If
            Next j

            ' SAFE_MODE_OLD = True のときは "old" だけに当たった行は削除しない。
            ' 代わりに別ファイルに書き出して、人が目で確認できるようにする。
            ' （"old" は folder / holder などにも部分一致してしまうため）
            If killedByOldOnly Then
                oldBuf(nOld) = s
                nOld = nOld + 1
                If Not SAFE_MODE_OLD Then killed = True
            End If

            If killed Then
                deleted = deleted + 1
            Else
                outBuf(nOut) = ConvertLine(s)
                nOut = nOut + 1
            End If
        End If
    Next i

    ' ---- 出力 ----
    Dim baseName As String, outPath As String, oldPath As String
    baseName = BaseNameOf(inPath)
    outPath = baseName & "_converted.txt"
    oldPath = baseName & "_old_candidates.txt"

    Dim res As String
    If nOut > 0 Then
        ReDim Preserve outBuf(0 To nOut - 1)
        res = Join(outBuf, vbCrLf) & vbCrLf
    Else
        res = ""
    End If
    WriteTextFile outPath, res, OUT_CHARSET, OUT_UTF8_BOM

    If nOld > 0 Then
        ReDim Preserve oldBuf(0 To nOld - 1)
        WriteTextFile oldPath, Join(oldBuf, vbCrLf) & vbCrLf, OUT_CHARSET, OUT_UTF8_BOM
    End If

    ' ---- 結果表示 ----
    Dim msg As String
    msg = "【変換完了】" & vbCrLf & vbCrLf
    msg = msg & "入力 : " & inPath & vbCrLf
    msg = msg & "出力 : " & outPath & vbCrLf & vbCrLf
    msg = msg & "読込   : " & total & " 行（1行目と空行 " & blankLine & " 行を除く）" & vbCrLf
    msg = msg & "削除   : " & deleted & " 行" & vbCrLf
    msg = msg & "出力   : " & nOut & " 行" & vbCrLf & vbCrLf

    msg = msg & "--- 削除パターン別の該当行数 ---" & vbCrLf
    For j = LBound(pat) To UBound(pat)
        msg = msg & "  " & PatternLabel(j) & " : " & hit(j)
        If j = oldIdx And SAFE_MODE_OLD Then msg = msg & "  ※削除せず別ファイルへ"
        msg = msg & vbCrLf
    Next j
    msg = msg & "（1行が複数パターンに当たることがあるため合計は一致しません）" & vbCrLf

    If nOld > 0 Then
        msg = msg & vbCrLf & "★ 'old' のみに該当した " & nOld & " 行は削除していません。" & vbCrLf
        msg = msg & "   " & oldPath & vbCrLf
        msg = msg & "   を開いて、本当に削除してよいか確認してください。" & vbCrLf
        msg = msg & "   （folder / holder などに部分一致している可能性があります）" & vbCrLf
        msg = msg & "   確認後、本当に削除するなら SAFE_MODE_OLD を False にして再実行。"
    End If

    MsgBox msg, vbInformation, "ConvertLibInfo"
    Debug.Print msg
End Sub
