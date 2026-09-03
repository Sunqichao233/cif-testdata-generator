<#
.SYNOPSIS
    CIF情報ファイル テストデータ生成ツール (Amazon WorkSpaces 用)

.DESCRIPTION
    設計書「CIF情報ファイルの項目」に基づき、タブ区切り13項目/1レコードの
    テストデータファイルを生成し、生成結果を自己検証します。

.EXAMPLE
    .\Generate-CifTestData.ps1
    .\Generate-CifTestData.ps1 -Count 500 -DataProfile full
    .\Generate-CifTestData.ps1 -Target 001,002 -Encoding utf-8
    .\Generate-CifTestData.ps1 -ValidateOnly
#>
[CmdletBinding()]
param(
    [string]   $ConfigPath,
    [int]      $Count = 0,
    [ValidateSet('minimal','full','matching')][string] $DataProfile,
    [string]   $OutputDir,
    [ValidateSet('shift_jis','utf-8','utf-8-bom')][string] $Encoding,
    [string[]] $Target,
    [switch]   $IncludeBoundary,
    [switch]   $ValidateOnly,
    [switch]   $NoValidate
)

$ErrorActionPreference = 'Stop'

# =============================================================
# 0. 項目定義 (設計書「CIF情報ファイルの項目」より)
# =============================================================
$Script:FieldDefs = @(
    [pscustomobject]@{ No=1;  Name='レコードID';                 Max=30;  Fixed=$false; Type='c'; Required=$true  }
    [pscustomobject]@{ No=2;  Name='タイプ';                     Max=1;   Fixed=$false; Type='c'; Required=$false }
    [pscustomobject]@{ No=3;  Name='名称（フリガナ）';           Max=200; Fixed=$false; Type='k'; Required=$false }
    [pscustomobject]@{ No=4;  Name='名称（漢字）';               Max=200; Fixed=$false; Type='j'; Required=$false }
    [pscustomobject]@{ No=5;  Name='名称（英字）';               Max=200; Fixed=$false; Type='c'; Required=$false }
    [pscustomobject]@{ No=6;  Name='電話番号';                   Max=30;  Fixed=$false; Type='c'; Required=$false }
    [pscustomobject]@{ No=7;  Name='ID情報';                     Max=30;  Fixed=$false; Type='c'; Required=$false }
    [pscustomobject]@{ No=8;  Name='ID情報種類';                 Max=2;   Fixed=$true;  Type='n'; Required=$false }
    [pscustomobject]@{ No=9;  Name='国内郵便番号';               Max=7;   Fixed=$true;  Type='n'; Required=$false }
    [pscustomobject]@{ No=10; Name='国内住所（市区町村以降）';   Max=200; Fixed=$false; Type='j'; Required=$false }
    [pscustomobject]@{ No=11; Name='英文住所';                   Max=256; Fixed=$false; Type='c'; Required=$false }
    [pscustomobject]@{ No=12; Name='設立/生年月日';              Max=8;   Fixed=$true;  Type='n'; Required=$false }
    [pscustomobject]@{ No=13; Name='部店コード';                 Max=32;  Fixed=$false; Type='c'; Required=$true  }
)

# 文字種: c=半角英数記号(半角スペース可) / k=全角カタカナ(半角スペース可) / j=全角文字 / n=半角数値
$Script:CharPattern = @{
    'c' = '^[\x20-\x7E]*$'
    'k' = '^[ァ-ヺー・ ]*$'
    'j' = '^[^\t\r\n]*$'
    'n' = '^[0-9]*$'
}

# =============================================================
# 1. ログ出力
# =============================================================
$Script:LogLines = New-Object System.Collections.Generic.List[string]

function Write-Log {
    param([string]$Message, [ValidateSet('INFO','WARN','ERROR','OK','HEAD')][string]$Level='INFO')
    $stamp = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    [void]$Script:LogLines.Add("[$stamp][$Level] $Message")
    switch ($Level) {
        'ERROR' { Write-Host "  [NG]   $Message" -ForegroundColor Red }
        'WARN'  { Write-Host "  [WARN] $Message" -ForegroundColor Yellow }
        'OK'    { Write-Host "  [OK]   $Message" -ForegroundColor Green }
        'HEAD'  { Write-Host ""; Write-Host $Message -ForegroundColor Cyan }
        default { Write-Host "  $Message" }
    }
}

# =============================================================
# 2. 設定読み込み
# =============================================================
$Script:Root = Split-Path -Parent $PSScriptRoot
if (-not $ConfigPath) { $ConfigPath = Join-Path $Script:Root 'config\config.json' }
if (-not (Test-Path $ConfigPath)) { throw "設定ファイルが見つかりません: $ConfigPath" }

$cfg = [System.IO.File]::ReadAllText($ConfigPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json

# コマンドライン引数による上書き
if ($DataProfile) { $cfg.generation.profile = $DataProfile }
if ($OutputDir)   { $cfg.output.directory   = $OutputDir }
if ($Encoding)    { $cfg.output.encoding    = $Encoding }
if ($Count -gt 0) { foreach ($t in $cfg.targets) { $t.count = $Count } }

function Resolve-Dir {
    param([string]$Path)
    if ([System.IO.Path]::IsPathRooted($Path)) { return $Path }
    return (Join-Path $Script:Root $Path)
}
$outRoot = Resolve-Dir $cfg.output.directory
$logRoot = Resolve-Dir $cfg.output.logDirectory
foreach ($d in @($outRoot, $logRoot)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Path $d -Force | Out-Null }
}

function Get-TextEncoding {
    param([string]$Name)
    switch ($Name.ToLower()) {
        'shift_jis' { return [System.Text.Encoding]::GetEncoding(932) }
        'sjis'      { return [System.Text.Encoding]::GetEncoding(932) }
        'cp932'     { return [System.Text.Encoding]::GetEncoding(932) }
        'utf-8'     { return New-Object System.Text.UTF8Encoding($false) }
        'utf-8-bom' { return New-Object System.Text.UTF8Encoding($true) }
        default     { throw "未対応の文字コードです: $Name" }
    }
}

# =============================================================
# 3. ダミーデータ素材
# =============================================================
$Script:Surnames = @(
    @('ヤマダ','山田','YAMADA'),      @('サトウ','佐藤','SATO'),        @('スズキ','鈴木','SUZUKI'),
    @('タカハシ','高橋','TAKAHASHI'), @('タナカ','田中','TANAKA'),      @('イトウ','伊藤','ITO'),
    @('ワタナベ','渡辺','WATANABE'),  @('ナカムラ','中村','NAKAMURA'),  @('コバヤシ','小林','KOBAYASHI'),
    @('カトウ','加藤','KATO'),        @('ヨシダ','吉田','YOSHIDA'),     @('ヤマモト','山本','YAMAMOTO'),
    @('ササキ','佐々木','SASAKI'),    @('ヤマグチ','山口','YAMAGUCHI'), @('マツモト','松本','MATSUMOTO'),
    @('イノウエ','井上','INOUE'),     @('キムラ','木村','KIMURA'),      @('ハヤシ','林','HAYASHI'),
    @('シミズ','清水','SHIMIZU'),     @('ヤマザキ','山崎','YAMAZAKI')
)
$Script:GivenNames = @(
    @('タロウ','太郎','TARO'),      @('ジロウ','次郎','JIRO'),      @('ハナコ','花子','HANAKO'),
    @('イチロウ','一郎','ICHIRO'),  @('ミサキ','美咲','MISAKI'),    @('ユウタ','悠太','YUTA'),
    @('サクラ','咲良','SAKURA'),    @('ケンジ','健二','KENJI'),     @('アヤカ','彩香','AYAKA'),
    @('ダイスケ','大輔','DAISUKE'), @('ナオコ','直子','NAOKO'),     @('ショウタ','翔太','SHOTA'),
    @('ユカリ','由香里','YUKARI'),  @('タクミ','拓海','TAKUMI'),    @('マナブ','学','MANABU'),
    @('リョウ','涼','RYO'),         @('エミ','恵美','EMI'),         @('コウジ','浩二','KOJI'),
    @('チヒロ','千尋','CHIHIRO'),   @('ノブオ','信夫','NOBUO')
)
# 郵便番号(7桁) / 国内住所(市区町村以降) / 英文住所(国名は記入しない)
$Script:Addresses = @(
    @('2200011','横浜市西区高島２－１－１','2-1-1 Takashima, Nishi-ku, Yokohama'),
    @('2310021','横浜市中区日本大通１１','11 Nihon-odori, Naka-ku, Yokohama'),
    @('1000005','千代田区丸の内１－９－１','1-9-1 Marunouchi, Chiyoda-ku, Tokyo'),
    @('5420076','大阪市中央区難波３－１－１','3-1-1 Namba, Chuo-ku, Osaka'),
    @('4600008','名古屋市中区栄３－１５－３３','3-15-33 Sakae, Naka-ku, Nagoya'),
    @('0600042','札幌市中央区大通西４－１','4-1 Odori-nishi, Chuo-ku, Sapporo'),
    @('8100001','福岡市中央区天神２－１１－１','2-11-1 Tenjin, Chuo-ku, Fukuoka'),
    @('9800021','仙台市青葉区中央１－１－１','1-1-1 Chuo, Aoba-ku, Sendai')
)
# ID情報種類: 運転免許証・旅券などに対応するコード値 (導入時ヒアリングシートで確定 / 暫定値)
$Script:IdKinds = @('01','02','03','04','05')

$Script:KanaDigits    = @('ア','イ','ウ','エ','オ','カ','キ','ク','ケ','コ')
$Script:ZenkakuDigits = @('０','１','２','３','４','５','６','７','８','９')

function ConvertTo-KanaNumber {
    param([int]$Value)
    $b = ''
    foreach ($c in ([string]$Value).ToCharArray()) { $b += $Script:KanaDigits[[int]::Parse($c)] }
    return $b
}
function ConvertTo-ZenkakuNumber {
    param([int]$Value)
    $b = ''
    foreach ($c in ([string]$Value).ToCharArray()) { $b += $Script:ZenkakuDigits[[int]::Parse($c)] }
    return $b
}

# =============================================================
# 4. レコード生成
# =============================================================
function New-CifRecord {
    # 1レコード(13項目の配列)を生成する。$Index はファイル内0起算、$Seq は全体通番。
    param(
        [int]$Index, [int]$Seq, [string]$ButenCode,
        [string]$Profile, [int]$RecordIdDigits, [bool]$SetRecordId
    )

    $sCount = $Script:Surnames.Count
    $gCount = $Script:GivenNames.Count
    $combo  = $sCount * $gCount
    $sn     = $Script:Surnames[$Index % $sCount]
    $gn     = $Script:GivenNames[[int][math]::Floor($Index / $sCount) % $gCount]
    $cycle  = [int][math]::Floor($Index / $combo)   # 氏名の重複を避けるための連番サイクル

    $rec = New-Object 'string[]' 13
    for ($i = 0; $i -lt 13; $i++) { $rec[$i] = '' }

    # No.1 レコードID : 連番 (検索対象者はシステム自動採番のため任意)
    if ($SetRecordId) { $rec[0] = ([string]$Seq).PadLeft($RecordIdDigits, '0') }

    # No.3 名称（フリガナ） : 同一ファイル内で重複しない (半角スペース区切り)
    $kana = "$($sn[0]) $($gn[0])"
    if ($cycle -gt 0) { $kana += ' ' + (ConvertTo-KanaNumber $cycle) }
    $rec[2] = $kana

    # No.13 部店コード : 必須
    $rec[12] = $ButenCode

    if ($Profile -eq 'full') {
        $adr = $Script:Addresses[$Index % $Script:Addresses.Count]

        $rec[1]  = 'I'                                                      # No.2  タイプ (個人)
        $rec[3]  = "$($sn[1])　$($gn[1])"                                   # No.4  名称（漢字）
        $rec[4]  = "$($sn[2]) $($gn[2])"                                    # No.5  名称（英字）
        if ($cycle -gt 0) {
            $rec[3] += (ConvertTo-ZenkakuNumber $cycle)
            $rec[4] += " $cycle"
        }
        $rec[5]  = '045' + (($Seq % 10000000).ToString().PadLeft(7,'0'))    # No.6  電話番号
        $rec[6]  = 'ID' + (($Seq % 100000000).ToString().PadLeft(8,'0'))    # No.7  ID情報
        $rec[7]  = $Script:IdKinds[$Index % $Script:IdKinds.Count]          # No.8  ID情報種類 (固定2桁)
        $rec[8]  = $adr[0]                                                  # No.9  国内郵便番号 (固定7桁)
        $rec[9]  = $adr[1]                                                  # No.10 国内住所（市区町村以降）
        $rec[10] = $adr[2]                                                  # No.11 英文住所
        $y = 1950 + ($Index % 56); $m = 1 + ($Index % 12); $d = 1 + ($Index % 28)
        $rec[11] = '{0:0000}{1:00}{2:00}' -f $y, $m, $d                     # No.12 設立/生年月日 (YYYYMMDD)
    }
    return ,$rec
}

function New-BoundaryRecords {
    # 最大桁数ちょうどの境界値レコード(正常系)を生成する。
    param([int]$StartSeq, [string]$ButenCode, [bool]$SetRecordId)
    $list = New-Object System.Collections.ArrayList

    $r1 = New-Object 'string[]' 13
    for ($i = 0; $i -lt 13; $i++) { $r1[$i] = '' }
    if ($SetRecordId) { $r1[0] = ('B' + [string]$StartSeq).PadRight(30, '9') }
    $r1[2]  = ('キョウカイチ' * 34).Substring(0, 200)          # フリガナ 200文字ちょうど
    $r1[12] = $ButenCode
    [void]$list.Add($r1)

    $r2 = New-Object 'string[]' 13
    for ($i = 0; $i -lt 13; $i++) { $r2[$i] = '' }
    if ($SetRecordId) { $r2[0] = ('B' + [string]($StartSeq + 1)).PadRight(30, '8') }
    $r2[1]  = 'I'
    $r2[2]  = ('キョウカイチ' * 34).Substring(0, 199) + 'ニ'    # フリガナ 200文字 (重複回避)
    $r2[3]  = ('境界値検証用漢字氏名' * 20).Substring(0, 200)   # 漢字 200文字ちょうど
    $r2[4]  = ('BOUNDARY MAX LENGTH NAME ' * 10).Substring(0, 200)
    $r2[5]  = ''.PadRight(30, '9')                              # 電話番号 30桁
    $r2[6]  = ''.PadRight(30, 'Z')                              # ID情報 30桁
    $r2[7]  = '01'
    $r2[8]  = '2200011'
    $r2[9]  = ('横浜市西区高島２－１－１' * 20).Substring(0, 200)
    $r2[10] = ('2-1-1 Takashima, Nishi-ku, Yokohama ' * 10).Substring(0, 256)
    $r2[11] = '19800101'
    $r2[12] = $ButenCode.PadRight(32, '0').Substring(0, 32)     # 部店コード 32桁
    [void]$list.Add($r2)

    return $list
}

function New-MatchingRecords {
    # 照合パターン(名寄せ)テスト用レコードを生成する。
    #   レコードID = 1<f1>2<f2>3<f3>[h]
    #     f1 : 名称（フリガナ）  A=一致 / B=不一致
    #     f2 : 名称（漢字）      A=一致 / B=不一致 / C=未設定
    #     f3 : 設立/生年月日     A=一致 / B=不一致 / C=未設定
    #     h  : ヒット想定（フリガナ一致 かつ 漢字が不一致でない かつ 生年月日が不一致でない）
    param($MatchCfg, [string]$ButenCode)

    $list = New-Object System.Collections.ArrayList
    foreach ($f1 in @('A','B')) {
        foreach ($f3 in @('A','B','C')) {
            foreach ($f2 in @('A','B','C')) {
                $hit = ($f1 -eq 'A') -and ($f2 -ne 'B') -and ($f3 -ne 'B')
                $id  = '1' + $f1 + '2' + $f2 + '3' + $f3
                if ($hit) { $id += 'h' }

                $rec = New-Object 'string[]' 13
                for ($i = 0; $i -lt 13; $i++) { $rec[$i] = '' }
                $rec[0]  = $id
                $rec[2]  = $MatchCfg.furigana.$f1
                $rec[3]  = $MatchCfg.kanji.$f2
                $rec[11] = $MatchCfg.birthday.$f3
                $rec[12] = $ButenCode
                [void]$list.Add($rec)
            }
        }
    }
    return $list
}

function Add-ExtraColumns {
    # 設計書 No.14 以降の追加列を末尾に連結する (config.extraColumns)。
    param([string[]]$Record, $ExtraCfg)
    if (-not $ExtraCfg -or -not $ExtraCfg.enabled) { return $Record }
    $extra = @($ExtraCfg.columns | ForEach-Object { [string]$_.value })
    return @($Record + $extra)
}

# =============================================================
# 5. 検証
# =============================================================
function Test-CifFile {
    param([string]$Path, [string]$EncodingName, [string]$LengthMode, [bool]$RequireRecordId,
          $ValidationCfg, [int]$ExpectedFieldCount = 13)

    $enc   = Get-TextEncoding $EncodingName
    $text  = $enc.GetString([System.IO.File]::ReadAllBytes($Path)).TrimEnd("`r", "`n")
    $lines = $text -split "`r`n|`n"
    $sjis  = [System.Text.Encoding]::GetEncoding(932)

    $errors  = New-Object System.Collections.Generic.List[string]
    $idSet   = New-Object 'System.Collections.Generic.HashSet[string]'
    $kanaSet = New-Object 'System.Collections.Generic.HashSet[string]'
    $lineNo  = 0

    foreach ($line in $lines) {
        $lineNo++
        if ($line -eq '') { [void]$errors.Add("L$lineNo : 空行が含まれています"); continue }

        $f = $line -split "`t", -1
        if ($f.Count -ne $ExpectedFieldCount) {
            [void]$errors.Add("L$lineNo : 項目数が$($ExpectedFieldCount)ではありません (実際: $($f.Count))"); continue
        }

        foreach ($def in $Script:FieldDefs) {
            $v = $f[$def.No - 1]
            if ($v -eq '') {
                if ($def.Required -and ($def.No -ne 1 -or $RequireRecordId)) {
                    [void]$errors.Add("L$lineNo No.$($def.No) $($def.Name) : 必須項目が未設定です")
                }
                continue
            }
            $len = if ($LengthMode -eq 'byte') { $sjis.GetByteCount($v) } else { $v.Length }
            if ($def.Fixed) {
                if ($len -ne $def.Max) {
                    [void]$errors.Add("L$lineNo No.$($def.No) $($def.Name) : 固定$($def.Max)桁ではありません (実際: $len)")
                }
            } elseif ($len -gt $def.Max) {
                [void]$errors.Add("L$lineNo No.$($def.No) $($def.Name) : 最大桁数$($def.Max)超過 (実際: $len)")
            }
            if ($v -notmatch $Script:CharPattern[$def.Type]) {
                [void]$errors.Add("L$lineNo No.$($def.No) $($def.Name) : 文字種[$($def.Type)]に違反しています")
            }
            if ($def.No -eq 12) {
                $dt = [datetime]::MinValue
                if (-not [datetime]::TryParseExact($v, 'yyyyMMdd', [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::None, [ref]$dt)) {
                    [void]$errors.Add("L$lineNo No.12 設立/生年月日 : 存在しない日付です ($v)")
                }
            }
            if ($def.No -eq 1 -and $v -ne $v.TrimEnd(' ')) {
                [void]$errors.Add("L$lineNo No.1 レコードID : 後続スペースはカットしてください")
            }
        }

        if ($f[0] -ne '' -and $ValidationCfg.checkDuplicateRecordId) {
            if (-not $idSet.Add($f[0])) { [void]$errors.Add("L$lineNo No.1 レコードID : ファイル内で重複しています ($($f[0]))") }
        }
        if ($f[2] -ne '' -and $ValidationCfg.checkDuplicateFurigana) {
            if (-not $kanaSet.Add($f[2])) { [void]$errors.Add("L$lineNo No.3 名称（フリガナ） : ファイル内で重複しています ($($f[2]))") }
        }
        if ($f[2] -eq '' -and $f[3] -eq '' -and $f[4] -eq '') {
            [void]$errors.Add("L$lineNo No.3/4/5 名称 : フリガナ・漢字・英字のうち少なくとも1項目は必須です")
        }
    }

    return [pscustomobject]@{ Path = $Path; RecordCount = $lineNo; Errors = $errors }
}

# =============================================================
# 6. メイン処理
# =============================================================
$startTime = Get-Date
Write-Host ''
Write-Host '==================================================================' -ForegroundColor White
Write-Host ' CIF情報ファイル テストデータ生成ツール' -ForegroundColor White
Write-Host '==================================================================' -ForegroundColor White
Write-Log "設定ファイル : $ConfigPath"
Write-Log "出力先       : $outRoot"
Write-Log "文字コード   : $($cfg.output.encoding) / 改行 : $($cfg.output.newline)"
Write-Log "プロファイル : $($cfg.generation.profile)"

$enc     = Get-TextEncoding $cfg.output.encoding
$nl      = if ($cfg.output.newline -eq 'LF') { "`n" } else { "`r`n" }
$dateTag = if ($cfg.output.useDateInName) { '_' + (Get-Date).ToString('yyyyMMdd') + '_' } else { '_' }

$extraCfg   = $cfg.extraColumns
$extraCount = 0
if ($extraCfg -and $extraCfg.enabled) {
    $extraCount = @($extraCfg.columns).Count
    Write-Log "追加列       : 有効 (No.14〜No.$(13 + $extraCount) / 計 $(13 + $extraCount) 項目)" 'WARN'
}

if ($cfg.generation.profile -eq 'matching') {
    Write-Log '照合パターン : 2x3x3 = 18 レコード/ファイル (config.json の count は使用しません)'
    if ($cfg.validation.checkDuplicateFurigana) {
        Write-Log '名称（フリガナ）の重複チェックは照合パターンでは意図的に重複するため無効化します' 'WARN'
        $cfg.validation.checkDuplicateFurigana = $false
    }
}

$targets = @($cfg.targets | Where-Object { $_.enabled })
if ($Target) { $targets = @($targets | Where-Object { $Target -contains $_.id }) }
if ($targets.Count -eq 0) { throw '対象ファイルが1件もありません。config.json の targets または -Target 引数を確認してください。' }

$results = @()
$seq = [int]$cfg.generation.startNumber

foreach ($t in $targets) {
    $prefix   = if ($t.kind -eq 'search') { $cfg.output.searchPrefix } else { $cfg.output.cifPrefix }
    $fileName = "$prefix$dateTag$($t.id)$($cfg.output.extension)"
    $filePath = Join-Path $outRoot $fileName
    $setRecId = if ($t.kind -eq 'search') { [bool]$cfg.generation.setRecordIdForSearch } else { $true }

    Write-Log "[$($t.id)] $($t.label)  部店コード=$($t.butenCode)  件数=$($t.count)  -> $fileName" 'HEAD'

    if (-not $ValidateOnly) {
        if ((Test-Path $filePath) -and (-not $cfg.generation.overwrite)) {
            Write-Log "既存ファイルが存在するためスキップしました : $fileName" 'WARN'
            continue
        }

        $sb = New-Object System.Text.StringBuilder
        if ($cfg.generation.profile -eq 'matching') {
            $mButen = if ($cfg.matching.butenCodeOverride) { $cfg.matching.butenCodeOverride } else { $t.butenCode }
            foreach ($mr in (New-MatchingRecords -MatchCfg $cfg.matching -ButenCode $mButen)) {
                [void]$sb.Append(((Add-ExtraColumns -Record $mr -ExtraCfg $extraCfg) -join "`t")).Append($nl)
                $seq++
            }
        } else {
            for ($i = 0; $i -lt [int]$t.count; $i++) {
                $rec = New-CifRecord -Index $i -Seq $seq -ButenCode $t.butenCode `
                                     -Profile $cfg.generation.profile `
                                     -RecordIdDigits ([int]$cfg.generation.recordIdDigits) `
                                     -SetRecordId $setRecId
                [void]$sb.Append(((Add-ExtraColumns -Record $rec -ExtraCfg $extraCfg) -join "`t")).Append($nl)
                $seq++
            }
            if ($IncludeBoundary) {
                foreach ($br in (New-BoundaryRecords -StartSeq $seq -ButenCode $t.butenCode -SetRecordId $setRecId)) {
                    [void]$sb.Append(((Add-ExtraColumns -Record $br -ExtraCfg $extraCfg) -join "`t")).Append($nl)
                    $seq++
                }
            }
        }
        $outText = $sb.ToString()
        if (-not $cfg.output.trailingNewline) { $outText = $outText.TrimEnd("`r", "`n") }
        [System.IO.File]::WriteAllBytes($filePath, $enc.GetBytes($outText))
        $sizeKb = [math]::Round((Get-Item $filePath).Length / 1KB, 1)
        Write-Log "生成完了 : $fileName ($sizeKb KB)" 'OK'
    }

    if ((-not $NoValidate) -and $cfg.validation.enabled) {
        if (-not (Test-Path $filePath)) { Write-Log "検証対象ファイルがありません : $fileName" 'WARN'; continue }
        $r = Test-CifFile -Path $filePath -EncodingName $cfg.output.encoding `
                          -LengthMode $cfg.generation.lengthCheckMode `
                          -RequireRecordId $setRecId -ValidationCfg $cfg.validation `
                          -ExpectedFieldCount (13 + $extraCount)
        $results += [pscustomobject]@{ Id = $t.id; File = $fileName; Records = $r.RecordCount; ErrorCount = $r.Errors.Count }
        if ($r.Errors.Count -eq 0) {
            Write-Log "検証OK : $($r.RecordCount) 件すべて項目定義に適合しています" 'OK'
        } else {
            Write-Log "検証NG : $($r.Errors.Count) 件のエラーを検出しました" 'ERROR'
            $r.Errors | Select-Object -First 20 | ForEach-Object { Write-Log $_ 'ERROR' }
            if ($r.Errors.Count -gt 20) { Write-Log "... 他 $($r.Errors.Count - 20) 件 (詳細はログファイルを参照)" 'ERROR' }
            $r.Errors | ForEach-Object { [void]$Script:LogLines.Add("[detail] $_") }
            if ($cfg.validation.stopOnError) { throw '検証エラーのため処理を中断しました。' }
        }
    }
}

# =============================================================
# 7. サマリ
# =============================================================
Write-Host ''
Write-Host '------------------------------ 集計 ------------------------------' -ForegroundColor White
if ($results.Count -gt 0) { $results | Format-Table -AutoSize | Out-String | Write-Host }

$totalErr = ($results | Measure-Object -Property ErrorCount -Sum).Sum
$totalRec = ($results | Measure-Object -Property Records    -Sum).Sum
$elapsed  = ((Get-Date) - $startTime).TotalSeconds

Write-Log "総レコード数 : $totalRec 件 / 出力ファイル数 : $($results.Count)"
Write-Log ('処理時間     : {0:N1} 秒' -f $elapsed)

$logPath = Join-Path $logRoot ('generate_{0}.log' -f (Get-Date).ToString('yyyyMMdd_HHmmss'))
[System.IO.File]::WriteAllLines($logPath, $Script:LogLines, (New-Object System.Text.UTF8Encoding($true)))
Write-Host "  ログ         : $logPath"
Write-Host ''

if ($totalErr -gt 0) {
    Write-Host " 結果 : NG ($totalErr 件のエラー)" -ForegroundColor Red
    exit 1
}
Write-Host ' 結果 : OK' -ForegroundColor Green
exit 0
