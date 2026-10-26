# ======================================================================
# 検索インデックス作成ツール（PowerShell 5.1用）
# 指定フォルダ配下のフォルダ名・ファイル名・サイズ・更新日を調べ、
# データを埋め込んだ「検索画面.html」を1枚出力します。
# 出力されたHTMLをNASに置けば、ダブルクリックで誰でも検索できます。
# ======================================================================
param(
    [string]$TargetFolder = "",
    [string]$OutputDirectory = ""
)

# ■1. 設定

# 一覧から除外する拡張子（名前だけの記録なので、画像・動画なども通常は含めてかまいません）
$ExcludeExtensions = @(".tmp", ".db", ".ini", ".sys", ".lnk", ".bak")

# 一覧から除外するファイル名の先頭（Officeの一時ファイルなど）
$ExcludeNamePrefixes = @("~$")

# 中に入らないフォルダ名（ごみ箱・NASの管理用フォルダなど）
$ExcludeDirNames = @('$RECYCLE.BIN', "System Volume Information", "#recycle", "@eaDir", ".snapshot")

# 出力ファイル名
$OutputFileName = "検索画面.html"

# ■2. 補助関数

function Select-Folder {
    param([string]$Description)
    Add-Type -AssemblyName System.Windows.Forms
    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Description
    $dialog.ShowNewFolderButton = $true
    try {
        if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) { return $dialog.SelectedPath }
        return $null
    } finally { $dialog.Dispose() }
}

# JavaScriptの文字列リテラル用にエスケープする（</script> 対策の "</" も処理）
function Esc-Js {
    param([string]$s)
    return $s.Replace('\', '\\').Replace('"', '\"').Replace("`t", '\t').Replace("`r", '\r').Replace("`n", '\n').Replace('</', '<\/').Replace([string][char]0x2028, '\u2028').Replace([string][char]0x2029, '\u2029')
}

# ■3. 実行

$templatePath = Join-Path $PSScriptRoot "検索画面テンプレート.html"
if (-not (Test-Path -LiteralPath $templatePath)) {
    Write-Host "テンプレートが見つかりません: $templatePath" -ForegroundColor Red
    return
}

if (-not $TargetFolder) {
    Write-Host "対象フォルダのパスを貼り付けてEnterを押してください。" -ForegroundColor Cyan
    Write-Host "（NASは \\サーバー名\共有名 の形式。何も入力せずEnterで選択画面が開きます）"
    $typed = (Read-Host "対象フォルダ").Trim().Trim('"')
    if ($typed) { $TargetFolder = $typed }
    else { $TargetFolder = Select-Folder "検索対象にする【フォルダ】を選択してください" }
}
if (-not $TargetFolder) { Write-Host "対象フォルダが選択されなかったため終了します。" -ForegroundColor Yellow; return }
if (-not (Test-Path -LiteralPath $TargetFolder -PathType Container)) {
    Write-Host "指定された対象フォルダが見つかりません: $TargetFolder" -ForegroundColor Red
    return
}

if (-not $OutputDirectory) {
    $OutputDirectory = Select-Folder "検索画面.html の【保存先フォルダ】を選択してください（NAS上でもOK）"
}
if (-not $OutputDirectory) { Write-Host "保存先フォルダが選択されなかったため終了します。" -ForegroundColor Yellow; return }
if (-not (Test-Path -LiteralPath $OutputDirectory)) {
    New-Item -ItemType Directory -Path $OutputDirectory | Out-Null
}

$TargetFolder = $TargetFolder.TrimEnd('\')
if ($TargetFolder -match '^[A-Za-z]:$') { $TargetFolder += '\' }   # "D:" のままだとカレントディレクトリ扱いになるため
$rootDisplay = $TargetFolder.TrimEnd('\')

$outPath = Join-Path $OutputDirectory $OutputFileName
$utf8 = New-Object System.Text.UTF8Encoding($false)
$template = [System.IO.File]::ReadAllText($templatePath, $utf8)
$marker = "/*__DATA__*/null"
$pos = $template.IndexOf($marker)
if ($pos -lt 0) { Write-Host "テンプレートに差し込み位置($marker)がありません。" -ForegroundColor Red; return }

$excludeExt = @{}
foreach ($e in $ExcludeExtensions) { $excludeExt[$e.ToLower()] = $true }
$excludeDir = @{}
foreach ($d in $ExcludeDirNames) { $excludeDir[$d.ToLower()] = $true }

$dirList = New-Object 'System.Collections.Generic.List[string]'
$dirList.Add("")
$stack = New-Object 'System.Collections.Generic.Stack[object]'
$stack.Push(@((New-Object System.IO.DirectoryInfo($TargetFolder)), 0, ""))

$fileCount = 0
$skipped = 0
$epochTicks = 621355968000000000
$writer = New-Object System.IO.StreamWriter($outPath, $false, $utf8)

try {
    Write-Host "スキャンを開始しました。しばらくお待ちください..." -ForegroundColor Cyan
    $writer.Write($template.Substring(0, $pos))
    $scanned = Get-Date -Format "yyyy-MM-dd HH:mm"
    $writer.Write('{"root":"' + (Esc-Js $rootDisplay) + '","scanned":"' + $scanned + '","files":[')
    $writer.Write("`n")

    while ($stack.Count -gt 0) {
        $cur = $stack.Pop()
        $di = $cur[0]; $dirIdx = [int]$cur[1]; $rel = [string]$cur[2]

        try { $entries = $di.GetFileSystemInfos() }
        catch { $skipped++; continue }

        foreach ($e in $entries) {
            $isDir = ($e.Attributes -band [System.IO.FileAttributes]::Directory) -ne 0
            if ($isDir) {
                # シンボリックリンク等は無限ループ防止のため辿らない
                if (($e.Attributes -band [System.IO.FileAttributes]::ReparsePoint) -ne 0) { continue }
                if ($excludeDir.ContainsKey($e.Name.ToLower())) { continue }
                $childRel = if ($rel -eq "") { $e.Name } else { "$rel\$($e.Name)" }
                $dirList.Add($childRel)
                $stack.Push(@($e, ($dirList.Count - 1), $childRel))
            } else {
                if ($excludeExt.ContainsKey($e.Extension.ToLower())) { continue }
                $skip = $false
                foreach ($p in $ExcludeNamePrefixes) { if ($e.Name.StartsWith($p)) { $skip = $true; break } }
                if ($skip) { continue }
                $mtime = [long](($e.LastWriteTimeUtc.Ticks - $epochTicks) / 10000000)
                $writer.Write("[$dirIdx,`"$(Esc-Js $e.Name)`",$($e.Length),$mtime],`n")
                $fileCount++
                if ($fileCount % 20000 -eq 0) { Write-Host "  ... $fileCount ファイル" }
            }
        }
    }

    $writer.Write('],"dirs":[')
    $writer.Write("`n")
    foreach ($d in $dirList) { $writer.Write('"' + (Esc-Js $d) + "`",`n") }
    $writer.Write(']}')
    $writer.Write($template.Substring($pos + $marker.Length))
} catch {
    Write-Host "エラーが発生しました: $_" -ForegroundColor Red
    return
} finally {
    $writer.Close()
    $writer.Dispose()
}

$sizeMB = [math]::Round((Get-Item -LiteralPath $outPath).Length / 1MB, 1)
Write-Host ""
Write-Host "[完了] ファイル $fileCount 件・フォルダ $($dirList.Count - 1) 件を記録しました。" -ForegroundColor Green
if ($skipped -gt 0) { Write-Host "※アクセスできず読み取れなかったフォルダが $skipped 件ありました。" -ForegroundColor Yellow }
Write-Host "出力: $outPath （$sizeMB MB）" -ForegroundColor Yellow
Write-Host "この HTML をNASなど共有できる場所に置いてください。ダブルクリックで検索できます。"
