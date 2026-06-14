# ====== 批量独立 Chrome 登录窗口生成器 ======
# 使用方法：填写同目录的“学生电话表.xlsx”，然后双击“双击生成登录窗口.bat”。
# Excel 表头只需要两列：姓名、手机。

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$excelPath = Join-Path $scriptDir "学生电话表.xlsx"
$root = Join-Path $scriptDir "生成结果_登录窗口"
$url = "https://icc.cffpd.org.cn/activity?competitionType=theme"
$groupSize = 5

function Convert-ExcelValueToText($value) {
    if ($null -eq $value) { return "" }
    if ($value -is [double] -or $value -is [single] -or $value -is [decimal]) {
        return ([decimal]$value).ToString("0")
    }
    return ([string]$value).Trim()
}

function Release-ComObject($obj) {
    if ($null -ne $obj) {
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj)
    }
}

function Write-BatFile([string]$path, [string[]]$lines) {
    $content = ($lines -join "`r`n") + "`r`n"
    [System.IO.File]::WriteAllText($path, $content, [System.Text.Encoding]::Default)
}

if (!(Test-Path -LiteralPath $excelPath)) {
    throw "找不到学生电话表：$excelPath"
}

$students = @()
$excel = $null
$workbook = $null
$worksheet = $null
$usedRange = $null

try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $workbook = $excel.Workbooks.Open($excelPath)
    $worksheet = $workbook.Worksheets.Item(1)
    $usedRange = $worksheet.UsedRange
    $values = $usedRange.Value2

    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count
    if ($rowCount -lt 2) { throw "学生电话表没有可读取的数据行。" }

    $nameCol = $null
    $phoneCol = $null
    for ($col = 1; $col -le $colCount; $col++) {
        $header = Convert-ExcelValueToText $values[1, $col]
        if ($header -eq "姓名") { $nameCol = $col }
        if ($header -eq "手机" -or $header -eq "手机号" -or $header -eq "电话" -or $header -eq "电话号码") { $phoneCol = $col }
    }
    if ($null -eq $nameCol -or $null -eq $phoneCol) {
        throw "表头必须包含[姓名]和[手机]两列。"
    }

    for ($row = 2; $row -le $rowCount; $row++) {
        $name = Convert-ExcelValueToText $values[$row, $nameCol]
        $phone = Convert-ExcelValueToText $values[$row, $phoneCol]
        if ($name -and $phone) {
            $students += [pscustomobject]@{
                Index = $students.Count + 1
                Name = $name
                Phone = $phone
            }
        }
    }
}
finally {
    if ($workbook) { $workbook.Close($false) | Out-Null }
    if ($excel) { $excel.Quit() | Out-Null }
    Release-ComObject $usedRange
    Release-ComObject $worksheet
    Release-ComObject $workbook
    Release-ComObject $excel
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}

if ($students.Count -eq 0) {
    throw "没有从学生电话表读取到姓名和手机。"
}

$dataDir = Join-Path $root "data"
$labelDir = Join-Path $root "labels"
$launcherPath = Join-Path $root "launch_account.ps1"
$enc = New-Object System.Text.UTF8Encoding($false)

New-Item -ItemType Directory -Force -Path $root, $dataDir, $labelDir | Out-Null
Get-ChildItem -LiteralPath $root -Filter "*.bat" -File -ErrorAction SilentlyContinue | Remove-Item -Force
Get-ChildItem -LiteralPath $labelDir -Filter "*.html" -File -ErrorAction SilentlyContinue | Remove-Item -Force

$launcher = @"
param(
    [Parameter(Mandatory=`$true)][string]`$Account
)

`$ErrorActionPreference = "Stop"
`$root = Split-Path -Parent `$MyInvocation.MyCommand.Path
`$chromeCandidates = @(
    "`$env:ProgramFiles\Google\Chrome\Application\chrome.exe",
    "`${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
    "`$env:LocalAppData\Google\Chrome\Application\chrome.exe"
)
`$chrome = `$chromeCandidates | Where-Object { `$_ -and (Test-Path -LiteralPath `$_) } | Select-Object -First 1
if (-not `$chrome) { `$chrome = "chrome.exe" }

`$dataDir = Join-Path `$root "data\`$Account"
`$labelPath = Join-Path `$root "labels\`$Account.html"
`$url = "$url"

New-Item -ItemType Directory -Force -Path `$dataDir | Out-Null

& `$chrome "--user-data-dir=`$dataDir" "--no-first-run" "--no-default-browser-check" "--new-window" `$url

Start-Sleep -Milliseconds 800

if (Test-Path -LiteralPath `$labelPath) {
    `$labelUri = (New-Object System.Uri(`$labelPath)).AbsoluteUri
    & `$chrome "--user-data-dir=`$dataDir" `$labelUri
}
"@
[System.IO.File]::WriteAllText($launcherPath, $launcher, (New-Object System.Text.UTF8Encoding($true)))

$mapLines = New-Object System.Collections.Generic.List[string]
$mapLines.Add("序号`t姓名`t手机`t数据目录`t启动器")

foreach ($student in $students) {
    $nn = "{0:D2}" -f $student.Index
    $acc = "acc$nn"
    $accountDataDir = Join-Path $dataDir $acc
    $labelPath = Join-Path $labelDir "$acc.html"
    $batPath = Join-Path $root "$acc.bat"

    New-Item -ItemType Directory -Force -Path $accountDataDir | Out-Null

    $html = @"
<!doctype html>
<html>
<head>
  <meta charset="utf-8">
  <title>$nn $($student.Name) $($student.Phone)</title>
</head>
<body style="font-family:'Microsoft YaHei',Arial,sans-serif;text-align:center;margin-top:11vh;background:#f8fbff;color:#14213d">
  <div style="font-size:92px;color:#1a73e8;font-weight:700">$nn</div>
  <div style="font-size:50px;font-weight:700;margin-top:12px">$($student.Name)</div>
  <div style="font-size:34px;margin-top:18px">手机号：$($student.Phone)</div>
  <div style="font-size:22px;margin-top:34px;color:#5f6b7a">批量登录窗口</div>
</body>
</html>
"@
    [System.IO.File]::WriteAllText($labelPath, $html, $enc)

    Write-BatFile $batPath @(
        "@echo off",
        "powershell -NoProfile -ExecutionPolicy Bypass -File ""%~dp0launch_account.ps1"" -Account ""$acc"""
    )
    $mapLines.Add("$nn`t$($student.Name)`t$($student.Phone)`t$accountDataDir`t$batPath")
}

$groupCount = [Math]::Ceiling($students.Count / $groupSize)
for ($group = 1; $group -le $groupCount; $group++) {
    $start = (($group - 1) * $groupSize) + 1
    $end = [Math]::Min($group * $groupSize, $students.Count)
    $groupBatPath = Join-Path $root ("group{0}_no{1:D2}-{2:D2}.bat" -f $group, $start, $end)
    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add("@echo off")
    $lines.Add("cd /d ""%~dp0""")
    for ($i = $start; $i -le $end; $i++) {
        $nn = "{0:D2}" -f $i
        $lines.Add("call ""acc$nn.bat""")
        $lines.Add("timeout /t 1 /nobreak >nul")
    }
    Write-BatFile $groupBatPath $lines.ToArray()
}

$readme = @"
========================================
  批量独立 Chrome 登录窗口 · 使用说明
========================================

生成时间：$(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
目标网址：$url
账号数量：$($students.Count)

【怎么用】
1. 修改“学生电话表.xlsx”，只保留或填写两列：姓名、手机。
2. 双击上一级文件夹里的“双击生成登录窗口.bat”。
3. 打开本文件夹：$root
4. 双击 group1_no01-05.bat、group2_no06-10.bat 这类分组启动器。

【说明】
- 每个 accNN 都是一个独立 Chrome 登录环境，登录状态保存在 data\accNN。
- data 文件夹不要删；删掉后对应账号会变成未登录状态。
- labels 文件夹是姓名和手机号标牌页。
- 对照表.txt 可以核对序号、姓名、手机和数据目录。
"@
[System.IO.File]::WriteAllText((Join-Path $root "使用说明.txt"), $readme, $enc)
[System.IO.File]::WriteAllLines((Join-Path $root "对照表.txt"), $mapLines, $enc)

Write-Host ""
Write-Host "完成：已生成 $($students.Count) 个登录窗口。" -ForegroundColor Green
Write-Host "输出位置：$root"
Write-Host "建议使用：group1_no01-05.bat、group2_no06-10.bat... 分组启动器。"
