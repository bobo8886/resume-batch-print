<#
  CI 用：确认仓库里没有跟踪隐私文件 / 大二进制。

  为什么不直接写在 workflow 的 run: 里？
  因为 GitHub Actions 会把 run: 的内容写成**不带 BOM** 的 .ps1，
  而 Windows PowerShell 5.1 遇到无 BOM 的 UTF-8 会按 ANSI/GBK 解码，
  脚本里的中文会变成乱码并导致语法错误。
  所以凡是带中文的逻辑都放进仓库脚本（这些文件都带 BOM），workflow 只留 ASCII 调用。
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

$Root = Split-Path -Parent $PSScriptRoot

$forbidden = @(
    'config.json'
    '打印记录/printed.json'
    'bin/SumatraPDF.exe'
    'bin/SumatraPDF-settings.txt'
)

Push-Location $Root
try {
    # core.quotepath=false：否则中文文件名会被 git 转义成 \346\211\223... 不好比对
    $tracked = @(& git -c core.quotepath=false ls-files)
    if ($tracked.Count -eq 0) { throw 'git ls-files 没有返回任何文件，仓库状态异常' }

    $bad = @()
    foreach ($p in $forbidden) {
        if ($tracked -contains $p) { $bad += $p }
    }

    if ($bad.Count -gt 0) {
        Write-Host ('[FAIL] 以下文件不该出现在仓库里：' + ($bad -join '、')) -ForegroundColor Red
        exit 1
    }

    Write-Host ('[OK] 仓库跟踪 ' + $tracked.Count + ' 个文件，没有隐私文件或大二进制')
    exit 0
} finally {
    Pop-Location
}
