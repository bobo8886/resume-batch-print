<#
  打包发布：生成可以直接发给用户、无需联网的完整包（含 PDF 引擎）。

  用法：
      powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\build-release.ps1

  产物：
      dist\resume-batch-print-v<版本>.zip

  注意：这个包会**一并分发 SumatraPDF（GPLv3）**，所以脚本会自动把 GPLv3 许可证全文
  和第三方说明塞进压缩包，满足许可证要求。
#>
[CmdletBinding()]
param(
    [string]$Version = '1.0.0',
    [switch]$SkipLicenseFetch
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

$RepoRoot = Split-Path -Parent $PSScriptRoot
$DistDir  = Join-Path $RepoRoot 'dist'
$StageDir = Join-Path $DistDir ('stage-' + [guid]::NewGuid().ToString('N'))
$ZipPath  = Join-Path $DistDir ("resume-batch-print-v$Version.zip")

Write-Host ''
Write-Host "打包 resume-batch-print v$Version" -ForegroundColor Cyan
Write-Host ''

# ---------- 检查引擎 ----------
$engine = Join-Path $RepoRoot 'bin\SumatraPDF.exe'
if (-not (Test-Path $engine)) {
    Write-Host '本机没有 bin\SumatraPDF.exe，无法打"离线完整包"。' -ForegroundColor Yellow
    Write-Host '请先运行一次工具，在网页里点「下载 PDF 引擎」，然后再打包。' -ForegroundColor Yellow
    exit 1
}
$hash = (Get-FileHash -LiteralPath $engine -Algorithm SHA256).Hash
Write-Host ("  引擎 SHA-256: " + $hash) -ForegroundColor DarkGray

# ---------- 准备临时目录 ----------
if (Test-Path $StageDir) { Remove-Item $StageDir -Recurse -Force }
New-Item -ItemType Directory -Path $StageDir -Force | Out-Null
$pkg = Join-Path $StageDir 'resume-batch-print'
New-Item -ItemType Directory -Path $pkg -Force | Out-Null

# ---------- 复制"该有的东西" ----------
$includeFiles = @('server.ps1', '一键打印简历.bat', '使用说明.md', 'README.md',
                  'LICENSE', 'THIRD-PARTY.md', 'SECURITY.md', 'CHANGELOG.md')
foreach ($f in $includeFiles) {
    $src = Join-Path $RepoRoot $f
    if (Test-Path $src) { Copy-Item -LiteralPath $src -Destination $pkg }
    else { Write-Host ("  跳过不存在的文件: " + $f) -ForegroundColor DarkGray }
}

foreach ($d in @('web', 'bin')) {
    $src = Join-Path $RepoRoot $d
    if (Test-Path $src) { Copy-Item -LiteralPath $src -Destination $pkg -Recurse }
}
New-Item -ItemType Directory -Path (Join-Path $pkg '打印记录') -Force | Out-Null

# 明确不带这些：运行时产物 + 仓库里给开发者看的说明
foreach ($junk in @(
        (Join-Path $pkg 'config.json'),
        (Join-Path $pkg '打印记录\printed.json'),
        (Join-Path $pkg 'bin\README.md')
    )) {
    if (Test-Path -LiteralPath $junk) { Remove-Item -LiteralPath $junk -Recurse -Force -ErrorAction SilentlyContinue }
}
Get-ChildItem -LiteralPath (Join-Path $pkg '打印记录') -Force -ErrorAction SilentlyContinue |
    Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

# ---------- GPLv3 合规：把许可证全文放进去 ----------
$gpl = Join-Path $pkg 'bin\SumatraPDF-GPLv3.txt'
$gotLicense = $false
if (-not $SkipLicenseFetch) {
    $old = $ProgressPreference
    try {
        $ProgressPreference = 'SilentlyContinue'
        [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -Uri 'https://www.gnu.org/licenses/gpl-3.0.txt' -OutFile $gpl -UseBasicParsing -TimeoutSec 60
        $gotLicense = $true
        Write-Host '  已附带 GPLv3 许可证全文' -ForegroundColor Green
    } catch {
        Write-Host ('  下载 GPLv3 全文失败：' + $_.Exception.Message) -ForegroundColor Yellow
    } finally { $ProgressPreference = $old }
}
if (-not $gotLicense) {
    [System.IO.File]::WriteAllText($gpl,
        "This package bundles SumatraPDF, which is licensed under GNU GPL v3.`r`n" +
        "Full license text: https://www.gnu.org/licenses/gpl-3.0.txt`r`n" +
        "Source code: https://github.com/sumatrapdfreader/sumatrapdf`r`n",
        (New-Object System.Text.UTF8Encoding($false)))
    Write-Host '  已写入 GPLv3 指路文件（未能下载全文）' -ForegroundColor Yellow
}

# ---------- 附一张"从哪来"的说明 ----------
$notice = @"
resume-batch-print / 简历一键打印  v$Version
=====================================================

这是含 PDF 引擎的离线完整包，解压后双击「一键打印简历.bat」即可使用，无需联网。

本包内的第三方组件
------------------
bin\SumatraPDF.exe        SumatraPDF 3.5.2 (64 位)
                          作者: Krzysztof Kowalczyk 及贡献者
                          许可证: GNU GPL v3
                          官网: https://www.sumatrapdfreader.org/
                          源码: https://github.com/sumatrapdfreader/sumatrapdf
                          许可证全文见 bin\SumatraPDF-GPLv3.txt

                          SHA-256:
                          $hash

                          本工具以独立进程 + 命令行参数的方式调用它，
                          属于 GPL 所称的 "mere aggregation"，不影响本项目自身的 MIT 许可。

本项目源码: MIT 许可证，详见 LICENSE
安全说明:   详见 SECURITY.md
"@
[System.IO.File]::WriteAllText((Join-Path $pkg '本包说明.txt'), $notice, (New-Object System.Text.UTF8Encoding($true)))

# ---------- 压缩 ----------
if (-not (Test-Path $DistDir)) { New-Item -ItemType Directory -Path $DistDir -Force | Out-Null }
if (Test-Path $ZipPath) { Remove-Item $ZipPath -Force }
Compress-Archive -Path $pkg -DestinationPath $ZipPath -CompressionLevel Optimal

Remove-Item $StageDir -Recurse -Force -ErrorAction SilentlyContinue

$size = [math]::Round((Get-Item $ZipPath).Length / 1MB, 1)
Write-Host ''
Write-Host ("完成: " + $ZipPath + "  (" + $size + " MB)") -ForegroundColor Green
Write-Host ''
Write-Host '下一步：' -ForegroundColor DarkGray
Write-Host '  1. 到 GitHub 仓库页面 → Releases → Draft a new release' -ForegroundColor DarkGray
Write-Host ("  2. Tag 填 v$Version，把上面这个 zip 拖进去当附件") -ForegroundColor DarkGray
Write-Host '  3. README 里的下载链接指向该 Release 即可' -ForegroundColor DarkGray
Write-Host ''
