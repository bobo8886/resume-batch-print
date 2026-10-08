<#
  CI / 本地用：下载固定版本的 SumatraPDF 打印引擎并校验 SHA-256。

  两个作用：
    1. 让 CI 能真正覆盖 PDF 打印链路（裸 runner 上没有引擎）
    2. 顺带验证 server.ps1 里固定的下载 URL 与哈希**现在依然有效**
       —— 上游如果换了文件，这里会立刻失败，而不是等用户下载时才发现

  用法：
      powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\ci-fetch-engine.ps1
#>
[CmdletBinding()]
param(
    [string]$Destination = ''
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

$Root = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($Destination)) { $Destination = Join-Path $Root 'bin\SumatraPDF.exe' }

# 从 server.ps1 里读取"唯一事实来源"，避免两处硬编码走偏
$srv  = Get-Content (Join-Path $Root 'server.ps1') -Raw -Encoding UTF8
$url  = [regex]::Match($srv, "\`$script:EngineUrl\s*=\s*'([^']+)'").Groups[1].Value
$hash = [regex]::Match($srv, "\`$script:EngineSha256\s*=\s*'([0-9A-Fa-f]{64})'").Groups[1].Value
if (-not $url -or -not $hash) { throw '没能从 server.ps1 里解析出引擎 URL / SHA-256' }

Write-Host "URL : $url"
Write-Host "SHA : $hash"

$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('rbp-engine-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
$oldProgress = $ProgressPreference
try {
    $ProgressPreference = 'SilentlyContinue'
    [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12

    $zip = Join-Path $tmp 'engine.zip'
    Invoke-WebRequest -Uri $url -OutFile $zip -UseBasicParsing -TimeoutSec 300
    Expand-Archive -LiteralPath $zip -DestinationPath $tmp -Force

    $exe = @(Get-ChildItem -LiteralPath $tmp -Filter '*.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1)
    if ($exe.Count -eq 0) { throw '下载包里没有找到可执行文件' }

    $actual = (Get-FileHash -LiteralPath $exe[0].FullName -Algorithm SHA256).Hash
    if ($actual -ne $hash) {
        throw ("[FAIL] 引擎哈希不匹配！期望 " + $hash + "，实际 " + $actual + " —— 上游文件可能已变更，请核实后再更新 server.ps1")
    }
    Write-Host '[OK] 引擎下载成功，SHA-256 校验通过'

    $dir = Split-Path -Parent $Destination
    if (-not (Test-Path $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    Copy-Item -LiteralPath $exe[0].FullName -Destination $Destination -Force
    Write-Host ('     已放置到: ' + $Destination)
} finally {
    $ProgressPreference = $oldProgress
    Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
}
