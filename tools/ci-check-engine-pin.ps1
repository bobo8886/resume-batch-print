<#
  校验"打包脚本会不会接受一个哈希不符的引擎"。

  背景：tools\build-release.ps1 以前只要 bin\SumatraPDF.exe **存在**就把它打进
  "离线完整包"，只把哈希打印出来写进包内说明，**从不与 server.ps1 里固定的值比对**。
  于是任何一个放在 bin\ 下的可执行文件都会被当成官方 3.5.2 分发出去。

  这个脚本在一个临时副本里放一个**长度相同但内容不同**的假引擎，
  断言 build-release 会拒绝打包；再断言逃生舱 -AllowUnpinnedEngine 仍然可用。

  用法：powershell -NoProfile -ExecutionPolicy Bypass -File .\tools\ci-check-engine-pin.ps1
#>
$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

$RepoRoot = Split-Path -Parent $PSScriptRoot
$fail = 0
function Check([string]$name, [bool]$ok, [string]$detail = '') {
    if ($ok) { Write-Host "  [PASS] $name" -ForegroundColor Green }
    else { Write-Host "  [FAIL] $name  $detail" -ForegroundColor Red; $script:fail++ }
}

Write-Host ''
Write-Host '── 校验：打包脚本必须拒绝未固定的引擎 ─────────────────────────' -ForegroundColor Cyan

# ---- 1. 固定值本身要读得到 ----
$srcText = Get-Content -LiteralPath (Join-Path $RepoRoot 'server.ps1') -Raw -Encoding UTF8
$m = [regex]::Match($srcText, "EngineSha256\s*=\s*'([0-9A-Fa-f]{64})'")
Check 'server.ps1 里有 64 位 SHA-256 固定值' $m.Success
if (-not $m.Success) { Write-Host '  无法继续'; exit 1 }
$pinned = $m.Groups[1].Value.ToUpperInvariant()

# ---- 2. 当前引擎必须与固定值一致 ----
$realEngine = Join-Path $RepoRoot 'bin\SumatraPDF.exe'
Check '本机 bin\SumatraPDF.exe 存在' (Test-Path -LiteralPath $realEngine)
if (-not (Test-Path -LiteralPath $realEngine)) { Write-Host '  无法继续'; exit 1 }
$realHash = (Get-FileHash -LiteralPath $realEngine -Algorithm SHA256).Hash.ToUpperInvariant()
Check '本机引擎哈希与固定值一致' ($realHash -eq $pinned) "实际 $realHash / 固定 $pinned"
$realLen = (Get-Item -LiteralPath $realEngine).Length
$realBytes = [System.IO.File]::ReadAllBytes($realEngine)

# ---- 3. 造一个临时副本，把假引擎放进去 ----
$tmp = Join-Path $env:TEMP ('rbp-pincheck-' + [guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Path (Join-Path $tmp 'tools') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tmp 'bin') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $tmp 'web') -Force | Out-Null
    Copy-Item -LiteralPath (Join-Path $RepoRoot 'server.ps1') -Destination $tmp
    Copy-Item -LiteralPath (Join-Path $RepoRoot 'tools\build-release.ps1') -Destination (Join-Path $tmp 'tools')
    Copy-Item -Path (Join-Path $RepoRoot 'web\*') -Destination (Join-Path $tmp 'web')

    # 假引擎：与真引擎**长度完全相同**，但内容不同。
    # 长度相同是攻击者可满足的，所以这条对"按长度复用/按存在性打包"都是有效的反例。
    $fake = New-Object byte[] $realLen
    [Array]::Copy($realBytes, $fake, $realLen)
    $fake[$realLen - 1] = [byte](($fake[$realLen - 1] + 1) % 256)
    [System.IO.File]::WriteAllBytes((Join-Path $tmp 'bin\SumatraPDF.exe'), $fake)
    $fakeHash = (Get-FileHash -LiteralPath (Join-Path $tmp 'bin\SumatraPDF.exe') -Algorithm SHA256).Hash.ToUpperInvariant()
    Check '假引擎长度与真引擎相同但哈希不同' (($fake.Length -eq $realLen) -and ($fakeHash -ne $pinned))

    # ---- 4. 默认必须拒绝 ----
    $log = Join-Path $tmp 'out.log'
    $p = Start-Process -FilePath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" `
        -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $tmp 'tools\build-release.ps1'), '-SkipLicenseFetch' `
        -Wait -PassThru -NoNewWindow -RedirectStandardOutput $log -RedirectStandardError (Join-Path $tmp 'err.log')
    Check '哈希不符时打包脚本拒绝（退出码非 0）' ($p.ExitCode -ne 0) "退出码 $($p.ExitCode)"
    $zips = @(Get-ChildItem -Path (Join-Path $tmp 'dist') -Filter '*.zip' -ErrorAction SilentlyContinue)
    Check '哈希不符时没有产出 zip' ($zips.Count -eq 0) "竟然产出了 $($zips.Count) 个"

    # ---- 5. 逃生舱仍然可用 ----
    $p2 = Start-Process -FilePath "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe" `
        -ArgumentList '-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $tmp 'tools\build-release.ps1'), '-SkipLicenseFetch', '-AllowUnpinnedEngine' `
        -Wait -PassThru -NoNewWindow -RedirectStandardOutput (Join-Path $tmp 'out2.log') -RedirectStandardError (Join-Path $tmp 'err2.log')
    $zips2 = @(Get-ChildItem -Path (Join-Path $tmp 'dist') -Filter '*.zip' -ErrorAction SilentlyContinue)
    Check '显式 -AllowUnpinnedEngine 时可以继续打包' (($p2.ExitCode -eq 0) -and ($zips2.Count -ge 1)) "退出码 $($p2.ExitCode), zip $($zips2.Count)"
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

Write-Host ''
if ($fail -eq 0) { Write-Host '  引擎固定校验全部通过' -ForegroundColor Green; exit 0 }
Write-Host "  失败 $fail 项" -ForegroundColor Red
exit 1
