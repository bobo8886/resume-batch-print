<#
  resume-batch-print 自动化测试
  零依赖（不需要 Pester / PSScriptAnalyzer），在 Windows PowerShell 5.1 下直接跑：

      powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\run-tests.ps1

  它会：
    1) 静态检查：BOM、语法、隐私信息、引擎哈希、.gitignore 覆盖
    2) 把仓库复制到临时目录（模拟"用户 clone 下来"的干净副本）
    3) 在干净副本上起服务，跑 API 端到端 + 安全用例
    4) 全程只做"试运行"，不会真的打印任何东西
#>
[CmdletBinding()]
param(
    [switch]$KeepTemp
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }

$RepoRoot = Split-Path -Parent $PSScriptRoot
$script:Pass = 0
$script:Fail = 0
$script:Failures = @()

function Section([string]$t) {
    Write-Host ''
    Write-Host ('── ' + $t + ' ' + ('─' * [Math]::Max(0, 56 - $t.Length))) -ForegroundColor Cyan
}
function Ok([string]$name) {
    $script:Pass++
    Write-Host ('  [PASS] ' + $name) -ForegroundColor Green
}
function Bad([string]$name, [string]$why) {
    $script:Fail++
    $script:Failures += ($name + '  ::  ' + $why)
    Write-Host ('  [FAIL] ' + $name) -ForegroundColor Red
    Write-Host ('         ' + $why) -ForegroundColor DarkGray
}
function Check([string]$name, [bool]$cond, [string]$why = '条件不成立') {
    if ($cond) { Ok $name } else { Bad $name $why }
}
function Check-Throws([string]$name, [scriptblock]$body) {
    try { & $body | Out-Null; Bad $name '期望抛异常但没有' }
    catch { Ok $name }
}

# ============================================================
Section '静态检查：文件编码与语法'
# ============================================================

$ps1Files = @(Get-ChildItem -LiteralPath $RepoRoot -Recurse -File |
              Where-Object {
                  $_.Extension -in @('.ps1', '.psm1', '.psd1') -and
                  $_.FullName -notmatch '\\\.git\\'
              })

Check '仓库里能找到 .ps1 文件' ($ps1Files.Count -gt 0) ('找到 ' + $ps1Files.Count + ' 个')

foreach ($f in $ps1Files) {
    $rel = $f.FullName.Substring($RepoRoot.Length).TrimStart('\')
    if ($f.Extension -eq '.psd1') {
        # 数据文件只检查编码，不做脚本语法解析
        $b = [System.IO.File]::ReadAllBytes($f.FullName)
        $hasBom = ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF)
        Check ("UTF-8 BOM: $rel") $hasBom '缺少 BOM，中文会变乱码'
        $okData = $true
        try { [void](Import-PowerShellDataFile -LiteralPath $f.FullName -ErrorAction Stop) } catch { $okData = $false }
        Check ("数据文件可解析: $rel") $okData 'Import-PowerShellDataFile 失败'
        continue
    }

    # UTF-8 BOM 检查（Windows PowerShell 5.1 必须有 BOM 才正确读中文）
    $b = [System.IO.File]::ReadAllBytes($f.FullName)
    $hasBom = ($b.Length -ge 3 -and $b[0] -eq 0xEF -and $b[1] -eq 0xBB -and $b[2] -eq 0xBF)
    Check ("UTF-8 BOM: $rel") $hasBom '缺少 BOM，中文会变乱码'

    # 语法检查
    $err = $null
    $text = [System.IO.File]::ReadAllText($f.FullName, (New-Object System.Text.UTF8Encoding($true)))
    [void][System.Management.Automation.PSParser]::Tokenize($text, [ref]$err)
    Check ("语法正确: $rel") ($err.Count -eq 0) (($err | ForEach-Object { "行$($_.Token.StartLine):$($_.Message)" }) -join '; ')
}

# ============================================================
Section '静态检查：workflow 的 run: 块必须是纯 ASCII'
# ============================================================
# 这个坑踩过一次：GitHub Actions 会把 `run:` 的内容写成**不带 BOM** 的 .ps1，
# Windows PowerShell 5.1 遇到无 BOM 的 UTF-8 会按 ANSI/GBK 解码，
# 中文直接变乱码并语法报错（CI 上表现为 "Unexpected token ..."）。
# 所以 run: 里只能写 ASCII；带中文的逻辑要放进仓库脚本（那些文件带 BOM）。

$wfDir = Join-Path $RepoRoot '.github\workflows'
$wfFiles = @(Get-ChildItem -LiteralPath $wfDir -Filter '*.yml' -File -ErrorAction SilentlyContinue)
Check '能找到 workflow 文件' ($wfFiles.Count -gt 0) ('目录: ' + $wfDir)

foreach ($wf in $wfFiles) {
    $lines = @(Get-Content -LiteralPath $wf.FullName -Encoding UTF8)
    $bad = @()
    $inRunBlock = $false
    $runIndent = 0
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $ln = $lines[$i]
        if (-not $inRunBlock) {
            if ($ln -match '^(\s*)run:\s*[|>]\s*$') {
                $inRunBlock = $true
                $runIndent = $Matches[1].Length
                continue
            }
            if ($ln -match '^\s*run:\s*(.+)$') {
                if ($Matches[1] -match '[^\x00-\x7F]') { $bad += ('行' + ($i + 1) + ': ' + $Matches[1].Trim()) }
            }
            continue
        }
        if ([string]::IsNullOrWhiteSpace($ln)) { continue }
        $ind = $ln.Length - $ln.TrimStart().Length
        if ($ind -le $runIndent) { $inRunBlock = $false; $i--; continue }
        if ($ln -match '[^\x00-\x7F]') { $bad += ('行' + ($i + 1) + ': ' + $ln.Trim()) }
    }
    Check ("run: 块纯 ASCII: " + $wf.Name) ($bad.Count -eq 0) (($bad | Select-Object -First 5) -join '  |  ')
}

# ============================================================
Section '静态检查：隐私与密钥'
# ============================================================

$scanFiles = @(Get-ChildItem -LiteralPath $RepoRoot -Recurse -File |
               Where-Object {
                   $_.FullName -notmatch '\\\.git\\' -and
                   $_.FullName -notmatch '\\tests\\' -and
                   $_.FullName -notmatch '\\docs\\' -and
                   $_.Extension -in @('.ps1', '.md', '.js', '.css', '.html', '.bat', '.json', '.yml', '.yaml')
               })

Check '有可扫描的文本文件' ($scanFiles.Count -gt 0) ('共 ' + $scanFiles.Count + ' 个')

$piiPatterns = [ordered]@{
    'Windows 用户桌面路径' = '[A-Za-z]:\\Users\\[^\\\s"'']+\\Desktop'
    'Windows 用户目录路径' = '[A-Za-z]:\\Users\\[A-Za-z0-9_.-]+\\'
    '私钥内容'             = '-----BEGIN [A-Z ]*PRIVATE KEY-----'
    '疑似硬编码密钥'       = '(?i)(password|passwd|secret|api[_-]?key|access[_-]?token)\s*[:=]\s*["''][^"'']{8,}'
    '内网 IP'              = '\b(192\.168\.\d+\.\d+|10\.\d+\.\d+\.\d+)\b'
}
foreach ($kv in $piiPatterns.GetEnumerator()) {
    $hits = @()
    foreach ($f in $scanFiles) {
        $t = Get-Content -LiteralPath $f.FullName -Raw -Encoding UTF8 -ErrorAction SilentlyContinue
        if ($t -and $t -match $kv.Value) {
            $hits += $f.FullName.Substring($RepoRoot.Length).TrimStart('\')
        }
    }
    Check ('不含' + $kv.Key) ($hits.Count -eq 0) ('出现在: ' + ($hits -join ', '))
}

# ============================================================
Section '静态检查：.gitignore 覆盖'
# ============================================================

$ignore = Get-Content -LiteralPath (Join-Path $RepoRoot '.gitignore') -Raw -Encoding UTF8
foreach ($must in @('config.json', '打印记录', 'SumatraPDF.exe')) {
    Check (".gitignore 忽略 $must") ($ignore -match [regex]::Escape($must)) '没有被排除，可能把隐私/大文件提交上去'
}

# ============================================================
Section '静态检查：PDF 引擎哈希'
# ============================================================

$srvText = Get-Content -LiteralPath (Join-Path $RepoRoot 'server.ps1') -Raw -Encoding UTF8
$m = [regex]::Match($srvText, "\`$script:EngineSha256\s*=\s*'([0-9A-Fa-f]{64})'")
Check 'server.ps1 里固定了 64 位 SHA-256' $m.Success '没找到或格式不对'
if ($m.Success) {
    $pinned = $m.Groups[1].Value.ToUpperInvariant()
    $localExe = Join-Path $RepoRoot 'bin\SumatraPDF.exe'
    if (Test-Path $localExe) {
        $actual = (Get-FileHash -LiteralPath $localExe -Algorithm SHA256).Hash
        Check '本地引擎哈希与代码里固定的一致' ($actual -eq $pinned) ("代码=$pinned 实际=$actual")
    } else {
        Write-Host '  [SKIP] 本地没有 bin\SumatraPDF.exe（仓库本就不该有）' -ForegroundColor DarkGray
    }
}

# ============================================================
Section '干净副本：模拟用户 clone 下来的状态'
# ============================================================

$tmp = Join-Path $env:TEMP ('rbp-test-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null

# 只复制"仓库里应该有的东西"：server.ps1 + web + bin/README.md + 文档
Copy-Item -LiteralPath (Join-Path $RepoRoot 'server.ps1') -Destination $tmp
Copy-Item -LiteralPath (Join-Path $RepoRoot 'web') -Destination $tmp -Recurse
New-Item -ItemType Directory -Path (Join-Path $tmp 'bin') -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $tmp '打印记录') -Force | Out-Null

# 有引擎就一起复制，没有就测"缺引擎"的状态
$localExe = Join-Path $RepoRoot 'bin\SumatraPDF.exe'
$hasEngine = Test-Path $localExe
if ($hasEngine) {
    Copy-Item -LiteralPath $localExe -Destination (Join-Path $tmp 'bin\SumatraPDF.exe')
}

Check '干净副本里没有 config.json' (-not (Test-Path (Join-Path $tmp 'config.json')))
Check '干净副本里没有打印记录' (-not (Test-Path (Join-Path $tmp '打印记录\printed.json')))

# 造一个测试用的"简历"文件夹（内容不重要，测试只用 dryRun）
$resume = Join-Path $tmp '_resumes'
New-Item -ItemType Directory -Path $resume -Force | Out-Null
'a' | Out-File (Join-Path $resume 'a-张三-产品经理.pdf') -Encoding utf8
'b' | Out-File (Join-Path $resume 'b-李四-前端工程师.docx') -Encoding utf8
'c' | Out-File (Join-Path $resume 'c-不是简历.txt.exe') -Encoding utf8
'd' | Out-File (Join-Path $resume 'd-忽略我.xlsx') -Encoding utf8

# ============================================================
Section '启动服务'
# ============================================================

$listener = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, 0)
$listener.Start()
$Port = $listener.LocalEndpoint.Port
$listener.Stop()

$psExe = Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
$proc = Start-Process -FilePath $psExe -PassThru -WindowStyle Hidden -ArgumentList @(
    '-NoProfile', '-ExecutionPolicy', 'Bypass',
    '-File', (Join-Path $tmp 'server.ps1'),
    '-NoBrowser', '-Port', "$Port"
)

$Base = "http://127.0.0.1:$Port"
$up = $false
for ($i = 0; $i -lt 40; $i++) {
    Start-Sleep -Milliseconds 500
    try {
        $t = New-Object System.Net.Sockets.TcpClient
        $t.Connect('127.0.0.1', $Port); $t.Close(); $up = $true; break
    } catch { }
}
Check '服务已启动并监听' $up ("端口 $Port 无法连接")

function Api([string]$method, [string]$path, $body = $null) {
    $u = $Base + $path
    if ($null -eq $body) { return Invoke-RestMethod -Uri $u -Method $method -TimeoutSec 30 }
    return Invoke-RestMethod -Uri $u -Method $method -TimeoutSec 120 -ContentType 'application/json; charset=utf-8' `
        -Body ([System.Text.Encoding]::UTF8.GetBytes(($body | ConvertTo-Json -Compress -Depth 5)))
}

# 直接发原始 TCP，用来构造"恶意"请求
function RawRequest([hashtable]$headers, [string]$method, [string]$path, [string]$body = '') {
    $c = New-Object System.Net.Sockets.TcpClient
    $c.Connect('127.0.0.1', $Port)
    try {
        $s = $c.GetStream()
        $sb = New-Object System.Text.StringBuilder
        [void]$sb.Append("$method $path HTTP/1.1`r`n")
        foreach ($k in $headers.Keys) { [void]$sb.Append("$k`: $($headers[$k])`r`n") }
        if ($body) { [void]$sb.Append("Content-Length: $([System.Text.Encoding]::UTF8.GetByteCount($body))`r`n") }
        [void]$sb.Append("`r`n")
        if ($body) { [void]$sb.Append($body) }
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($sb.ToString())
        $s.Write($bytes, 0, $bytes.Length); $s.Flush()
        $sr = New-Object System.IO.StreamReader($s, [System.Text.Encoding]::UTF8)
        return $sr.ReadToEnd()
    } finally { $c.Close() }
}
function StatusOf([string]$resp) { if ($resp -match '^HTTP/1\.1 (\d{3})') { return [int]$Matches[1] } return 0 }

try {
    # --------------------------------------------------------
    Section 'API 端到端'
    # --------------------------------------------------------

    $info = Api 'GET' '/api/info'
    Check '/api/info 返回打印机列表' ($null -ne $info.printers)
    Check '/api/info 返回引擎状态' ($null -ne $info.engines)
    Check '/api/info 返回 engineMissing 标志' ($info.PSObject.Properties.Name -contains 'engineMissing')
    Check '/api/info 返回打印浓度设置' ($info.PSObject.Properties.Name -contains 'enhance')
    if ($hasEngine) {
        Check '引擎被正确识别为可用' ($info.engines.sumatra -eq $true) '有引擎但报告不可用'
        Check 'engineMissing = false' ($info.engineMissing -eq $false)
    } else {
        Check '缺引擎时 engineMissing = true' ($info.engineMissing -eq $true)
    }

    $st = Api 'POST' '/api/selftest' @{}
    Check '/api/selftest 返回引擎报告' ($null -ne $st.engines)

    # 自检之后引擎状态才会变成真实结果，这里必须重新取一次 info，
    # 否则拿到的是"还没检测"的旧值（全 false）。
    $info = Api 'GET' '/api/info'
    if ($hasEngine) {
        Check '自检后引擎被标记为可用' ($info.engines.sumatra -eq $true) '有引擎但报告不可用'
    }
    Check '自检后 tested = true' ($info.engines.tested -eq $true) 'selftest 没有回写状态'

    $dir = Api 'POST' '/api/list-dir' @{ path = '' }
    Check '/api/list-dir 根目录返回驱动器' ((@($dir.drives)).Count -gt 0)
    Check '/api/list-dir 返回快捷入口' ((@($dir.shortcuts)).Count -gt 0)

    $dir2 = Api 'POST' '/api/list-dir' @{ path = $tmp }
    Check '/api/list-dir 能列出指定目录' ($dir2.ok -eq $true -and $dir2.path -eq $tmp)

    $dirBad = Api 'POST' '/api/list-dir' @{ path = 'Z:\绝对不存在的目录' }
    Check '/api/list-dir 对不存在目录返回错误' ($dirBad.ok -eq $false -and $dirBad.error)

    $scan = Api 'POST' '/api/scan' @{ folder = $resume; recursive = $true }
    Check '/api/scan 只认支持的格式' ((@($scan.files)).Count -eq 2) ('实际 ' + (@($scan.files)).Count + ' 个（应为 2）')
    Check '/api/scan 统计被忽略的文件' ($scan.otherCount -eq 2) ('otherCount=' + $scan.otherCount)
    Check '/api/scan 默认标记为未打印' (@($scan.files | Where-Object { $_.printed }).Count -eq 0)

    $pdf = @($scan.files | Where-Object { $_.ext -eq '.pdf' })[0]
    $docx = @($scan.files | Where-Object { $_.ext -eq '.docx' })[0]

    # 这台机器到底有什么引擎？（CI 的裸 runner 上可能一个都没有，那就要断言"明确报错"）
    $anyPdf = ($info.engines.sumatra -eq $true)
    $anyOffice = (($info.engines.word -eq $true) -or ($info.engines.wps -eq $true))
    $anyEngine = ($anyPdf -or $anyOffice)
    Write-Host ("        可用引擎: PDF=" + $anyPdf + " Office=" + $anyOffice + "  (加深通道=" + $info.enhanceAvailable + ")") -ForegroundColor DarkGray

    $dry = Api 'POST' '/api/print-one' @{ path = $pdf.path; printer = ''; copies = 1; dryRun = $true; enhance = 'auto' }
    if ($anyEngine) {
        Check '试运行 PDF 返回成功' ($dry.ok -eq $true) $dry.detail
        Check '试运行不会真的打印' ($dry.detail -match '试运行')
        Check '试运行给出了打印方式' (-not [string]::IsNullOrWhiteSpace($dry.method)) 'method 为空'
    } else {
        Check '无任何引擎时 PDF 明确报错' ($dry.ok -eq $false -and $dry.detail -match '没有可用的打印引擎') $dry.detail
    }

    $dry2 = Api 'POST' '/api/print-one' @{ path = $docx.path; printer = ''; copies = 3; dryRun = $true; enhance = 'auto' }
    if ($anyOffice) {
        Check '试运行 Word 文档返回成功' ($dry2.ok -eq $true) $dry2.detail
    } else {
        Check '无 Word/WPS 时 Word 文档明确报错' ($dry2.ok -eq $false) $dry2.detail
    }

    # 「加深」通道最终仍要靠 SumatraPDF 出纸，没有它就必须报错，不能给假阳性
    foreach ($mode in @('auto', 'normal', 'dark', 'darker')) {
        $r = Api 'POST' '/api/print-one' @{ path = $pdf.path; printer = ''; copies = 1; dryRun = $true; enhance = $mode }
        if ($anyEngine) {
            Check "打印浓度 $mode 可用" ($r.ok -eq $true) $r.detail
        } else {
            Check "无引擎时打印浓度 $mode 明确报错" ($r.ok -eq $false) $r.detail
        }
    }
    if ($anyPdf -or $anyOffice) {
        # 有引擎时，试运行的结果必须是"真的能打"的那条路径
        $r = Api 'POST' '/api/print-one' @{ path = $pdf.path; printer = ''; copies = 1; dryRun = $true; enhance = 'auto' }
        Check '试运行给出的方式不是空壳' (-not [string]::IsNullOrWhiteSpace($r.method)) 'method 为空'
    }

    $bad = Api 'POST' '/api/print-one' @{ path = 'C:\Windows\win.ini'; printer = ''; copies = 1; dryRun = $true }
    Check '拒绝打印不在扫描结果里的文件' ($bad.ok -eq $false) '竟然通过了'

    $bad2 = Api 'POST' '/api/print-one' @{ path = ''; printer = ''; copies = 1; dryRun = $true }
    Check '拒绝空路径' ($bad2.ok -eq $false)

    $bad3 = Api 'POST' '/api/print-one' @{ path = (Join-Path $resume '不存在的.pdf'); printer = ''; copies = 1; dryRun = $true }
    Check '拒绝不存在的文件' ($bad3.ok -eq $false)

    $reset = Api 'POST' '/api/reset-printed' @{}
    Check '/api/reset-printed 可用' ($reset.ok -eq $true)

    $end = Api 'POST' '/api/end-batch' @{}
    Check '/api/end-batch 可用' ($end.ok -eq $true)

    # --------------------------------------------------------
    Section '静态资源'
    # --------------------------------------------------------

    foreach ($p in @('/', '/app.js', '/style.css')) {
        $w = Invoke-WebRequest -Uri ($Base + $p) -UseBasicParsing -TimeoutSec 15
        Check "静态资源 $p 可访问" ($w.StatusCode -eq 200 -and $w.RawContentLength -gt 0)
    }

    # --------------------------------------------------------
    Section '安全：路径穿越'
    # --------------------------------------------------------

    foreach ($p in @('/../server.ps1', '/..%2fserver.ps1', '/....//server.ps1', '/web/../server.ps1', '/%2e%2e/server.ps1')) {
        $r = RawRequest @{ Host = "127.0.0.1:$Port" } 'GET' $p
        Check "路径穿越被拦截: $p" ((StatusOf $r) -eq 404) ('返回 ' + (StatusOf $r))
    }

    # --------------------------------------------------------
    Section '安全：DNS 重绑定 / CSRF'
    # --------------------------------------------------------

    $json = '{"folder":"' + $resume.Replace('\', '\\') + '"}'

    $r1 = RawRequest @{ Host = 'evil.example.com'; 'Content-Type' = 'application/json' } 'POST' '/api/scan' $json
    Check 'Host 非回环被拒绝（DNS 重绑定）' ((StatusOf $r1) -eq 403) ('返回 ' + (StatusOf $r1))

    $r2 = RawRequest @{ Host = "127.0.0.1:$Port"; Origin = 'http://evil.example.com'; 'Content-Type' = 'application/json' } 'POST' '/api/scan' $json
    Check '跨站 Origin 被拒绝（CSRF）' ((StatusOf $r2) -eq 403) ('返回 ' + (StatusOf $r2))

    $r3 = RawRequest @{ Host = "127.0.0.1:$Port"; 'Content-Type' = 'text/plain' } 'POST' '/api/print-one' '{"path":"C:\\Windows\\win.ini"}'
    Check '表单式 text/plain 写请求被拒绝（CSRF）' ((StatusOf $r3) -eq 403) ('返回 ' + (StatusOf $r3))

    $r4 = RawRequest @{ Host = "127.0.0.1:$Port"; 'Content-Type' = 'application/x-www-form-urlencoded' } 'POST' '/api/print-one' 'path=x'
    Check 'urlencoded 写请求被拒绝（CSRF）' ((StatusOf $r4) -eq 403) ('返回 ' + (StatusOf $r4))

    $r5 = RawRequest @{ Host = "127.0.0.1:$Port"; 'Content-Type' = 'application/json' } 'POST' '/api/info' '{}'
    Check '正常回环请求不被误伤' ((StatusOf $r5) -eq 200) ('返回 ' + (StatusOf $r5))

    # --------------------------------------------------------
    Section '安全：请求体上限'
    # --------------------------------------------------------

    $big = '{"x":"' + ('A' * 2000000) + '"}'
    $r6 = ''
    try { $r6 = RawRequest @{ Host = "127.0.0.1:$Port"; 'Content-Type' = 'application/json' } 'POST' '/api/scan' $big }
    catch { $r6 = '(连接被重置)' }
    $st6 = StatusOf $r6
    Check '超大请求体被拒绝' ($st6 -eq 413 -or $st6 -eq 400 -or $r6 -eq '(连接被重置)') ('返回 ' + $st6)

    # --------------------------------------------------------
    Section '隐私：不在磁盘上乱写'
    # --------------------------------------------------------

    $tmpFiles = @(Get-ChildItem -LiteralPath $tmp -Recurse -File -ErrorAction SilentlyContinue |
                  Where-Object { $_.FullName -notmatch '\\_resumes\\' })
    Check '没有把简历复制到别处' (
        @($tmpFiles | Where-Object { $_.Name -match '张三|李四|产品经理' }).Count -eq 0
    ) '发现了简历副本'

} catch {
    Bad '测试执行过程' ('' + $_.Exception.Message)
    Write-Host $_.ScriptStackTrace -ForegroundColor DarkGray
} finally {
    try { if ($proc -and -not $proc.HasExited) { Stop-Process -Id $proc.Id -Force -ErrorAction SilentlyContinue } } catch { }
    Start-Sleep -Milliseconds 500
    if (-not $KeepTemp) { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
    else { Write-Host ('临时目录保留在: ' + $tmp) -ForegroundColor DarkGray }
}

# ============================================================
Write-Host ''
Write-Host ('══════════════════════════════════════════════════') -ForegroundColor Cyan
Write-Host ("  通过 $script:Pass 项，失败 $script:Fail 项") -ForegroundColor $(if ($script:Fail -eq 0) { 'Green' } else { 'Red' })
if ($script:Fail -gt 0) {
    Write-Host ''
    Write-Host '  失败明细：' -ForegroundColor Red
    foreach ($f in $script:Failures) { Write-Host ('   - ' + $f) -ForegroundColor Red }
}
Write-Host ('══════════════════════════════════════════════════') -ForegroundColor Cyan
Write-Host ''

if ($script:Fail -gt 0) { exit 1 }
exit 0
