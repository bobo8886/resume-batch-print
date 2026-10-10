<#
  简历一键打印 —— 本地打印服务 (Windows PowerShell 5.1，零依赖、免安装、可整包拷贝)

  打印链路：
    .pdf                  → SumatraPDF 静默打印
    .doc/.docx/.rtf/...   → Word(COM) → WPS(COM) → 转成 PDF 再用 SumatraPDF 打
    打印机选择             → 通过 COM 的 ActivePrinter 指定，不动系统默认打印机

  两个关键实现细节：
  1) 所有 Office 调用都走「纯 IDispatch」(见下方 C# 助手)，彻底绕开 Office 互操作程序集(PIA)。
     很多电脑上 PIA 是坏的或没装，用 PowerShell 原生 $app.Documents 会拿到 null，
     而纯 IDispatch 不受影响。
  2) SumatraPDF 在「自身程序路径含非 ASCII 字符」时 StartDoc 必定失败，
     所以启动时若发现程序目录含中文，会自动把它复制到纯英文目录再调用。
#>
[CmdletBinding()]
param(
    [int]$Port = 17890,
    [switch]$NoBrowser
)

$ErrorActionPreference = 'Stop'
try { [Console]::OutputEncoding = [System.Text.Encoding]::UTF8 } catch { }
try { $Host.UI.RawUI.WindowTitle = '简历一键打印（关闭本窗口即退出）' } catch { }

$Root           = $PSScriptRoot
$WebDir         = Join-Path $Root 'web'
$BinDir         = Join-Path $Root 'bin'
$RecordDir      = Join-Path $Root '打印记录'
$ConfigFile     = Join-Path $Root 'config.json'
$PrintedFile    = Join-Path $Root '打印记录\printed.json'
$SumatraBundled = Join-Path $BinDir 'SumatraPDF.exe'
$TempDir        = Join-Path $env:TEMP 'ResumePrintWork'
$SupportedExt   = @('.pdf', '.doc', '.docx', '.rtf', '.txt', '.odt', '.wps')

foreach ($d in @($RecordDir, $TempDir)) {
    if (-not (Test-Path $d)) { New-Item -ItemType Directory -Force -Path $d | Out-Null }
}

# ============================ 控制台日志 ============================

function Log([string]$msg, [string]$level = 'info') {
    switch ($level) {
        'err' { Write-Host ("  ! " + $msg) -ForegroundColor Red }
        'warn' { Write-Host ("  ~ " + $msg) -ForegroundColor Yellow }
        'ok' { Write-Host ("  + " + $msg) -ForegroundColor Green }
        default { Write-Host ("    " + $msg) -ForegroundColor DarkGray }
    }
}

# ============================ 通用工具 ============================

function Read-JsonFile([string]$Path) {
    if (Test-Path $Path) {
        try { return (Get-Content -Raw -Encoding UTF8 $Path | ConvertFrom-Json) } catch { return $null }
    }
    return $null
}
function Write-JsonFile([string]$Path, $Obj) {
    $json = $Obj | ConvertTo-Json -Depth 8
    # 原子写：先写同目录下的临时文件，再用 File.Replace 整体替换。
    # 以前直接 WriteAllText（先截断后写入），崩在中间会留下一个空/截断的
    # config.json 或 printed.json，下次启动静默回落默认值。
    # 注意：不能给 File.Replace 的第三个参数传 $null —— PowerShell 会把它
    # 转成空字符串，而空串是非法路径，会直接抛"路径的形式不合法"。
    $tmp = $Path + '.tmp'
    $bak = $Path + '.bak'
    [System.IO.File]::WriteAllText($tmp, $json, (New-Object System.Text.UTF8Encoding($false)))
    try {
        if ([System.IO.File]::Exists($Path)) {
            [System.IO.File]::Replace($tmp, $Path, $bak)
        } else {
            [System.IO.File]::Move($tmp, $Path)
        }
    } finally {
        Remove-Item -LiteralPath $bak -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
    }
}
function Get-Prop($Obj, [string]$Name, $Default) {
    if ($null -eq $Obj) { return $Default }
    if ($Obj.PSObject.Properties.Name -contains $Name) { return $Obj.$Name }
    return $Default
}
function Test-AsciiPath([string]$p) { return ($p -match '^[\x20-\x7E]+$') }

# ============================ 配置 / 记录 ============================

$script:Config = @{ lastFolder = ''; printer = ''; copies = 1; recursive = $true; recentFolders = @(); enhance = 'auto' }
$c = Read-JsonFile $ConfigFile
if ($c) {
    foreach ($k in @('lastFolder', 'printer', 'copies', 'recursive', 'recentFolders', 'enhance')) {
        if ($c.PSObject.Properties.Name -contains $k) { $script:Config[$k] = $c.$k }
    }
}
if ($null -eq $script:Config.recentFolders) { $script:Config.recentFolders = @() }
function Save-Config { Write-JsonFile $ConfigFile $script:Config }
function Add-RecentFolder([string]$Folder) {
    if ([string]::IsNullOrWhiteSpace($Folder)) { return }
    $list = @($Folder)
    foreach ($f in @($script:Config.recentFolders)) {
        if ($f -and $f -ne $Folder -and -not ($list -contains $f)) { $list += $f }
    }
    $script:Config.recentFolders = @($list | Select-Object -First 8)
}

$script:Printed = @{}
# 打印历史的上限。以前没有任何上限：每成功打印一份就加一个键，
# 然后把整个 map 重写一遍，安装期内文件和内存都无界增长。
$script:PrintedMax = 5000
$pr = Read-JsonFile $PrintedFile
if ($pr) { foreach ($p in $pr.PSObject.Properties) { $script:Printed[$p.Name] = $p.Value } }
function Save-Printed {
    # 超上限时按值里的时间戳丢掉最旧的，只留最近 $script:PrintedMax 条
    if ($script:Printed.Count -gt $script:PrintedMax) {
        $keep = @($script:Printed.GetEnumerator() |
                  Sort-Object -Property @{ Expression = { [string]$_.Value } } -Descending |
                  Select-Object -First $script:PrintedMax)
        $fresh = @{}
        foreach ($e in $keep) { $fresh[$e.Key] = $e.Value }
        $script:Printed = $fresh
    }
    Write-JsonFile $PrintedFile $script:Printed
}

# 本次会话里"被扫描列出来过"的文件，只有这些才允许打印
$script:KnownFiles = @{}

# ============================ 纯 IDispatch 助手 ============================
# 用 IDispatch::GetIDsOfNames + Invoke 直接调用 COM 成员，不经过任何互操作程序集。
# 遇到「被调用方忙」(RPC_E_CALL_REJECTED / SERVERCALL_RETRYLATER) 会自动重试。

$script:DispReady = $false
$csDisp = @'
using System;
using System.Runtime.InteropServices;

public static class Disp
{
    [StructLayout(LayoutKind.Sequential)]
    private struct DISPPARAMS { public IntPtr rgvarg; public IntPtr rgdispidNamedArgs; public int cArgs; public int cNamedArgs; }

    [ComImport, Guid("00020400-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IDispatch
    {
        [PreserveSig] int GetTypeInfoCount(out int pctinfo);
        [PreserveSig] int GetTypeInfo(int iTInfo, int lcid, out IntPtr ppTInfo);
        [PreserveSig] int GetIDsOfNames(ref Guid riid, [MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] rgszNames, int cNames, int lcid, [Out, MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
        [PreserveSig] int Invoke(int dispIdMember, ref Guid riid, int lcid, short wFlags, ref DISPPARAMS pDispParams, IntPtr pVarResult, IntPtr pExcepInfo, IntPtr puArgErr);
    }

    [DllImport("oleaut32.dll", ExactSpelling = true)]
    private static extern int VariantClear(IntPtr pvarg);

    private const short M_GET = 0x2;
    private const short M_PUT = 0x4;
    private const short M_CALL = 0x1;
    private const int LCID = 0x0400;
    private const int DISPID_PROPERTYPUT = -3;
    private const int RPC_E_CALL_REJECTED = unchecked((int)0x80010001);
    private const int RPC_E_SERVERCALL_RETRYLATER = unchecked((int)0x8001010A);

    private static IDispatch D(object o)
    {
        if (o == null) throw new ArgumentNullException("o");
        IDispatch d = o as IDispatch;
        if (d != null) return d;
        IntPtr p = Marshal.GetIDispatchForObject(o);
        return (IDispatch)Marshal.GetObjectForIUnknown(p);
    }

    private static object Once(object o, string name, short flags, object[] args, bool want)
    {
        IDispatch d = D(o);
        Guid iid = Guid.Empty;
        int[] ids = new int[1];
        int hr = d.GetIDsOfNames(ref iid, new string[] { name }, 1, LCID, ids);
        if (hr != 0) Marshal.ThrowExceptionForHR(hr);

        int n = (args == null) ? 0 : args.Length;
        int vs = (IntPtr.Size == 8) ? 24 : 16;
        DISPPARAMS dp = new DISPPARAMS();
        IntPtr rg = IntPtr.Zero, named = IntPtr.Zero, res = IntPtr.Zero;
        try
        {
            if (n > 0)
            {
                rg = Marshal.AllocCoTaskMem(vs * n);
                for (int i = 0; i < n; i++)
                    Marshal.GetNativeVariantForObject(args[n - 1 - i], (IntPtr)((long)rg + i * vs));
                dp.rgvarg = rg;
                dp.cArgs = n;
            }
            if (flags == M_PUT)
            {
                named = Marshal.AllocCoTaskMem(4);
                Marshal.WriteInt32(named, DISPID_PROPERTYPUT);
                dp.rgdispidNamedArgs = named;
                dp.cNamedArgs = 1;
            }
            if (want) res = Marshal.AllocCoTaskMem(vs);
            hr = d.Invoke(ids[0], ref iid, LCID, flags, ref dp, res, IntPtr.Zero, IntPtr.Zero);
            if (hr != 0) Marshal.ThrowExceptionForHR(hr);
            if (!want || res == IntPtr.Zero) return null;
            // 有些成员会回一个 .NET 无法封送的 VARIANT（例如属性写入成功时的返回值），
            // 这不是调用失败，直接当作"没有返回值"。
            try { return Marshal.GetObjectForNativeVariant(res); }
            catch { return null; }
        }
        finally
        {
            if (rg != IntPtr.Zero)
            {
                for (int i = 0; i < n; i++) VariantClear((IntPtr)((long)rg + i * vs));
                Marshal.FreeCoTaskMem(rg);
            }
            if (named != IntPtr.Zero) Marshal.FreeCoTaskMem(named);
            if (res != IntPtr.Zero) { VariantClear(res); Marshal.FreeCoTaskMem(res); }
        }
    }

    private static object Run(object o, string name, short flags, object[] args, bool want)
    {
        int attempt = 0;
        while (true)
        {
            try { return Once(o, name, flags, args, want); }
            catch (COMException ex)
            {
                if ((ex.HResult == RPC_E_CALL_REJECTED || ex.HResult == RPC_E_SERVERCALL_RETRYLATER) && attempt < 60)
                {
                    attempt++;
                    System.Threading.Thread.Sleep(150);
                    continue;
                }
                throw;
            }
        }
    }

    public static object Get(object o, string name) { return Run(o, name, M_GET, null, true); }
    public static object Call(object o, string name, params object[] a) { return Run(o, name, M_CALL, a, true); }
    public static void Set(object o, string name, object v) { Run(o, name, M_PUT, new object[] { v }, false); }
}
'@

try {
    Add-Type -TypeDefinition $csDisp -Language CSharp -ErrorAction Stop
    $script:DispReady = $true
} catch {
    $script:DispReady = ($null -ne ('Disp' -as [type]))
}

# ============================ 扫描件加深助手 ============================
# 很多简历 PDF 是"手机拍的/扫描的图片"，正文是浅灰色，直接打印会发虚发淡。
# 这里把 PDF 每页重新渲染成位图、按 gamma 加深，再拼成一个新 PDF 交给 SumatraPDF 打印。
# gamma 用 256 项查找表实现，PDF 用 zlib(FlateDecode) 压缩，全在 C# 里跑，够快。

$script:EnhReady = $false
$csEnh = @'
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;

public static class EnhPdf
{
    public static byte[] ToRgbGamma(Bitmap bmp, double gamma)
    {
        int w = bmp.Width, h = bmp.Height;
        var rect = new Rectangle(0, 0, w, h);
        var data = bmp.LockBits(rect, ImageLockMode.ReadOnly, PixelFormat.Format24bppRgb);
        byte[] outp = new byte[w * 3 * h];
        try
        {
            byte[] lut = new byte[256];
            for (int i = 0; i < 256; i++)
            {
                double v = 255.0 * Math.Pow(i / 255.0, gamma);
                lut[i] = (byte)(v < 0 ? 0 : (v > 255 ? 255 : Math.Round(v)));
            }
            int stride = data.Stride;
            byte[] row = new byte[stride];
            for (int y = 0; y < h; y++)
            {
                Marshal.Copy(IntPtr.Add(data.Scan0, y * stride), row, 0, stride);
                int o = y * w * 3;
                for (int x = 0; x < w; x++)
                {
                    outp[o] = lut[row[x * 3]];
                    outp[o + 1] = lut[row[x * 3 + 1]];
                    outp[o + 2] = lut[row[x * 3 + 2]];
                    o += 3;
                }
            }
        }
        finally { bmp.UnlockBits(data); }
        return outp;
    }

    private static byte[] Zlib(byte[] raw)
    {
        using (var ms = new MemoryStream())
        {
            ms.WriteByte(0x78); ms.WriteByte(0x9C);
            using (var ds = new DeflateStream(ms, CompressionMode.Compress, true)) ds.Write(raw, 0, raw.Length);
            uint a = 1, b = 0;
            for (int i = 0; i < raw.Length; i++)
            {
                a = (a + raw[i]) % 65521;
                b = (b + a) % 65521;
            }
            uint adler = (b << 16) | a;
            ms.WriteByte((byte)(adler >> 24)); ms.WriteByte((byte)(adler >> 16));
            ms.WriteByte((byte)(adler >> 8));  ms.WriteByte((byte)adler);
            return ms.ToArray();
        }
    }

    public class Builder
    {
        private readonly List<byte[]> _imgs = new List<byte[]>();
        private readonly List<int> _w = new List<int>();
        private readonly List<int> _h = new List<int>();
        private readonly List<double> _pw = new List<double>();
        private readonly List<double> _ph = new List<double>();

        public void AddPage(byte[] rgb, int w, int h, double pageWpt, double pageHpt)
        {
            _imgs.Add(Zlib(rgb)); _w.Add(w); _h.Add(h); _pw.Add(pageWpt); _ph.Add(pageHpt);
        }
        public int PageCount { get { return _imgs.Count; } }

        public void Save(string path)
        {
            int n = _imgs.Count;
            var offsets = new List<long>();
            using (var fs = new FileStream(path, FileMode.Create, FileAccess.Write))
            {
                Action<string> Wr = s => { var b = Encoding.ASCII.GetBytes(s); fs.Write(b, 0, b.Length); };
                Action<int> Obj = id => { offsets.Add(fs.Position); Wr(id + " 0 obj\n"); };
                Wr("%PDF-1.4\n");

                Obj(1); Wr("<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");

                var kids = new StringBuilder();
                for (int i = 0; i < n; i++) kids.Append((3 + i * 3) + " 0 R ");
                Obj(2); Wr("<< /Type /Pages /Count " + n + " /Kids [" + kids.ToString().Trim() + "] >>\nendobj\n");

                var inv = CultureInfo.InvariantCulture;
                for (int i = 0; i < n; i++)
                {
                    int pageId = 3 + i * 3, contId = 4 + i * 3, imgId = 5 + i * 3;
                    string pw = _pw[i].ToString("0.###", inv);
                    string ph = _ph[i].ToString("0.###", inv);

                    Obj(pageId);
                    Wr("<< /Type /Page /Parent 2 0 R /MediaBox [0 0 " + pw + " " + ph + "] "
                       + "/Resources << /XObject << /Im0 " + imgId + " 0 R >> >> /Contents " + contId + " 0 R >>\nendobj\n");

                    string content = "q " + pw + " 0 0 " + ph + " 0 0 cm /Im0 Do Q\n";
                    var cb = Encoding.ASCII.GetBytes(content);
                    Obj(contId); Wr("<< /Length " + cb.Length + " >>\nstream\n"); fs.Write(cb, 0, cb.Length); Wr("endstream\nendobj\n");

                    var img = _imgs[i];
                    Obj(imgId);
                    Wr("<< /Type /XObject /Subtype /Image /Width " + _w[i] + " /Height " + _h[i]
                       + " /ColorSpace /DeviceRGB /BitsPerComponent 8 /Filter /FlateDecode /Length " + img.Length + " >>\nstream\n");
                    fs.Write(img, 0, img.Length); Wr("\nendstream\nendobj\n");
                }

                long xref = fs.Position;
                Wr("xref\n0 " + (offsets.Count + 1) + "\n");
                Wr("0000000000 65535 f \n");
                foreach (var o in offsets) Wr(o.ToString("0000000000") + " 00000 n \n");
                Wr("trailer\n<< /Size " + (offsets.Count + 1) + " /Root 1 0 R >>\nstartxref\n" + xref + "\n%%EOF\n");
            }
        }
    }
}
'@

try {
    Add-Type -TypeDefinition $csEnh -Language CSharp -ReferencedAssemblies System.Drawing, System.IO.Compression -ErrorAction Stop
    $script:EnhReady = $true
} catch {
    $script:EnhReady = ($null -ne ('EnhPdf' -as [type]))
}

# ============================ PDF 引擎定位 ============================

function Resolve-SumatraEngine([string]$Bundled) {
    if (-not (Test-Path $Bundled)) { return $Bundled }
    if (Test-AsciiPath $Bundled) { return $Bundled }

    $cands = @()
    if ($env:TEMP)        { $cands += (Join-Path $env:TEMP 'ResumePrintEngine') }
    if ($env:PUBLIC)      { $cands += (Join-Path $env:PUBLIC 'ResumePrintEngine') }
    if ($env:ProgramData) { $cands += (Join-Path $env:ProgramData 'ResumePrintEngine') }

    foreach ($d in $cands) {
        if (-not (Test-AsciiPath $d)) { continue }
        try {
            if (-not (Test-Path $d)) { New-Item -ItemType Directory -Force -Path $d | Out-Null }
            $target = Join-Path $d 'SumatraPDF.exe'
            # 复用条件从"字节长度相同"改成"哈希与内置副本相同"。
            # 长度相等是攻击者可满足的：往一个本机可写的搬迁目录里放一个
            # 填充到同样长度的伪造引擎，就会被一直复用并在每次打印时执行。
            $need = $true
            if (Test-Path $target) {
                try {
                    if ((Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -eq (Get-FileHash -LiteralPath $Bundled -Algorithm SHA256).Hash) { $need = $false }
                } catch { $need = $true }
            }
            if ($need) { Copy-Item -LiteralPath $Bundled -Destination $target -Force }
            # 复制/复用后仍然只有哈希相同才认，否则退回内置路径，不去执行一个来路不明的文件
            try {
                if ((Get-FileHash -LiteralPath $target -Algorithm SHA256).Hash -ne (Get-FileHash -LiteralPath $Bundled -Algorithm SHA256).Hash) { continue }
            } catch { continue }
            $s = Join-Path $BinDir 'SumatraPDF-settings.txt'
            if (Test-Path $s) { Copy-Item -LiteralPath $s -Destination (Join-Path $d 'SumatraPDF-settings.txt') -Force }
            if (Test-Path $target) { return $target }
        } catch { }
    }
    return $Bundled
}

$script:SumatraPath = Resolve-SumatraEngine $SumatraBundled
$script:SumatraRelocated = ($script:SumatraPath -ne $SumatraBundled)

# ---- 引擎按需下载（仓库里不带 15MB 的二进制，首次使用时下载并校验 SHA-256） ----
$script:EngineUrl       = 'https://www.sumatrapdfreader.org/dl/rel/3.5.2/SumatraPDF-3.5.2-64.zip'
$script:EngineUrlBackup = 'https://github.com/sumatrapdfreader/sumatrapdf/releases/download/3.5.2/SumatraPDF-3.5.2-64.zip'
$script:EngineSha256    = '290E4AA7ED64C728138711C011E89AAB7AA48DBC1AE430371DC2BE4100B92BF0'
$script:EngineVersion   = '3.5.2'

function Get-AsciiEngineDir {
    $cands = @()
    if ($env:TEMP)        { $cands += (Join-Path $env:TEMP 'ResumePrintEngine') }
    if ($env:PUBLIC)      { $cands += (Join-Path $env:PUBLIC 'ResumePrintEngine') }
    if ($env:ProgramData) { $cands += (Join-Path $env:ProgramData 'ResumePrintEngine') }
    foreach ($d in $cands) {
        if (-not (Test-AsciiPath $d)) { continue }
        try {
            if (-not (Test-Path $d)) { New-Item -ItemType Directory -Force -Path $d | Out-Null }
            return $d
        } catch { }
    }
    return ''
}

function Test-DirWritable([string]$Dir) {
    try {
        if (-not (Test-Path $Dir)) { New-Item -ItemType Directory -Force -Path $Dir | Out-Null }
        $probe = Join-Path $Dir ('.w' + [guid]::NewGuid().ToString('N'))
        [System.IO.File]::WriteAllText($probe, 'x')
        Remove-Item $probe -Force -ErrorAction SilentlyContinue
        return $true
    } catch { return $false }
}

function Write-EngineSettings([string]$Dir) {
    $s = @(
        'CheckForUpdates = false', 'RestoreSession = false',
        'RememberOpenedFiles = false', 'ShowToc = false', 'ShowStartPage = false'
    ) -join "`r`n"
    [System.IO.File]::WriteAllText((Join-Path $Dir 'SumatraPDF-settings.txt'), $s, (New-Object System.Text.UTF8Encoding($false)))
}

function Install-PdfEngine {
    $targetDir = $BinDir
    if (-not (Test-DirWritable $targetDir)) {
        $targetDir = Get-AsciiEngineDir
        if (-not $targetDir) { throw '程序目录不可写，也找不到可写的临时目录来存放打印引擎' }
    }

    $tmp = Join-Path $env:TEMP ('sumatra_' + [guid]::NewGuid().ToString('N'))
    New-Item -ItemType Directory -Force -Path $tmp | Out-Null
    $oldProgress = $ProgressPreference
    try {
        $ProgressPreference = 'SilentlyContinue'
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch { }

        $zip = Join-Path $tmp 'engine.zip'
        $downloaded = $false
        $lastErr = ''
        foreach ($u in @($script:EngineUrl, $script:EngineUrlBackup)) {
            try {
                Invoke-WebRequest -Uri $u -OutFile $zip -UseBasicParsing -TimeoutSec 300 -ErrorAction Stop
                if ((Test-Path $zip) -and (Get-Item $zip).Length -gt 1MB) { $downloaded = $true; break }
            } catch { $lastErr = '' + $_.Exception.Message }
        }
        if (-not $downloaded) { throw "下载失败：$lastErr" }

        Expand-Archive -LiteralPath $zip -DestinationPath $tmp -Force
        $exe = @(Get-ChildItem -LiteralPath $tmp -Filter '*.exe' -Recurse -ErrorAction SilentlyContinue | Select-Object -First 1)
        if ($exe.Count -eq 0) { throw '下载包里没有找到打印引擎可执行文件' }

        $hash = (Get-FileHash -LiteralPath $exe[0].FullName -Algorithm SHA256).Hash
        if ($hash -ne $script:EngineSha256) {
            throw "安全校验未通过：文件哈希与预期不符，已放弃安装（预期 $($script:EngineSha256)，实际 $hash）"
        }

        $final = Join-Path $targetDir 'SumatraPDF.exe'
        Copy-Item -LiteralPath $exe[0].FullName -Destination $final -Force
        Write-EngineSettings $targetDir

        $script:SumatraPath = Resolve-SumatraEngine $SumatraBundled
        # 解析出来的路径必须真的就是刚才校验通过的那个文件；
        # 否则说明搬迁目录里有个不同的二进制，宁可退回 $final，也不去执行它。
        if (-not (Test-Path $script:SumatraPath)) {
            $script:SumatraPath = $final
        } elseif ((Get-FileHash -LiteralPath $script:SumatraPath -Algorithm SHA256).Hash -ne $script:EngineSha256) {
            Write-EngineSettings $targetDir
            $script:SumatraPath = $final
        }
        $script:SumatraRelocated = ($script:SumatraPath -ne $SumatraBundled)
        $script:EngineStatus.sumatra = (Test-Path $script:SumatraPath)
        return $final
    } finally {
        $ProgressPreference = $oldProgress
        Remove-Item $tmp -Recurse -Force -ErrorAction SilentlyContinue
    }
}

# ============================ 引擎探测 ============================

function Test-ComProgId([string]$ProgId) {
    foreach ($hive in @('HKLM:\SOFTWARE\Classes', 'HKCU:\SOFTWARE\Classes')) {
        if (Test-Path (Join-Path $hive $ProgId)) { return $true }
    }
    return $false
}

$script:HasWord = Test-ComProgId 'Word.Application'
$script:HasWps = ((Test-ComProgId 'KWPS.Application') -or (Test-ComProgId 'WPS.Application'))
$script:EngineStatus = [ordered]@{
    sumatra = (Test-Path $script:SumatraPath)
    word    = $false
    wps     = $false
    tested  = $false
    message = ''
    disp    = $script:DispReady
}
$script:OfficeOrder = @()
if ($script:HasWord) { $script:OfficeOrder += 'Word.Application' }
if ($script:HasWps) { $script:OfficeOrder += 'KWPS.Application' }

function Get-EngineReport {
    [ordered]@{
        sumatra     = [bool]$script:EngineStatus.sumatra
        word        = [bool]$script:EngineStatus.word
        wps         = [bool]$script:EngineStatus.wps
        wordPresent = [bool]$script:HasWord
        wpsPresent  = [bool]$script:HasWps
        tested      = [bool]$script:EngineStatus.tested
        message     = [string]$script:EngineStatus.message
        disp        = [bool]$script:EngineStatus.disp
    }
}

function Invoke-EngineSelfTest {
    $notes = @()
    $script:EngineStatus.word = $false
    $script:EngineStatus.wps = $false
    Close-OfficeApp

    if (-not $script:DispReady) {
        $notes += 'COM 调用助手未能加载'
    } else {
        foreach ($prog in @('Word.Application', 'KWPS.Application')) {
            $present = if ($prog -eq 'Word.Application') { $script:HasWord } else { $script:HasWps }
            if (-not $present) { continue }
            $label = if ($prog -like '*KWPS*') { 'WPS' } else { 'Word' }
            $app = $null
            $before = Get-OfficeProcessIds
            $owned = @()
            try {
                $app = New-Object -ComObject $prog
                $owned = Get-NewOfficeProcessIds $before
                $docs = $null
                for ($i = 0; $i -lt 24 -and $null -eq $docs; $i++) {
                    try { $docs = [Disp]::Get($app, 'Documents') } catch { $docs = $null }
                    if ($null -eq $docs) { Start-Sleep -Milliseconds 500 }
                }
                if ($null -ne $docs) {
                    if ($prog -like '*KWPS*') { $script:EngineStatus.wps = $true } else { $script:EngineStatus.word = $true }
                } else {
                    $notes += "$label 组件响应异常"
                }
            } catch {
                $notes += "$label 不可用"
            } finally {
                if ($app) { try { [Disp]::Call($app, 'Quit') | Out-Null } catch { }; try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($app) } catch { } }
                Start-Sleep -Milliseconds 400
                # 只终止本次自检**自己启动**的进程。
                # 以前这里是 `Get-Process WINWORD, wps | Stop-Process -Force`，
                # 会把操作者自己打开、还有未保存内容的文档一起强杀掉。
                Stop-OfficeProcessId $owned
            }
        }
    }

    $order = @()
    if ($script:EngineStatus.word) { $order += 'Word.Application' }
    if ($script:EngineStatus.wps) { $order += 'KWPS.Application' }
    foreach ($p in @('Word.Application', 'KWPS.Application')) {
        if (($order -notcontains $p) -and ($script:OfficeOrder -contains $p)) { $order += $p }
    }
    $script:OfficeOrder = $order

    $script:EngineStatus.tested = $true
    $script:EngineStatus.message = ($notes -join '；')
    return (Get-EngineReport)
}

# ============================ 打印机 ============================

function Get-PrinterKind([string]$Name, [string]$Port, [string]$Driver) {
    $p = ([string]$Port).ToLowerInvariant()
    if ($p -eq '' -or $p -eq 'nul:' -or $p -eq 'portprompt:' -or $p -eq 'file:') { return 'virtual' }
    $words = 'OneNote|Print To PDF|PrintToPDF|WPS PDF|Kingsoft|XPS|Virtual|虚拟|传真|Fax'
    if ($Name -match $words -or $Driver -match $words) { return 'virtual' }
    return 'real'
}

function Get-CurrentDefaultPrinter {
    try {
        $d = (Get-ItemProperty -Path 'HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Windows' -Name Device -ErrorAction Stop).Device
        if ($d) { return [string](($d -split ',')[0]) }
    } catch { }
    return ''
}

function Get-PrinterList {
    $names = @()
    try {
        $names = @(Get-ChildItem 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers' -ErrorAction Stop |
                   Select-Object -ExpandProperty PSChildName)
    } catch { $names = @() }
    if ($names.Count -eq 0) {
        try { $names = @(Get-CimInstance -ClassName Win32_Printer -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Name) } catch { }
    }

    $def = Get-CurrentDefaultPrinter
    $list = @()
    foreach ($n in ($names | Sort-Object)) {
        $port = ''; $driver = ''
        try {
            $pk = Get-ItemProperty ("HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\" + $n) -ErrorAction SilentlyContinue
            if ($pk) {
                if ($pk.Port) { $port = [string]$pk.Port }
                if ($pk.'Printer Driver') { $driver = [string]$pk.'Printer Driver' }
            }
        } catch { }
        $list += [ordered]@{
            name    = [string]$n
            default = ([string]$n -eq $def)
            port    = $port
            driver  = $driver
            kind    = (Get-PrinterKind $n $port $driver)
            network = ($port -like '\\*' -or $port -like 'IP_*' -or $port -like 'WSD*')
        }
    }
    if (-not $def -and $list.Count -gt 0) { $def = [string]$list[0].name }
    return @{ printers = $list; defaultPrinter = $def }
}

function Get-PrinterPortName([string]$Printer) {
    try {
        # -LiteralPath：打印机名里如果带 [ ] * ? 会被当成通配符模式，必须按字面查
        $pk = Get-ItemProperty -LiteralPath ("HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers\" + $Printer) -ErrorAction SilentlyContinue
        if ($pk -and $pk.Port) { return [string]$pk.Port }
    } catch { }
    return ''
}

function Test-PrinterExists([string]$Name) {
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    # 精确成员测试，不做通配符路径解析。
    # 以前是 Test-Path 拼注册表路径，而 Test-Path 把 -Path 当**通配符模式**，
    # 于是 "*" / "HP*" / "[a-z]*" 这类值也能"通过校验"，随后被原样拼进
    # Start-Process 的参数串里（也让配置里存下一个根本不是打印机的名字）。
    # 现在只接受与已安装打印机键名**完全相等**的值。
    if ($Name -match '["*?\[\]]') { return $false }
    try {
        $keys = @(Get-ChildItem 'HKLM:\SYSTEM\CurrentControlSet\Control\Print\Printers' -ErrorAction SilentlyContinue |
                  Select-Object -ExpandProperty PSChildName)
        return ($keys -ccontains $Name)
    } catch { return $false }
}

function Get-SuggestedPrinter($pl) {
    $stored = [string]$script:Config.printer
    $storedItem = @($pl.printers | Where-Object { $_.name -eq $stored } | Select-Object -First 1)
    if ($storedItem.Count -gt 0 -and $storedItem[0].kind -eq 'real') { return $stored }
    $defItem = @($pl.printers | Where-Object { $_.name -eq $pl.defaultPrinter } | Select-Object -First 1)
    if ($defItem.Count -gt 0 -and $defItem[0].kind -eq 'real') { return [string]$pl.defaultPrinter }
    $firstReal = @($pl.printers | Where-Object { $_.kind -eq 'real' } | Select-Object -First 1)
    if ($firstReal.Count -gt 0) { return [string]$firstReal[0].name }
    if ($storedItem.Count -gt 0) { return $stored }
    return [string]$pl.defaultPrinter
}

# ============================ Office 实例复用 ============================

$script:OfficeApp = $null
$script:OfficeProgId = ''
$script:OfficeLastUse = [datetime]::MinValue
# 只记录**本工具自己启动**的 Office 进程 id。绝不做"按映像名无差别强杀"。
$script:OfficeOwnedPids = @()

function Get-OfficeProcessIds {
    return @(Get-Process -Name WINWORD, wps -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
}

# 创建 COM 实例前后各取一次快照，差集才是我们"新开出来"的进程。
# 如果 New-Object 附着到了操作者已经打开的实例上，差集就是空的 —— 那就一个都不杀。
function Get-NewOfficeProcessIds([int[]]$Before) {
    if ($null -eq $Before) { $Before = @() }
    return @(Get-OfficeProcessIds | Where-Object { $Before -notcontains $_ })
}

function Stop-OfficeProcessId([int[]]$Ids) {
    foreach ($procId in @($Ids)) {
        if (-not $procId) { continue }
        try {
            $p = Get-Process -Id $procId -ErrorAction SilentlyContinue
            if ($p) { $p.Kill(); [void]$p.WaitForExit(5000) }
        } catch { }
    }
}

function Close-OfficeApp {
    if ($script:OfficeApp) {
        try { [Disp]::Call($script:OfficeApp, 'Quit') | Out-Null } catch { }
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($script:OfficeApp) } catch { }
        $script:OfficeApp = $null
        $script:OfficeProgId = ''
        [GC]::Collect()
    }
    # Quit 可能被"是否保存"的模态框卡住。这时才需要强杀 —— 但**只杀我们自己起的那个 pid**。
    if ($script:OfficeOwnedPids.Count -gt 0) {
        Start-Sleep -Milliseconds 300
        Stop-OfficeProcessId $script:OfficeOwnedPids
        $script:OfficeOwnedPids = @()
    }
}

function Get-OfficeApp([string]$ProgId) {
    $now = Get-Date
    if ($script:OfficeApp -and $script:OfficeProgId -eq $ProgId) {
        if (($now - $script:OfficeLastUse).TotalSeconds -lt 180) {
            $script:OfficeLastUse = $now
            return $script:OfficeApp
        }
        Close-OfficeApp
    }
    if ($script:OfficeApp) { Close-OfficeApp }
    $before = Get-OfficeProcessIds
    $app = New-Object -ComObject $ProgId
    $script:OfficeApp = $app
    $script:OfficeProgId = $ProgId
    $script:OfficeLastUse = $now
    $script:OfficeOwnedPids = Get-NewOfficeProcessIds $before
    return $app
}

# ============================ 打印实现 ============================

function Invoke-SumatraPrint {
    param([string]$Path, [string]$Printer, [int]$Copies, [string]$Pages = '')
    $engine = $script:SumatraPath
    if (-not (Test-Path $engine)) { throw '未找到 PDF 静默打印引擎 bin\SumatraPDF.exe' }

    # -print-settings 是逗号分隔的令牌列表，页范围（如 1-3,5）本身就是令牌之一
    $settings = @()
    if (-not [string]::IsNullOrWhiteSpace($Pages)) { $settings += $Pages.Trim() }
    if ($Copies -gt 1) { $settings += ($Copies.ToString([System.Globalization.CultureInfo]::InvariantCulture) + 'x') }

    $sb = New-Object System.Text.StringBuilder
    if ($Printer) { [void]$sb.Append('-print-to "' + $Printer + '" ') } else { [void]$sb.Append('-print-to-default ') }
    [void]$sb.Append('-silent ')
    if ($settings.Count -gt 0) { [void]$sb.Append('-print-settings "' + ($settings -join ',') + '" ') }
    [void]$sb.Append('-exit-when-done ')
    [void]$sb.Append('"' + $Path + '"')
    $argString = $sb.ToString()
    $proc = Start-Process -FilePath $engine -ArgumentList $argString -PassThru -WindowStyle Hidden
    if (-not $proc.WaitForExit(300000)) { try { $proc.Kill() } catch { }; throw 'PDF 打印超时' }
    if ($proc.ExitCode -ne 0) { throw "PDF 打印引擎返回错误码 $($proc.ExitCode)（请检查打印机是否就绪、缺纸或脱机）" }
    return 'SumatraPDF 静默打印'
}

function Set-OfficePrinter($App, [string]$Printer) {
    if ([string]::IsNullOrWhiteSpace($Printer)) { return }
    $port = Get-PrinterPortName $Printer
    $lastErr = ''
    foreach ($cand in @("$Printer on $port", $Printer)) {
        if ([string]::IsNullOrWhiteSpace($cand) -or $cand -eq ' on ') { continue }
        try { [Disp]::Set($App, 'ActivePrinter', $cand) } catch { $lastErr = $_.Exception.Message; continue }
        # 回读**必须**成功且匹配才算设置成功。
        # 以前这里是"读不到不算失败"，于是一个根本没生效的打印机会被当成成功，
        # 任务就落到别的地方去了 —— 相当于把校验写成了装饰。
        $back = ''
        try { $back = [string][Disp]::Get($App, 'ActivePrinter') } catch { continue }
        if ([string]::IsNullOrWhiteSpace($back)) { continue }
        if ($back -ne $Printer -and $back -notlike ($Printer + '*')) { continue }
        return
    }
    if ($lastErr) { throw "无法把打印目标切到「$Printer」（$lastErr）" }
    throw "无法把打印目标切到「$Printer」（回读校验未通过，可能这台打印机不存在或不可用）"
}

# 打开不可信文档之前，先把 Office 的"自动化安全"收紧。
# 以前只传了 ReadOnly —— 那只挡住"回写到原文件"，对文档里携带的宏、
# 外部链接、打开时的字段更新**一概不设防**，而工具本身也不弹任何提示。
# 这里显式：关闭全部宏（msoAutomationSecurityForceDisable = 3）、关闭链接更新、
# 关闭转换确认与警告弹窗。WPS 上没有这些属性时逐项 try/catch 忽略，不影响正常打印。
function Set-OfficeSafety($App) {
    if ($null -eq $App) { return }
    try { [Disp]::Set($App, 'AutomationSecurity', 3) } catch { }
    try { [Disp]::Set($App, 'DisplayAlerts', 0) } catch { }
    try {
        $opts = [Disp]::Get($App, 'Options')
        if ($opts) {
            try { [Disp]::Set($opts, 'UpdateLinksAtOpen', $false) } catch { }
            try { [Disp]::Set($opts, 'ConfirmConversions', $false) } catch { }
        }
    } catch { }
}

function Invoke-OfficePrint {
    param([string]$ProgId, [string]$Path, [string]$Printer, [int]$Copies, [string]$Pages = '')
    $label = if ($ProgId -like '*KWPS*') { 'WPS' } else { 'Word' }
    $app = Get-OfficeApp $ProgId
    $docs = [Disp]::Get($app, 'Documents')
    if ($null -eq $docs) { throw "$label 组件没有就绪" }

    Set-OfficeSafety $app
    Set-OfficePrinter $app $Printer

    $doc = $null
    try {
        # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
        # 第 4 个参数把文档挡在"最近使用的文件"列表之外，避免打印一份简历就往
        # 操作者的 Word 历史里塞一条记录（简历路径属 PII）。
        $doc = [Disp]::Call($docs, 'Open', $Path, $false, $true, $false)
        if ($null -eq $doc) { throw "$label 打不开这个文件" }
        if (-not [string]::IsNullOrWhiteSpace($Pages)) {
            # PrintOut(Background, Append, Range=4(wdPrintRangeOfPages), OutputFileName,
            #         From, To, Item, Copies, Pages)
            [Disp]::Call($doc, 'PrintOut', $false, $false, 4, '', 0, 0, 0, $Copies, $Pages.Trim()) | Out-Null
        } elseif ($Copies -gt 1) {
            [Disp]::Call($doc, 'PrintOut', $false, $false, 0, '', 0, 0, 0, $Copies) | Out-Null
        } else {
            [Disp]::Call($doc, 'PrintOut', $false) | Out-Null
        }
    } finally {
        if ($doc) { try { [Disp]::Call($doc, 'Close', 0) | Out-Null } catch { } }
    }
    if ([string]::IsNullOrWhiteSpace($Pages)) { return "$label 已送出 $Copies 份" }
    return "$label 已送出 $Copies 份（第 $Pages 页）"
}

function Convert-OfficeToPdf {
    param([string]$ProgId, [string]$Path)
    $out = Join-Path $TempDir ([guid]::NewGuid().ToString('N') + '.pdf')
    $app = Get-OfficeApp $ProgId
    $docs = [Disp]::Get($app, 'Documents')
    if ($null -eq $docs) { throw '组件没有就绪' }
    Set-OfficeSafety $app
    $doc = $null
    try {
        $doc = [Disp]::Call($docs, 'Open', $Path, $false, $true, $false)
        if ($null -eq $doc) { throw '打不开这个文件' }
        [Disp]::Call($doc, 'ExportAsFixedFormat', $out, 17) | Out-Null
    } finally {
        if ($doc) { try { [Disp]::Call($doc, 'Close', 0) | Out-Null } catch { } }
    }
    if (-not (Test-Path $out)) { throw '转换 PDF 失败' }
    return $out
}

# ============================ PDF 加深打印（浅色 / 扫描件） ============================

$script:WinRtReady = $null
$script:WinRtAsTaskOp = $null
$script:WinRtAsTaskAct = $null

function Initialize-WinRt {
    if ($null -ne $script:WinRtReady) { return $script:WinRtReady }
    $script:WinRtReady = $false
    if (-not $script:EnhReady) { return $false }
    try {
        Add-Type -AssemblyName System.Drawing -ErrorAction Stop
        Add-Type -AssemblyName System.Runtime.WindowsRuntime -ErrorAction Stop
        $null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
        $null = [Windows.Storage.StorageFile, Windows.Storage, ContentType = WindowsRuntime]
        $null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
        $exts = [System.WindowsRuntimeSystemExtensions].GetMethods()
        $script:WinRtAsTaskOp = ($exts | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' })[0]
        $script:WinRtAsTaskAct = ($exts | Where-Object { $_.Name -eq 'AsTask' -and $_.GetParameters().Count -eq 1 -and $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' })[0]
        if ($script:WinRtAsTaskOp -and $script:WinRtAsTaskAct) { $script:WinRtReady = $true }
    } catch { $script:WinRtReady = $false }
    return $script:WinRtReady
}

function Await-WinRtOp($op, $type) {
    $m = $script:WinRtAsTaskOp.MakeGenericMethod($type)
    $task = $m.Invoke($null, @($op))
    $task.Wait(-1) | Out-Null
    return $task.Result
}
function Await-WinRtAction($act) {
    $task = $script:WinRtAsTaskAct.Invoke($null, @($act))
    $task.Wait(-1) | Out-Null
}

function Test-ImagePdf([string]$Path) {
    # 通篇找不到字体对象 → 基本可以断定是"扫描件/图片版" PDF
    try {
        $fs = [System.IO.File]::OpenRead($Path)
        try {
            $len = [int][Math]::Min(1048576, $fs.Length)
            $buf = New-Object byte[] $len
            [void]$fs.Read($buf, 0, $len)
            $txt = [System.Text.Encoding]::GetEncoding(28591).GetString($buf)
            if ($txt -match '/Type\s*/Font' -or $txt -match '/BaseFont' -or $txt -match '/FontDescriptor') { return $false }
            return $true
        } finally { $fs.Dispose() }
    } catch { return $false }
}

function ConvertTo-EnhancedPdf {
    param([string]$Path, [double]$Gamma = 2.0, [int[]]$PageIndexes = $null)
    if (-not (Initialize-WinRt)) { throw '本机不支持内置 PDF 渲染（需要 Windows 10 及以上）' }
    $out = Join-Path $TempDir ([guid]::NewGuid().ToString('N') + '_enh.pdf')
    $doc = $null
    try {
        $file = Await-WinRtOp ([Windows.Storage.StorageFile]::GetFileFromPathAsync($Path)) ([Windows.Storage.StorageFile])
        $doc = Await-WinRtOp ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])

        # 没指定就全打；指定了就只渲染这几页（0 基），并保持用户给的顺序
        $want = if ($null -ne $PageIndexes -and $PageIndexes.Count -gt 0) { $PageIndexes } else { @(0..([int]$doc.PageCount - 1)) }

        $builder = New-Object EnhPdf+Builder
        foreach ($i in $want) {
            if ($i -lt 0 -or $i -ge $doc.PageCount) { continue }
            $page = $doc.GetPage([uint32]$i)
            $pwpt = [double]$page.Size.Width * 0.75     # WinRT 给的是 DIP，换算成 pt
            $phpt = [double]$page.Size.Height * 0.75
            $reqW = [int][math]::Round($pwpt / 72.0 * 200)   # 实测会被放大到约 300 DPI
            if ($reqW -lt 800) { $reqW = 800 }
            $opts = New-Object Windows.Data.Pdf.PdfPageRenderOptions
            $opts.DestinationWidth = [uint32]$reqW
            $opts.DestinationHeight = [uint32][math]::Round($reqW * $phpt / $pwpt)
            $stream = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
            Await-WinRtAction ($page.RenderToStreamAsync($stream, $opts))
            $net = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($stream)
            $ms = New-Object System.IO.MemoryStream
            $net.CopyTo($ms)
            $bmp = $null
            try {
                $bmp = [System.Drawing.Bitmap]::FromStream((New-Object System.IO.MemoryStream(, $ms.ToArray())))
                $rgb = [EnhPdf]::ToRgbGamma($bmp, $Gamma)
                $builder.AddPage($rgb, $bmp.Width, $bmp.Height, $pwpt, $phpt)
            } finally {
                if ($bmp) { $bmp.Dispose() }
                $net.Dispose(); $ms.Dispose(); $stream.Dispose(); $page.Dispose()
            }
        }
        if ($builder.PageCount -eq 0) { throw '这个 PDF 没有可打印的页面' }
        $builder.Save($out)
    } catch {
        Remove-Item $out -Force -ErrorAction SilentlyContinue
        throw
    } finally {
        if ($doc) { try { $doc.Dispose() } catch { } }
    }
    return $out
}

function Invoke-PdfEnhancedPrint {
    param([string]$Path, [string]$Printer, [int]$Copies, [double]$Gamma, [int[]]$PageIndexes = $null)
    $tmp = ConvertTo-EnhancedPdf -Path $Path -Gamma $Gamma -PageIndexes $PageIndexes
    try {
        # 选页已经在重建 PDF 时做掉了，这里整份打即可，不能再传页范围
        [void](Invoke-SumatraPrint -Path $tmp -Printer $Printer -Copies $Copies)
    } finally {
        Remove-Item $tmp -Force -ErrorAction SilentlyContinue
    }
    return '加深处理后再静默打印'
}

# ============================ 页数统计 / 页码范围 ============================

$script:PageCountCache = @{}

function Get-PdfPageCount([string]$Path) {
    if (-not (Initialize-WinRt)) { return -1 }
    $doc = $null
    try {
        $file = Await-WinRtOp ([Windows.Storage.StorageFile]::GetFileFromPathAsync($Path)) ([Windows.Storage.StorageFile])
        $doc = Await-WinRtOp ([Windows.Data.Pdf.PdfDocument]::LoadFromFileAsync($file)) ([Windows.Data.Pdf.PdfDocument])
        return [int]$doc.PageCount
    } catch { return -1 }
    finally { if ($doc) { try { $doc.Dispose() } catch { } } }
}

function Get-OfficePageCount([string]$ProgId, [string]$Path) {
    $app = Get-OfficeApp $ProgId
    $docs = [Disp]::Get($app, 'Documents')
    if ($null -eq $docs) { return -1 }
    Set-OfficeSafety $app
    $doc = $null
    try {
        $doc = [Disp]::Call($docs, 'Open', $Path, $false, $true, $false)
        if ($null -eq $doc) { return -1 }
        return [int][Disp]::Call($doc, 'ComputeStatistics', 2)   # wdStatisticPages
    } catch { return -1 }
    finally { if ($doc) { try { [Disp]::Call($doc, 'Close', 0) | Out-Null } catch { } } }
}

function Get-FilePageCount([string]$Path) {
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path)) { return -1 }
    $key = $Path
    try { $key = $Path + '|' + (Get-Item -LiteralPath $Path).LastWriteTimeUtc.Ticks } catch { }
    if ($script:PageCountCache.ContainsKey($key)) { return [int]$script:PageCountCache[$key] }

    $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    $n = -1
    if ($ext -eq '.pdf') {
        $n = Get-PdfPageCount $Path
    } else {
        foreach ($prog in $script:OfficeOrder) {
            if ($prog -eq 'Word.Application' -and -not $script:EngineStatus.word) { continue }
            if ($prog -like '*KWPS*' -and -not $script:EngineStatus.wps) { continue }
            $n = Get-OfficePageCount $prog $Path
            if ($n -gt 0) { break }
        }
    }
    $script:PageCountCache[$key] = $n
    # 页数缓存也要封顶：以前只增不减，长年累月翻不同文件夹会一直涨。
    # 成功与失败的结果都会缓存，所以这里只做容量控制，不丢语义。
    if ($script:PageCountCache.Count -gt 2000) { $script:PageCountCache = @{} }
    return $n
}

# 页码的合法上界。真实文档不可能有这么多页，但它能挡住
# `9999999999` 这种"强转 [int] 直接抛原始转换异常"的输入，
# 让用户看到友好的范围提示而不是一个 .NET 报错。
$script:MaxPageNumber = 100000

function Test-PageSpec([string]$Spec, [int]$MaxPages) {
    # 返回 '' 表示合法，否则返回中文错误说明
    if ([string]::IsNullOrWhiteSpace($Spec)) { return '' }
    foreach ($raw in ($Spec -split ',')) {
        $p = $raw.Trim()
        if ($p -eq '') { return '页码范围里有空项（是不是多打了一个逗号？）' }
        if ($p -match '^(?i)(even|odd|last)$') { continue }
        if ($p -match '^-\d+$') {
            if ($p.Length -gt 8) { return "页码数字太大：$p" }
            continue
        }
        if ($p -match '^\d+$') {
            if ($p.Length -gt 9) { return "页码数字太大：$p（页码不可能这么大）" }
            $n = 0
            if (-not [long]::TryParse($p, [ref]$n)) { return "页码数字太大：$p" }
            if ($n -lt 1) { return "页码要从 1 开始：$p" }
            if ($n -gt $script:MaxPageNumber) { return "页码太大：$p（最大 $($script:MaxPageNumber)）" }
            if ($MaxPages -gt 0 -and $n -gt $MaxPages) { return "第 $n 页超出范围（这份文件共 $MaxPages 页）" }
            continue
        }
        if ($p -match '^(\d+)\s*-\s*(\d+)$') {
            $a = 0; $b = 0
            if (-not [long]::TryParse($Matches[1], [ref]$a) -or -not [long]::TryParse($Matches[2], [ref]$b)) { return "页码数字太大：$p" }
            if ($a -lt 1 -or $b -lt 1) { return "页码要从 1 开始：$p" }
            if ($a -gt $script:MaxPageNumber -or $b -gt $script:MaxPageNumber) { return "页码太大：$p（最大 $($script:MaxPageNumber)）" }
            if ($MaxPages -gt 0 -and ([Math]::Max($a, $b) -gt $MaxPages)) { return "第 $p 页超出范围（这份文件共 $MaxPages 页）" }
            continue
        }
        return "看不懂的页码：「$p」，正确写法像 1-3,5,8"
    }
    return ''
}

# 0 基页索引 -> "1-3,5,8-10" 这种紧凑写法（SumatraPDF 和 Word 都认）
function ConvertTo-PageSpec([int[]]$Indexes) {
    if ($null -eq $Indexes -or $Indexes.Count -eq 0) { return '' }
    $nums = @($Indexes | Sort-Object -Unique)
    $parts = @()
    $start = $nums[0]; $prev = $nums[0]
    for ($i = 1; $i -le $nums.Count; $i++) {
        $cur = if ($i -lt $nums.Count) { [int]$nums[$i] } else { [int]::MinValue }
        if ($cur -eq $prev + 1) { $prev = $cur; continue }
        $a = $start + 1; $b = $prev + 1
        if ($a -eq $b) { $parts += "$a" } else { $parts += "$a-$b" }
        $start = $cur; $prev = $cur
    }
    return ($parts -join ',')
}

# "1-3,5" / even / odd / last / -1 -> 0 基页索引数组（去重升序）；空串 = 全部页
function ConvertTo-PageIndexList([string]$Spec, [int]$PageCount) {
    if ($PageCount -le 0) { return @() }
    if ([string]::IsNullOrWhiteSpace($Spec)) { return @(0..($PageCount - 1)) }
    $set = New-Object 'System.Collections.Generic.SortedSet[int]'
    foreach ($raw in ($Spec -split ',')) {
        $p = $raw.Trim()
        if ($p -eq '') { continue }
        if ($p -match '^(?i)odd$') { for ($i = 1; $i -le $PageCount; $i += 2) { [void]$set.Add($i - 1) }; continue }
        if ($p -match '^(?i)even$') { for ($i = 2; $i -le $PageCount; $i += 2) { [void]$set.Add($i - 1) }; continue }
        if ($p -match '^(?i)last$') { [void]$set.Add($PageCount - 1); continue }
        if ($p -match '^-\d+$') {
            $n = [int]$p.Substring(1)
            if ($n -ge 1 -and $n -le $PageCount) { [void]$set.Add($PageCount - $n) }
            continue
        }
        if ($p -match '^\d+$') {
            $n = [int]$p
            if ($n -ge 1 -and $n -le $PageCount) { [void]$set.Add($n - 1) }
            continue
        }
        if ($p -match '^(\d+)\s*-\s*(\d+)$') {
            $a = [int]$Matches[1]; $b = [int]$Matches[2]
            $lo = [Math]::Min($a, $b); $hi = [Math]::Max($a, $b)
            for ($i = $lo; $i -le $hi; $i++) {
                if ($i -ge 1 -and $i -le $PageCount) { [void]$set.Add($i - 1) }
            }
            continue
        }
    }
    return @($set)
}

function Invoke-PrintFile {
    param(
        [string]$Path,
        [string]$Printer,
        [int]$Copies = 1,
        [bool]$DryRun = $false,
        [string]$Enhance = 'auto',
        [string]$Pages = '',
        [bool]$ClampPages = $false
    )
    $result = [ordered]@{
        path = $Path; ok = $false; method = ''; detail = ''
        time = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    }

    if ([string]::IsNullOrWhiteSpace($Path)) { $result.detail = '没有拿到文件路径（请重新扫描后再打印）'; return $result }
    if (-not (Test-Path -LiteralPath $Path)) { $result.detail = '文件不存在'; return $result }
    if (-not $script:KnownFiles.ContainsKey($Path)) { $result.detail = '这个文件不在本次扫描结果里，请重新扫描后再打印'; return $result }
    $ext = [System.IO.Path]::GetExtension($Path).ToLowerInvariant()
    if ($SupportedExt -notcontains $ext) { $result.detail = "不支持的格式 $ext"; return $result }
    if ($Copies -lt 1) { $Copies = 1 }
    if ($Printer -and -not (Test-PrinterExists $Printer)) { $result.detail = "找不到打印机「$Printer」"; return $result }

    # 页范围：先按"已知页数"校验一遍，防止用户手输 999 页把打印任务搞乱
    $Pages = ([string]$Pages).Trim()
    $pageCount = -1
    $pageCountKnown = $false
    if ($Pages -ne '') {
        $pageCount = Get-FilePageCount $Path
        $pageCountKnown = ($pageCount -gt 0)
        if ($ClampPages -and $pageCountKnown) {
            # 批量套用页码时走这条：某个文件页数更少，越界的那几页直接丢掉，
            # 而不是让整份文件失败（比如「都只打 1-3 页」遇到一份只有 2 页的简历）
            $idx = ConvertTo-PageIndexList $Pages $pageCount
            if ($idx.Count -eq 0) {
                $result.detail = "第 $Pages 页在这份文件里一页都不存在（它只有 $pageCount 页）"
                return $result
            }
            $Pages = ConvertTo-PageSpec $idx
        } else {
            # $MaxPages = -1 时只做**语法**校验 + 绝对上界校验，不做"是否超出这份文件"的校验。
            # 这是"页数拿不到"时的显式第三种模式：取值范围无法核对，但输入形态仍然受控。
            $bad = Test-PageSpec $Pages $pageCount
            if ($bad -ne '') { $result.detail = $bad; return $result }
            # 规范化：把 "5,1-3" 这类整理成 SumatraPDF/Word 都认的紧凑写法。
            # 只有页数已知时才重排/去重；页数未知时原样交给引擎，
            # 并把这件事明确写进结果里，而不是静默当成功。
            if ($pageCountKnown) {
                $idx = ConvertTo-PageIndexList $Pages $pageCount
                if ($idx.Count -eq 0) { $result.detail = '选中的页码在这份文件里一页都不存在'; return $result }
                $Pages = ConvertTo-PageSpec $idx
            } elseif ($ClampPages) {
                # 批量套用 + 页数未知：无法裁剪，但也不能假装成功
                $result.detail = "拿不到这份文件的页数，没法按「$Pages」裁剪。请对这份单独用「选页」确认页数后再打。"
                return $result
            }
        }
    }
    $result['pages'] = $Pages
    $result['pageCount'] = $pageCount
    $result['pageCountKnown'] = $pageCountKnown

    $usable = @()
    foreach ($prog in $script:OfficeOrder) {
        if ($prog -eq 'Word.Application' -and -not $script:EngineStatus.word) { continue }
        if ($prog -like '*KWPS*' -and -not $script:EngineStatus.wps) { continue }
        $usable += $prog
    }

    $plan = @()
    if ($ext -eq '.pdf') {
        # 决定要不要走"加深"通道
        $useEnh = $false
        $gamma = 2.0
        switch ($Enhance) {
            'normal' { $useEnh = $false }
            'dark' { $useEnh = $true; $gamma = 1.8 }
            'darker' { $useEnh = $true; $gamma = 2.6 }
            default { if (Test-ImagePdf $Path) { $useEnh = $true; $gamma = 2.0 } }
        }
        # 注意：「加深」通道是自己渲染 + 重建 PDF，最终仍然要靠 SumatraPDF 出纸，
        # 所以 SumatraPDF 不在时不能把它算作可用，否则试运行会给出假阳性。
        if ($useEnh -and $script:EngineStatus.sumatra -and (Initialize-WinRt)) {
            $plan += ('enh:' + $gamma.ToString([System.Globalization.CultureInfo]::InvariantCulture))
        }
        if ($script:EngineStatus.sumatra) { $plan += 'sumatra' }
        foreach ($prog in $usable) { $plan += ('office:' + $prog) }
    } else {
        foreach ($prog in $usable) { $plan += ('office:' + $prog) }
        foreach ($prog in $usable) { $plan += ('convert:' + $prog) }
    }

    if ($plan.Count -eq 0) {
        $result.detail = '本机没有可用的打印引擎（PDF 引擎缺失，也没有可用的 Word/WPS）。请先点右上角「重新检测」。'
        return $result
    }

    $methodName = @{
        'sumatra'                   = 'SumatraPDF'
        'office:Word.Application'   = 'Word'
        'office:KWPS.Application'   = 'WPS'
        'convert:Word.Application'  = 'Word转PDF'
        'convert:KWPS.Application'  = 'WPS转PDF'
    }

    if ($DryRun) {
        $result.ok = $true
        if ($plan[0] -like 'enh:*') { $result.method = '加深打印' }
        else { $result.method = $methodName[$plan[0]] }
        $scope = if ($Pages -eq '') { '全部页' } else { "第 $Pages 页" }
        $result.detail = "试运行（未真正打印，$scope，目标：$(if ($Printer) { $Printer } else { '默认打印机' })）"
        return $result
    }

    $errors = @()
    foreach ($step in $plan) {
        try {
            if ($step -eq 'sumatra') {
                $result.detail = Invoke-SumatraPrint -Path $Path -Printer $Printer -Copies $Copies -Pages $Pages
                $result.method = 'SumatraPDF'
            }
            elseif ($step -like 'enh:*') {
                $g = [double]::Parse($step.Substring(4), [System.Globalization.CultureInfo]::InvariantCulture)
                # 加深通道要自己重建 PDF，所以把"选哪几页"换算成 0 基索引传进去
                $idx = $null
                if ($Pages -ne '' -and $pageCount -gt 0) { $idx = ConvertTo-PageIndexList $Pages $pageCount }
                $result.detail = Invoke-PdfEnhancedPrint -Path $Path -Printer $Printer -Copies $Copies -Gamma $g -PageIndexes $idx
                $result.method = '加深打印'
            }
            elseif ($step -like 'office:*') {
                $prog = $step.Substring(7)
                $label = if ($prog -like '*KWPS*') { 'WPS' } else { 'Word' }
                $result.detail = Invoke-OfficePrint -ProgId $prog -Path $Path -Printer $Printer -Copies $Copies -Pages $Pages
                $result.method = $label
                $script:OfficeOrder = @($prog) + @($script:OfficeOrder | Where-Object { $_ -ne $prog })
            }
            elseif ($step -like 'convert:*') {
                $prog = $step.Substring(8)
                $tmpPdf = Convert-OfficeToPdf -ProgId $prog -Path $Path
                try {
                    # 转出来的 PDF 页序与原文档一致，页范围可以直接沿用
                    $result.detail = '转成 PDF 后 ' + (Invoke-SumatraPrint -Path $tmpPdf -Printer $Printer -Copies $Copies -Pages $Pages)
                    $result.method = '转PDF后静默打印'
                } finally { Remove-Item $tmpPdf -Force -ErrorAction SilentlyContinue }
            }
            $result.ok = $true
            break
        } catch {
            $errors += "$($methodName[$step]): $($_.Exception.Message)"
            Log ("$([System.IO.Path]::GetFileName($Path)) -> $($methodName[$step]) 失败：$($_.Exception.Message)") 'warn'
            if ($step -like 'office:*' -or $step -like 'convert:*') { Close-OfficeApp }
        }
    }

    if (-not $result.ok) {
        $result.detail = ($errors -join '　|　')
        if ([string]::IsNullOrWhiteSpace($result.detail)) { $result.detail = '所有打印方式都失败了' }
    } else {
        $script:Printed[$Path] = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
    }
    return $result
}

# ============================ 文件夹浏览 / 扫描 ============================

function Get-DirListing([string]$Path) {
    # 快捷入口只在"没有指定路径"（也就是停在根）时才返回。
    # 以前无论调用者查哪个目录都会无条件回带桌面/文档/下载的绝对路径，
    # 等于把用户配置文件的形态白送给任何人。
    $shortcuts = @()
    if ([string]::IsNullOrWhiteSpace($Path)) {
        foreach ($pair in @(
                @('桌面', [Environment]::GetFolderPath('Desktop')),
                @('我的文档', [Environment]::GetFolderPath('MyDocuments')),
                @('下载', (Join-Path $env:USERPROFILE 'Downloads')),
                @('此电脑', '')
            )) {
            if ($pair[1] -eq '' -or (Test-Path -LiteralPath $pair[1])) {
                $shortcuts += [ordered]@{ name = [string]$pair[0]; path = [string]$pair[1] }
            }
        }
    }

    $result = [ordered]@{
        ok = $true; path = ''; parent = ''; dirs = @(); drives = @()
        shortcuts = $shortcuts; error = ''
    }

    if ([string]::IsNullOrWhiteSpace($Path)) {
        try {
            foreach ($d in (Get-PSDrive -PSProvider FileSystem -ErrorAction SilentlyContinue | Sort-Object Name)) {
                $root = [string]$d.Root
                if ([string]::IsNullOrWhiteSpace($root)) { continue }
                $result.drives += [ordered]@{ name = ([string]$d.Name + ':'); path = $root }
            }
        } catch { }
        return $result
    }

    $full = $Path
    try { $full = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).Path }
    catch {
        $result.ok = $false
        $result.path = $Path
        $result.error = "打不开这个文件夹：$Path"
        try { $result.parent = [string](Split-Path -Parent $Path) } catch { }
        return $result
    }

    $result.path = $full
    $parent = ''
    try { $parent = [string](Split-Path -Parent $full) } catch { }
    if ($parent -eq $full) { $parent = '' }
    $result.parent = $parent

    $dirs = @()
    try {
        foreach ($d in (Get-ChildItem -LiteralPath $full -Directory -Force -ErrorAction SilentlyContinue | Sort-Object Name)) {
            $attr = $d.Attributes
            if ($attr -band [System.IO.FileAttributes]::System) { continue }
            if ($attr -band [System.IO.FileAttributes]::Hidden) { continue }
            $dirs += [ordered]@{ name = [string]$d.Name; path = [string]$d.FullName }
        }
    } catch { }
    $result.dirs = $dirs
    return $result
}

function Invoke-Scan([string]$Folder, [bool]$Recursive) {
    if ([string]::IsNullOrWhiteSpace($Folder)) { throw '请先填写文件夹路径' }
    if (-not (Test-Path -LiteralPath $Folder)) { throw "文件夹不存在：$Folder" }

    # 扫描上限。以前是无限递归：/api/scan 可以对任意存在的目录（含盘根）做无深度、
    # 无条目数的枚举，整个结果还要序列化进一个响应里 —— 单请求就能把内存和
    # 唯一那个 worker 占满。现在给出明确的深度与条目上界，超了就截断并告知。
    $maxDepth = 8
    $maxEntries = 20000
    $truncated = $false

    $items = New-Object System.Collections.Generic.List[System.IO.FileInfo]
    try {
        if ($Recursive) {
            $stack = New-Object System.Collections.Generic.Stack[object]
            $stack.Push(@($Folder, 0))
            while ($stack.Count -gt 0) {
                $cur = $stack.Pop()
                $dir = [string]$cur[0]
                $depth = [int]$cur[1]
                foreach ($f in @(Get-ChildItem -LiteralPath $dir -File -Force -ErrorAction SilentlyContinue)) { $items.Add($f) }
                if ($items.Count -ge $maxEntries) { $truncated = $true; break }
                if ($depth -ge $maxDepth) { $truncated = $true; continue }
                foreach ($d in @(Get-ChildItem -LiteralPath $dir -Directory -Force -ErrorAction SilentlyContinue)) {
                    if ($d.Attributes -band [System.IO.FileAttributes]::ReparsePoint) { continue }   # 不跟符号链接/联接走
                    $stack.Push(@($d.FullName, $depth + 1))
                }
            }
        } else {
            foreach ($f in @(Get-ChildItem -LiteralPath $Folder -File -Force -ErrorAction SilentlyContinue)) { $items.Add($f) }
        }
    } catch { }
    if ($items.Count -gt $maxEntries) { $items = $items.GetRange(0, $maxEntries); $truncated = $true }

    $files = @()
    $other = 0
    # 只允许打印"本次扫描列出来的文件"，避免这个本地服务被当成任意文件打印器
    $script:KnownFiles = @{}
    foreach ($f in $items) {
        $ext = $f.Extension.ToLowerInvariant()
        if ($SupportedExt -notcontains $ext) { $other++; continue }
        $script:KnownFiles[$f.FullName] = $true
        $files += [ordered]@{
            path    = $f.FullName
            name    = $f.Name
            ext     = $ext
            size    = $f.Length
            mtime   = $f.LastWriteTime.ToString('yyyy-MM-dd HH:mm')
            printed = $script:Printed.ContainsKey($f.FullName)
        }
    }
    $files = @($files | Sort-Object { $_.name })
    return @{
        folder = (Resolve-Path -LiteralPath $Folder).Path
        files = $files
        otherCount = $other
        truncated = $truncated
        truncateNote = $(if ($truncated) { "目录太深或文件太多，只扫描了前 $maxEntries 个条目（最深 $maxDepth 层）。如果确实有更多简历，请缩小文件夹范围。" } else { '' })
    }
}

# ============================ HTTP 服务 ============================

function Get-HeaderEndIndex([byte[]]$bytes) {
    for ($i = 3; $i -lt $bytes.Length; $i++) {
        if ($bytes[$i - 3] -eq 13 -and $bytes[$i - 2] -eq 10 -and $bytes[$i - 1] -eq 13 -and $bytes[$i] -eq 10) { return $i - 3 }
    }
    return -1
}

function Read-HttpRequest($client) {
    $stream = $client.GetStream()

    # 单次读取超时只约束"一次 Read"，挡不住"每 29 秒滴 1 字节"的连接。
    # 这个服务只有一个 worker（串行 accept），所以必须再加一个**整个请求的总时限**，
    # 否则任意一个本机连接就能把界面饿死。这里把单次超时压到 10 秒，
    # 让总时限最多晚 10 秒被发现，总上限 = 30 + 10 秒。
    $client.ReceiveTimeout = 10000
    $reqDeadline = (Get-Date).AddSeconds(30)
    $timedOut = $false

    $ms = New-Object System.IO.MemoryStream
    $buf = New-Object byte[] 8192
    $headerEnd = -1
    while ($headerEnd -lt 0) {
        if ((Get-Date) -ge $reqDeadline) { $timedOut = $true; break }
        $read = $stream.Read($buf, 0, $buf.Length)
        if ($read -le 0) { break }
        $ms.Write($buf, 0, $read)
        $all = $ms.ToArray()
        $headerEnd = Get-HeaderEndIndex $all
        if ($all.Length -gt 65536) { break }        # 请求头最大 64KB
    }
    try { $ms.Dispose() } catch { }
    $all = $ms.ToArray()
    if ($timedOut -or $headerEnd -lt 0) { return $null }

    $headerText = [System.Text.Encoding]::UTF8.GetString($all, 0, $headerEnd)
    $lines = $headerText -split "`r`n"
    $parts = $lines[0] -split ' '
    $method = $parts[0]
    $target = if ($parts.Length -gt 1) { $parts[1] } else { '/' }

    $headers = @{}
    $contentLength = 0
    for ($i = 1; $i -lt $lines.Length; $i++) {
        $ln = $lines[$i]
        $ci = $ln.IndexOf(':')
        if ($ci -le 0) { continue }
        $hn = $ln.Substring(0, $ci).Trim().ToLowerInvariant()
        $hv = $ln.Substring($ci + 1).Trim()
        $headers[$hn] = $hv
        if ($hn -eq 'content-length') { [void][int]::TryParse($hv, [ref]$contentLength) }
    }

    # 先解析出 path/query，再处理请求体。
    # 之前超限分支在 $path 赋值之前就读它，读到的是未定义变量（靠 Set-StrictMode 未启用来兜底），
    # 现在把它提到前面，任何分支返回的对象形状都一致。
    $path = $target
    $query = ''
    $qi = $target.IndexOf('?')
    if ($qi -ge 0) { $path = $target.Substring(0, $qi); $query = $target.Substring($qi + 1) }

    # 请求体上限 1MB，防止内存被撑爆
    if ($contentLength -lt 0) { $contentLength = 0 }
    if ($contentLength -gt 1048576) {
        # 太大：有界地把客户端已经发出来的数据读掉（最多 8MB，且仍受请求总时限约束），
        # 这样客户端能正常收到我们的 413，而不是被 RST；但绝不会无限读下去。
        $drained = 0
        try {
            while ($drained -lt $contentLength -and $drained -lt 8388608 -and (Get-Date) -lt $reqDeadline) {
                $n = $stream.Read($buf, 0, $buf.Length)
                if ($n -le 0) { break }
                $drained += $n
            }
        } catch { }
        return @{ tooLarge = $true; method = $method; path = $path; query = $query; body = ''; headers = $headers }
    }

    $bodyStart = $headerEnd + 4
    $bodyMs = New-Object System.IO.MemoryStream
    if ($all.Length -gt $bodyStart) { $bodyMs.Write($all, $bodyStart, $all.Length - $bodyStart) }
    while ($bodyMs.Length -lt $contentLength) {
        if ((Get-Date) -ge $reqDeadline) { break }
        $read = $stream.Read($buf, 0, $buf.Length)
        if ($read -le 0) { break }
        $bodyMs.Write($buf, 0, $read)
    }
    $body = [System.Text.Encoding]::UTF8.GetString($bodyMs.ToArray())
    try { $bodyMs.Dispose() } catch { }

    return @{ tooLarge = $false; method = $method; path = $path; query = $query; body = $body; headers = $headers }
}

# ---------------- 请求合法性校验（本地服务的安全边界） ----------------
# 这是一个"能读本地目录、还能指挥打印机"的本地服务，必须挡住：
#   1) DNS 重绑定：恶意域名解析到 127.0.0.1，浏览器就会带 Host: evil.com 来访问
#   2) 跨站请求伪造(CSRF)：别的网页偷偷 POST 过来让我们打印 / 翻目录
#   3) 方法混淆：以前 Content-Type 闸门只写在 "if method = POST" 里，
#      而分发是按路径的、完全不看方法，于是**不带 Origin 的 GET** 能把三项校验全部绕过
#      （浏览器对普通跨站 GET 不发送 Origin；目标就是 127.0.0.1，Host 也天然通过），
#      而 /api/selftest、/api/testprint、/api/reset-printed 等端点**空请求体就有破坏性默认行为**。
# 手段：
#   - Host 必须是回环地址（挡 DNS 重绑定）
#   - 有 Origin 就必须是回环来源
#   - **所有会改变状态的端点一律要求 POST**，非 POST 直接 405
#   - POST 必须是 Content-Type: application/json（HTML 表单发不出它，
#     跨域发它又会先触发预检，而本服务不响应预检、也不发任何 Access-Control-* 头）
#   - 纵深防御：现代浏览器会带 Sec-Fetch-Site，跨站一律拒绝

# 会改变状态（或有副作用）的端点：一律要求 POST。/api/info 是纯只读的，允许 GET。
$script:PostOnlyPaths = @(
    '/api/selftest', '/api/fetch-engine', '/api/list-dir', '/api/pagecount',
    '/api/scan', '/api/print-one', '/api/end-batch', '/api/restore-printer',
    '/api/testprint', '/api/open', '/api/reset-printed'
)

function Test-RequestAllowed($req) {
    $hostHdr = [string]$req.headers['host']
    if (-not [string]::IsNullOrWhiteSpace($hostHdr)) {
        $h = ($hostHdr -split ':')[0].Trim().Trim('[', ']').ToLowerInvariant()
        if ($h -ne '127.0.0.1' -and $h -ne 'localhost' -and $h -ne '::1') { return 'host' }
    }

    $origin = [string]$req.headers['origin']
    if (-not [string]::IsNullOrWhiteSpace($origin)) {
        if ($origin -notmatch '^(?i)https?://(127\.0\.0\.1|localhost|\[::1\])(:\d+)?$') { return 'origin' }
    }

    # 纵深防御：浏览器策略头。非浏览器客户端可以伪造，所以它只是额外一层，不作为唯一控制。
    $sfs = [string]$req.headers['sec-fetch-site']
    if ($sfs -eq 'cross-site') { return 'cross-site' }

    $path = [string]$req.path
    if ($path -like '/api/*') {
        $needPost = $script:PostOnlyPaths -contains $path
        if ($needPost) {
            if ($req.method -ne 'POST') { return 'method' }
            $ct = [string]$req.headers['content-type']
            if ($ct -notmatch '^(?i)\s*application/json') { return 'content-type' }
        } elseif ($req.method -ne 'GET' -and $req.method -ne 'POST') {
            return 'method'
        }
    }
    return ''
}

$script:ReasonPhrases = @{ 200 = 'OK'; 204 = 'No Content'; 400 = 'Bad Request'; 403 = 'Forbidden'; 404 = 'Not Found'; 405 = 'Method Not Allowed'; 413 = 'Payload Too Large'; 500 = 'Internal Server Error' }

function Send-Bytes($client, [int]$status, [string]$contentType, [byte[]]$body) {
    $stream = $client.GetStream()
    $reason = $script:ReasonPhrases[$status]
    if (-not $reason) { $reason = 'OK' }
    $head = "HTTP/1.1 $status $reason`r`nContent-Type: $contentType`r`nContent-Length: $($body.Length)`r`nCache-Control: no-store`r`nX-Content-Type-Options: nosniff`r`nConnection: close`r`n`r`n"
    $headBytes = [System.Text.Encoding]::ASCII.GetBytes($head)
    $stream.Write($headBytes, 0, $headBytes.Length)
    if ($body.Length -gt 0) { $stream.Write($body, 0, $body.Length) }
    $stream.Flush()
}

function Send-Json($client, $obj, [int]$status = 200) {
    $json = $obj | ConvertTo-Json -Depth 8 -Compress
    Send-Bytes $client $status 'application/json; charset=utf-8' ([System.Text.Encoding]::UTF8.GetBytes($json))
}

$script:MimeTypes = @{
    '.html' = 'text/html; charset=utf-8'; '.htm' = 'text/html; charset=utf-8'
    '.js'   = 'application/javascript; charset=utf-8'; '.css' = 'text/css; charset=utf-8'
    '.json' = 'application/json; charset=utf-8'; '.svg' = 'image/svg+xml'
    '.png'  = 'image/png'; '.ico' = 'image/x-icon'; '.woff2' = 'font/woff2'
}

function Send-Static($client, [string]$urlPath) {
    if ([string]::IsNullOrWhiteSpace($urlPath) -or $urlPath -eq '/') { $urlPath = '/index.html' }
    $rel = $urlPath.TrimStart('/').Replace('/', '\')
    if ([string]::IsNullOrWhiteSpace($rel) -or $rel -match '\.\.') {
        Send-Bytes $client 404 'text/plain; charset=utf-8' ([System.Text.Encoding]::UTF8.GetBytes('Not Found')); return
    }
    $full = ''
    try {
        $full = [System.IO.Path]::GetFullPath((Join-Path $WebDir $rel))
        $rootFull = [System.IO.Path]::GetFullPath($WebDir)
        # 解析后必须仍在 web 目录内，挡掉 ..\..\ 与绝对路径
        if (-not $full.StartsWith($rootFull + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) {
            Send-Bytes $client 404 'text/plain; charset=utf-8' ([System.Text.Encoding]::UTF8.GetBytes('Not Found')); return
        }
    } catch {
        Send-Bytes $client 404 'text/plain; charset=utf-8' ([System.Text.Encoding]::UTF8.GetBytes('Not Found')); return
    }
    if (-not (Test-Path -LiteralPath $full -PathType Leaf)) {
        Send-Bytes $client 404 'text/plain; charset=utf-8' ([System.Text.Encoding]::UTF8.GetBytes('Not Found')); return
    }
    # 前缀比较只作用在**字符串**上，不解析重解析点。
    # 所以 web\ 里如果被放进一个符号链接/联接，它的目标在 web\ 之外也照样会被下发。
    # 这里补一次实物检查：整条路径上任何一段带 ReparsePoint 属性就拒绝。
    try {
        $probe = Get-Item -LiteralPath $full -Force -ErrorAction Stop
        if ($probe.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
            Send-Bytes $client 404 'text/plain; charset=utf-8' ([System.Text.Encoding]::UTF8.GetBytes('Not Found')); return
        }
    } catch {
        Send-Bytes $client 404 'text/plain; charset=utf-8' ([System.Text.Encoding]::UTF8.GetBytes('Not Found')); return
    }
    $ext = [System.IO.Path]::GetExtension($full).ToLowerInvariant()
    # 只下发已知的界面资源类型。以前任何扩展名都会下发，未知类型回落
    # application/octet-stream —— 也就是 web\ 里出现任何文件都会变成可下载项。
    $mime = $script:MimeTypes[$ext]
    if (-not $mime) {
        Send-Bytes $client 404 'text/plain; charset=utf-8' ([System.Text.Encoding]::UTF8.GetBytes('Not Found')); return
    }
    Send-Bytes $client 200 $mime ([System.IO.File]::ReadAllBytes($full))
}

function Get-BodyJson($req) {
    if ([string]::IsNullOrWhiteSpace($req.body)) { return $null }
    try { return ($req.body | ConvertFrom-Json) } catch { return $null }
}

function Handle-Api($client, $req) {
    $p = $req.path

    if ($p -eq '/api/info') {
        $pl = Get-PrinterList
        $suggested = Get-SuggestedPrinter $pl
        $sugItem = @($pl.printers | Where-Object { $_.name -eq $suggested } | Select-Object -First 1)
        # 拷到别的电脑后老路径可能不存在，过滤掉，免得一进页面就报错
        $lastFolder = [string]$script:Config.lastFolder
        if ($lastFolder -and -not (Test-Path -LiteralPath $lastFolder)) { $lastFolder = '' }
        $recent = @()
        foreach ($rf in @($script:Config.recentFolders)) {
            if ($rf -and (Test-Path -LiteralPath $rf)) { $recent += $rf }
        }
        Send-Json $client ([ordered]@{
                printers         = $pl.printers
                defaultPrinter   = $pl.defaultPrinter
                engines          = (Get-EngineReport)
                enginePath       = $script:SumatraPath
                engineMoved      = $script:SumatraRelocated
                engineMissing    = (-not $script:EngineStatus.sumatra)
                engineVersion    = $script:EngineVersion
                lastFolder       = $lastFolder
                recentFolders    = $recent
                printer          = $suggested
                printerIsVirtual = ($sugItem.Count -gt 0 -and $sugItem[0].kind -eq 'virtual')
                copies           = [int]$script:Config.copies
                recursive        = [bool]$script:Config.recursive
                enhance          = [string]$script:Config.enhance
                enhanceAvailable = (Initialize-WinRt)
                printedCount     = $script:Printed.Count
                root             = $Root
            })
        return
    }

    if ($p -eq '/api/selftest') {
        Send-Json $client ([ordered]@{ ok = $true; engines = (Invoke-EngineSelfTest) })
        return
    }

    if ($p -eq '/api/fetch-engine') {
        try {
            $f = Install-PdfEngine
            Send-Json $client ([ordered]@{
                    ok = $true
                    detail = ('PDF 引擎已安装：' + $f)
                    engines = (Get-EngineReport)
                })
        } catch {
            Send-Json $client ([ordered]@{ ok = $false; detail = ('' + $_.Exception.Message) })
        }
        return
    }

    if ($p -eq '/api/list-dir') {
        $b = Get-BodyJson $req
        Send-Json $client (Get-DirListing ([string](Get-Prop $b 'path' '')))
        return
    }

    if ($p -eq '/api/pagecount') {
        $b = Get-BodyJson $req
        $path = [string](Get-Prop $b 'path' '')
        if ([string]::IsNullOrWhiteSpace($path) -or -not $script:KnownFiles.ContainsKey($path)) {
            Send-Json $client ([ordered]@{ ok = $false; count = -1; detail = '这个文件不在本次扫描结果里' })
            return
        }
        $n = Get-FilePageCount $path
        $result = [ordered]@{
            ok = ($n -gt 0)
            count = $n
            detail = if ($n -gt 0) { '' } else { '拿不到页数（可能是加密/损坏的文件，或本机缺少对应的阅读引擎）' }
        }
        Send-Json $client $result
        return
    }

    if ($p -eq '/api/scan') {
        $b = Get-BodyJson $req
        $folder = [string](Get-Prop $b 'folder' '')
        $rec = [bool](Get-Prop $b 'recursive' $true)
        $script:Config.lastFolder = $folder
        $script:Config.recursive = $rec
        $scanResult = Invoke-Scan -Folder $folder -Recursive $rec
        Add-RecentFolder $scanResult.folder
        Save-Config
        Send-Json $client $scanResult
        return
    }

    if ($p -eq '/api/print-one') {
        $b = Get-BodyJson $req
        $path = [string](Get-Prop $b 'path' '')
        $printer = [string](Get-Prop $b 'printer' '')
        $copies = [int](Get-Prop $b 'copies' 1)
        $dry = [bool](Get-Prop $b 'dryRun' $false)
        $enh = [string](Get-Prop $b 'enhance' 'auto')
        $pages = [string](Get-Prop $b 'pages' '')
        $clamp = [bool](Get-Prop $b 'pagesClamp' $false)
        if ([string]::IsNullOrWhiteSpace($enh)) { $enh = 'auto' }
        $script:Config.printer = $printer
        $script:Config.copies = $copies
        $script:Config.enhance = $enh
        Save-Config
        $r = Invoke-PrintFile -Path $path -Printer $printer -Copies $copies -DryRun $dry -Enhance $enh -Pages $pages -ClampPages $clamp
        if ($r.ok -and -not $dry) { Save-Printed }
        Send-Json $client $r
        return
    }

    if ($p -eq '/api/end-batch' -or $p -eq '/api/restore-printer') {
        Close-OfficeApp
        Send-Json $client ([ordered]@{ ok = $true })
        return
    }

    if ($p -eq '/api/testprint') {
        $b = Get-BodyJson $req
        $printer = [string](Get-Prop $b 'printer' '')
        # 之前这条路径完全绕过打印机校验，调用者给的字符串会直接进 COM 的 ActivePrinter。
        # 现在和 /api/print-one 用同一道门。
        if ($printer -and -not (Test-PrinterExists $printer)) {
            Send-Json $client ([ordered]@{ ok = $false; detail = "找不到打印机「$printer」" })
            return
        }
        $targets = @()
        if ($script:EngineStatus.word) { $targets += 'Word.Application' }
        if ($script:EngineStatus.wps) { $targets += 'KWPS.Application' }
        if ($targets.Count -eq 0) {
            Send-Json $client ([ordered]@{ ok = $false; detail = '没有可用的 Word/WPS，测不了测试页；可以直接试打一份 PDF 简历' })
            return
        }
        $lines = @(
            '简历一键打印 - 打印机测试页'
            '--------------------------------'
            ('时间：' + (Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))
            ('目标打印机：' + $(if ($printer) { $printer } else { '默认打印机' }))
            ('主机：' + $env:COMPUTERNAME)
            ''
            '如果你能看到这张纸，说明这台打印机和本工具都工作正常。'
        )
        $tmp = Join-Path $RecordDir ('测试页_' + (Get-Date).ToString('yyyyMMdd_HHmmss') + '.txt')
        [System.IO.File]::WriteAllText($tmp, ($lines -join "`r`n"), (New-Object System.Text.UTF8Encoding($true)))
        # 测试页文件也要封顶：以前每点一次"打印测试页"就往打印记录目录里留一个文件，没有上限。
        try {
            @(Get-ChildItem -LiteralPath $RecordDir -Filter '测试页_*.txt' -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending |
              Select-Object -Skip 10) | ForEach-Object { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue }
        } catch { }
        $errors = @()
        foreach ($prog in $targets) {
            try {
                $detail = Invoke-OfficePrint -ProgId $prog -Path $tmp -Printer $printer -Copies 1
                Send-Json $client ([ordered]@{ ok = $true; detail = $detail; file = $tmp })
                return
            } catch {
                $errors += $_.Exception.Message
                Close-OfficeApp
            }
        }
        Send-Json $client ([ordered]@{ ok = $false; detail = ($errors -join ' | ') })
        return
    }

    if ($p -eq '/api/open') {
        $b = Get-BodyJson $req
        $target = [string](Get-Prop $b 'path' '')
        # 只允许打开"打印记录目录"或本次扫描过的文件夹，不许拿它当资源管理器用。
        # 以前被拒或不存在的路径会被**静默改写成打印记录目录**，于是
        # "其余一律拒绝"这条腿实际上不存在：调用者以为开的是别处，实际开了这里。
        # 现在明确报错，并告诉调用者允许的范围。
        if ([string]::IsNullOrWhiteSpace($target)) { $target = $RecordDir }
        $allowed = @([System.IO.Path]::GetFullPath($RecordDir))
        if ($script:Config.lastFolder -and (Test-Path -LiteralPath $script:Config.lastFolder)) {
            $allowed += [System.IO.Path]::GetFullPath($script:Config.lastFolder)
        }
        if (-not (Test-Path -LiteralPath $target)) {
            Send-Json $client ([ordered]@{ ok = $false; detail = '这个路径不存在' })
            return
        }
        $full = [System.IO.Path]::GetFullPath($target)
        $okOpen = $false
        foreach ($a in $allowed) {
            if ($full -eq $a -or $full.StartsWith($a + [System.IO.Path]::DirectorySeparatorChar, [System.StringComparison]::OrdinalIgnoreCase)) { $okOpen = $true; break }
        }
        if (-not $okOpen) {
            Send-Json $client ([ordered]@{ ok = $false; detail = '只允许打开打印记录目录或本次扫描的文件夹' })
            return
        }
        Start-Process explorer.exe -ArgumentList ('"' + $target + '"')
        Send-Json $client ([ordered]@{ ok = $true })
        return
    }

    if ($p -eq '/api/reset-printed') {
        $script:Printed = @{}
        Save-Printed
        Send-Json $client ([ordered]@{ ok = $true })
        return
    }

    Send-Json $client ([ordered]@{ ok = $false; error = 'unknown api' }) 404
}

function Handle-Client($client) {
    try {
        $req = Read-HttpRequest $client
        if ($null -eq $req) {
            Send-Bytes $client 400 'text/plain; charset=utf-8' ([System.Text.Encoding]::UTF8.GetBytes('Bad Request'))
            return
        }
        # 先过请求合法性校验，再处理"请求体过大"。
        # 之前超限分支排在过滤器之前，于是"超大请求"成了唯一一种
        # Host / Origin / Content-Type 三项从不被评估的请求形态。
        $bad = Test-RequestAllowed $req
        if ($bad -ne '') {
            Log ("已拦截可疑请求（$bad）：$($req.method) $($req.path)  Host=$($req.headers['host'])  Origin=$($req.headers['origin'])") 'warn'
            $code = if ($bad -eq 'method') { 405 } else { 403 }
            Send-Json $client ([ordered]@{ ok = $false; error = "请求被拒绝（$bad）" }) $code
            return
        }
        if ($req.tooLarge) {
            Log '已拦截超大请求体' 'warn'
            Send-Json $client ([ordered]@{ ok = $false; error = '请求体过大' }) 413
            return
        }
        $p = $req.path
        if ($p -eq '/favicon.ico') { Send-Bytes $client 204 'image/x-icon' (New-Object byte[] 0); return }
        if ($p -like '/api/*') { Handle-Api $client $req; return }
        Send-Static $client $p
    } catch {
        try {
            $msg = '' + $_.Exception.Message
            Log $msg 'err'
            Send-Json $client ([ordered]@{ ok = $false; error = $msg }) 500
        } catch { }
    }
}

# ============================ 启动 ============================

$listener = $null
for ($tryPort = $Port; $tryPort -lt ($Port + 40); $tryPort++) {
    try {
        $l = New-Object System.Net.Sockets.TcpListener([System.Net.IPAddress]::Loopback, $tryPort)
        $l.Start()
        $listener = $l
        $Port = $tryPort
        break
    } catch { }
}
if ($null -eq $listener) {
    Write-Host '无法启动本地服务：端口被占用。' -ForegroundColor Red
    Read-Host '按回车退出'
    exit 1
}

$url = "http://127.0.0.1:$Port/"
Write-Host ''
Write-Host '  ============================================' -ForegroundColor Cyan
Write-Host '            简 历 一 键 打 印   已 启 动' -ForegroundColor Cyan
Write-Host '  ============================================' -ForegroundColor Cyan
Write-Host ("    网址：" + $url) -ForegroundColor Yellow
Write-Host ''
Write-Host ("    PDF 引擎 ：" + $(if ($script:EngineStatus.sumatra) { '已就绪（SumatraPDF ' + $script:EngineVersion + '）' } else { '未安装 —— 请在网页里点「下载 PDF 引擎」' })) -ForegroundColor $(if ($script:EngineStatus.sumatra) { 'Gray' } else { 'Yellow' })
if ($script:SumatraRelocated) { Write-Host ("               （程序目录含中文，已自动搬到 " + $script:SumatraPath + " 运行）") -ForegroundColor DarkGray }
Write-Host ("    Word     ：" + $(if ($script:HasWord) { '已安装（进页面后自动检测）' } else { '未安装' }))
Write-Host ("    WPS      ：" + $(if ($script:HasWps) { '已安装（进页面后自动检测）' } else { '未安装' }))
Write-Host ("    COM 助手 ：" + $(if ($script:DispReady) { 'OK（纯 IDispatch，不依赖 Office 互操作程序集）' } else { '加载失败' })) -ForegroundColor DarkGray
Write-Host ''
Write-Host '    ★ 关闭这个黑窗口就等于退出工具。' -ForegroundColor DarkGray
Write-Host ''

if (-not $NoBrowser) {
    try { Start-Process $url } catch { }
}

while ($true) {
    try {
        $client = $listener.AcceptTcpClient()
        try { Handle-Client $client } finally { try { $client.Close() } catch { } }
    } catch {
        Start-Sleep -Milliseconds 200
    }
}
