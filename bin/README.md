# 关于这个目录

`SumatraPDF.exe` **不在仓库里**（已被 `.gitignore` 排除），原因有两个：

1. 它是个 15 MB 的二进制，放进 git 会让仓库变胖、clone 变慢；
2. 它是 **GPLv3** 软件，单独走下载流程更清楚（见仓库根目录的 `THIRD-PARTY.md`）。

## 引擎是怎么来的

工具启动时如果发现这里没有 `SumatraPDF.exe`，网页顶部会提示缺少 PDF 引擎，
点一下「**下载 PDF 引擎**」就会：

1. 从 SumatraPDF 官方站点下载 `SumatraPDF-3.5.2-64.zip`
   （`https://www.sumatrapdfreader.org/dl/rel/3.5.2/SumatraPDF-3.5.2-64.zip`，
   失败时自动换 GitHub Releases 镜像）
2. 解压出可执行文件
3. **校验 SHA-256**：
   ```
   290E4AA7ED64C728138711C011E89AAB7AA48DBC1AE430371DC2BE4100B92BF0
   ```
   对不上就直接报错、不安装
4. 复制到本目录并改名成 `SumatraPDF.exe`

> 如果程序目录不可写（比如放在 `C:\Program Files` 下），会自动装到
> `%PUBLIC%\ResumePrintEngine\` 之类的可写英文目录。

## 想升级引擎版本？

改 `server.ps1` 顶部的这三行，并重新计算哈希：

```powershell
$script:EngineUrl       = '...'
$script:EngineUrlBackup = '...'
$script:EngineSha256    = '...'
$script:EngineVersion   = '3.5.2'
```

哈希算法：

```powershell
(Get-FileHash .\SumatraPDF.exe -Algorithm SHA256).Hash
```

同时记得更新根目录 `THIRD-PARTY.md` 里的版本与哈希。

## 手动放置（离线环境）

如果你是离线部署，可以把官方 `SumatraPDF-3.5.2-64.exe` 直接改名成
`SumatraPDF.exe` 放进这个目录，工具会直接使用，不会去下载。
