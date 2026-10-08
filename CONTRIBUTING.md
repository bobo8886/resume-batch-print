# 贡献指南

感谢你有兴趣改进这个工具。它是给"面试季 HR / 行政"用的，
所以**稳定 > 功能多**，任何"打不出来"的情况都比少一个按钮严重得多。

## 开发环境

- Windows 7+，Windows PowerShell 5.1（不要用 PowerShell 7 测试，API 有差异）
- 不需要额外安装任何东西；`tests\run-tests.ps1` 零依赖

```powershell
git clone https://github.com/bobo8886/resume-batch-print.git
cd resume-batch-print
powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\run-tests.ps1
```

## 提交前请自测

1. **跑测试**：`powershell -NoProfile -ExecutionPolicy Bypass -File .\tests\run-tests.ps1` 全绿
2. **手动过一遍**：双击 `一键打印简历.bat` → 选文件夹 → 扫描 → **试运行** → 真打一份
3. **确认没提交隐私**：
   ```powershell
   git status --porcelain
   ```
   `config.json`、`打印记录\`、`bin\SumatraPDF.exe` 都不应该出现在待提交列表里

## 代码约定

- 所有 `.ps1` 必须存成 **UTF-8 with BOM**。Windows PowerShell 5.1 读没有 BOM 的 UTF-8
  会把中文变乱码，甚至直接解析失败。
- 脚本里不要出现任何本机绝对路径、用户名、打印机名之类的个人信息。
- 中文注释写在"为什么"上，不是"做了什么"上。

---

## ⚠️ 千万别踩回去的坑

这几个都是实际调试了很久才定位到的问题，改代码时请务必保留对应处理。

### 1. SumatraPDF 的程序路径不能含非 ASCII 字符

只要 `SumatraPDF.exe` **自己所在的路径**有中文，打印就会失败：

```
PrintToDevice: StartDoc() failed with -1
```

注意是**程序路径**，不是简历文件路径——简历放中文目录完全没问题。
所以启动时 `Resolve-SumatraEngine` 会把它复制到 `%PUBLIC%\ResumePrintEngine\` 之类
的纯英文目录再调用。

### 2. Office 自动化不能走 PowerShell 原生写法

```powershell
$word = New-Object -ComObject Word.Application
$word.Documents          # ← 在 PIA 损坏的机器上返回 null！
```

很多电脑上 Office 互操作程序集（PIA）是坏的或没装。这时对象**能创建成功**，
但所有属性都返回 `null`，而且不一定报错，非常难查。

本项目改为内嵌一段 C#（`server.ps1` 里的 `Disp` 类）直接走
`IDispatch::GetIDsOfNames` + `IDispatch::Invoke`，完全绕开 PIA。
另外还要处理两种 HRESULT：

- `RPC_E_CALL_REJECTED (0x80010001)`、`RPC_E_SERVERCALL_RETRYLATER (0x8001010A)`
  → Office 正忙，要重试
- 属性写入（`DISPATCH_PROPERTYPUT`）成功时会返回一个 .NET 无法封送的 VARIANT
  → 这不是失败，不要读返回值

### 3. SumatraPDF 打印时一律栅格化整页

官方文档写得很清楚。这意味着**浅灰色内容会被冲淡**（被放大插值）。
「打印浓度」功能因此改成：用 WinRT `Windows.Data.Pdf` 自己渲染 → gamma 查表加深 →
用自带的极简 PDF 写入器重建 PDF → 再交给 SumatraPDF 打印。

顺带两个坑：

- `PdfPage.Size` 的单位是 **DIP（1/96 英寸）不是 pt**，要 ×0.75 才是 PDF 的 pt
- WinRT 的 `DestinationWidth` 实际会被放大（约 1.5 倍），别按字面理解

### 4. 不要用 GDI+ 的 PrintDocument 打 PDF

在测试机上它会 `Print()` 返回成功、打印队列也清空，但**打印机一个字节都没产出**。
所以 PDF 打印统一走 SumatraPDF，不要改成 `PrintDocument`。

### 5. 本地服务必须做来源校验

这是一个"能读本地目录、还能指挥打印机"的本地 HTTP 服务。少了
Host / Origin / Content-Type 校验，任何一个网页都能通过 CSRF / DNS 重绑定
让用户的打印机吐纸，或者枚举他的磁盘目录。详见 `SECURITY.md`。

对应实现是 `server.ps1` 里的 `Test-RequestAllowed`，**不要删**。

---

## 提交 PR

- 一个 PR 只做一件事，标题写清楚"修了什么/加了什么"
- 涉及行为变更的，同步更新 `README.md` 和 `CHANGELOG.md`
- 新增安全相关处理的，同步更新 `SECURITY.md`
