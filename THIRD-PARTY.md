# 第三方组件

本项目的**源代码**以 MIT 许可证发布（见 `LICENSE`）。
但它会调用 / 下载下面这个第三方组件，该组件有自己的许可证，**不受 MIT 覆盖**。

---

## SumatraPDF（PDF 静默打印引擎）

| 项目 | 说明 |
|---|---|
| 名称 | SumatraPDF |
| 版本 | **3.5.2**（64 位） |
| 作者 | Krzysztof Kowalczyk 及贡献者 |
| 许可证 | **GNU General Public License v3.0 (GPLv3)** |
| 官网 | <https://www.sumatrapdfreader.org/> |
| 源码 | <https://github.com/sumatrapdfreader/sumatrapdf> |
| 许可证全文 | <https://www.gnu.org/licenses/gpl-3.0.html> |

### 为什么需要它

SumatraPDF 官方文档明确说明：它打印时会把整页栅格化成位图再送给打印机，
因此对 PDF 的兼容性和"所见即所得"程度都很好。
本工具用它来完成 **PDF 的静默打印（不弹窗、不需要人工点确认）**。

### 它是怎么被分发的

**本仓库不包含 SumatraPDF 的任何二进制文件**（`bin/SumatraPDF.exe` 已被 `.gitignore` 排除）。

- 用户首次使用时，工具会**从 SumatraPDF 官方站点下载**官方发布的压缩包；
- 下载地址固定为：
  `https://www.sumatrapdfreader.org/dl/rel/3.5.2/SumatraPDF-3.5.2-64.zip`
  （备用镜像：GitHub Releases）
- 下载后会校验 **SHA-256**，与代码中固定的哈希值比对，不一致就拒绝安装：

  ```
  290E4AA7ED64C728138711C011E89AAB7AA48DBC1AE430371DC2BE4100B92BF0
  ```

### GPLv3 合规说明

- 本工具通过**独立进程调用** SumatraPDF（命令行参数），二者是**各自独立的程序**，
  属于 GPL 所称的 "mere aggregation"（单纯聚合），不构成衍生作品，
  因此本项目的 MIT 许可证不受影响。
- 如果你要**再分发**这个工具（例如打包成 zip 发给同事、或在 GitHub Releases 里附带完整包），
  请一并保留本节说明，并让接收者知道 SumatraPDF 是 GPLv3 软件、源码在哪里获取。
- 需要完整对应源码的用户，请前往上面的官方源码仓库，或联系官方发布者。
- SumatraPDF 的 GPLv3 许可证要求：分发二进制时须附带许可证全文。若你打包分发，
  请把 <https://www.gnu.org/licenses/gpl-3.0.txt> 一并放入压缩包。

---

## Microsoft Office / WPS Office（可选）

打印 `.doc` `.docx` `.rtf` `.txt` `.odt` `.wps` 时，本工具通过 **COM 自动化**调用
用户本机已安装的 Microsoft Word 或 WPS Writer。

- 本工具**不包含、也不再分发**这些软件的任何部分；
- 它们由用户自行安装并遵守各自的商业许可；
- 没有安装它们时，Word 系文档会走「转 PDF 再打印」或直接给出明确提示。

---

## Windows 系统组件

以下为操作系统自带能力，不涉及第三方分发：

- `Windows.Data.Pdf`（WinRT）—— PDF 栅格化，用于「打印浓度 / 加深」功能（Windows 10+）
- `System.Drawing`（GDI+）—— 位图处理
- PowerShell 5.1 / .NET Framework —— 运行环境
