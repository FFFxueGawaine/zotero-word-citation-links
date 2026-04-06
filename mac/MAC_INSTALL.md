# Mac Manual Install Guide

Jump to: [中文](#中文) | [English](#english)

<a id="中文"></a>

## 中文

这是当前项目的 **Mac 实验性支持说明**。

当前版本只支持：

- `Zotero + Microsoft Word for Mac`
- 预改好的 `Zotero.dotm`
- `.command` 一键安装脚本

当前版本不包含：

- `.pkg` 自动安装包
- LibreOffice / EndNote / Pages / WPS 支持

### 你需要先确认

1. 你的 Word 里已经能正常看到 `Zotero` 选项卡  
   如果官方 Zotero Word 插件本身都没有正常工作，请先修复官方插件。
2. 你已经从本项目的 Release 下载了：
   - `zotero-word-links-mac-template.zip`

### Mac 安装步骤

推荐优先使用一键安装脚本。

#### 方式一：一键安装

1. 完全关闭 `Microsoft Word`
2. 解压 `zotero-word-links-mac-template.zip`
3. 双击运行：
   - `install_mac.command`
4. 如果 macOS 首次拦截，右键脚本，选择 `Open`
5. 等脚本提示安装完成
6. 重新打开 `Word`
7. 打开 `Zotero` 选项卡，确认出现：
   - `Create Citation Links`
   - `Remove Citation Links`
   - `Set Link Color`

`Set Link Color` 会优先尝试 Word 自带的颜色对话框；如果当前 Mac 版本无法稳定取色，再回退到预设颜色 / 自定义 `RGB`。

#### 方式二：手工安装

1. 完全关闭 `Microsoft Word`
2. 找到你当前正在使用的 `Zotero.dotm`
3. 先备份原文件
4. 解压 `zotero-word-links-mac-template.zip`
5. 将压缩包中的 `Zotero.dotm` 复制到 Word Startup 模板目录
6. 重新打开 `Word`
7. 打开 `Zotero` 选项卡，确认出现：
   - `Create Citation Links`
   - `Remove Citation Links`
   - `Set Link Color`

如果原生颜色对话框在你的 Mac / Word 版本上不可用，`Set Link Color` 仍会保留当前的输入式兜底流程。

### 常见路径

Word for Mac 常见 Startup 路径通常类似：

`~/Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Word/`

如果你不确定路径，可以这样找：

1. 在 Finder 中按 `Shift + Command + G`
2. 输入：
   `~/Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Word/`
3. 回车后查看目录中是否存在 `Zotero.dotm`

### 恢复原版

如果安装后异常：

1. 关闭 `Word`
2. 优先双击运行：
   - `restore_mac.command`
3. 如果脚本无法使用，再手工删除当前替换进去的 `Zotero.dotm`
4. 把你之前备份的原始 `Zotero.dotm` 复制回去
5. 重新打开 `Word`

### 版本绑定说明

这个 Mac 模板不是从零生成的，而是基于 **Zotero 官方 Mac Word 模板** 修改得到。

- 如果你的 Zotero Word 集成版本不同，可能会出现：
  - 按钮不显示
  - Zotero 交互异常
  - 模板加载失败
- 如果出现异常，优先回滚到你原始备份的 `Zotero.dotm`

### 许可说明

Release 中提供的 Mac `Zotero.dotm` 是基于 Zotero 官方 Mac 模板的派生版本。  
相关上游来源、版本和许可信息会随压缩包一起提供。

<a id="english"></a>

## English

This is the **experimental Mac support** guide for the project.

Current Mac support is limited to:

- `Zotero + Microsoft Word for Mac`
- a prebuilt `Zotero.dotm`
- a `.command` one-click installer

It does not include:

- a `.pkg` auto-install package
- LibreOffice / EndNote / Pages / WPS support

### Before you start

1. Make sure the standard Zotero Word integration already works on your Mac  
   If the official Zotero plugin is not working yet, fix that first.
2. Download this release asset:
   - `zotero-word-links-mac-template.zip`

### Install Steps

Use the one-click script first if possible.

#### Option 1: One-click install

1. Fully quit `Microsoft Word`
2. Extract `zotero-word-links-mac-template.zip`
3. Double-click:
   - `install_mac.command`
4. If macOS blocks it the first time, right-click the script and choose `Open`
5. Wait for the installer to finish
6. Reopen `Word`
7. Open the `Zotero` tab and confirm these buttons are visible:
   - `Create Citation Links`
   - `Remove Citation Links`
   - `Set Link Color`

`Set Link Color` first tries Word's native color dialog. If the current Mac / Word version cannot return a stable color value, it falls back to preset colors / custom `RGB`.

#### Option 2: Manual install

1. Fully quit `Microsoft Word`
2. Locate the `Zotero.dotm` currently used by Word
3. Back up the original file first
4. Extract `zotero-word-links-mac-template.zip`
5. Copy the included `Zotero.dotm` into the Word Startup template folder
6. Reopen `Word`
7. Open the `Zotero` tab and confirm these buttons are visible:
   - `Create Citation Links`
   - `Remove Citation Links`
   - `Set Link Color`

If the native color dialog is not usable on your Mac / Word version, `Set Link Color` still keeps the current input-based fallback path.

### Common Path

A common Word Startup path on macOS is:

`~/Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Word/`

If you are unsure, use Finder:

1. Press `Shift + Command + G`
2. Paste:
   `~/Library/Group Containers/UBF8T346G9.Office/User Content/Startup/Word/`
3. Check whether `Zotero.dotm` exists there

### Restore the Original Template

If anything goes wrong:

1. Quit `Word`
2. Prefer running:
   - `restore_mac.command`
3. If the script is unavailable, delete the replaced `Zotero.dotm`
4. Copy your original backup back into the folder
5. Reopen `Word`

### Version Binding

This Mac template is a modified derivative of Zotero's official Mac Word template.

- If your Zotero Word integration version differs, you may see:
  - missing buttons
  - Zotero integration errors
  - template loading failures
- If that happens, restore your original `Zotero.dotm` first

### License Note

The Mac `Zotero.dotm` included in the release is a derived template based on Zotero's official Mac template.  
The package includes upstream source, version, and license notes.
