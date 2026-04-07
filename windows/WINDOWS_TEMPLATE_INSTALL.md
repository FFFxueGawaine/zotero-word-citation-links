# Windows Prebuilt Template Install Guide

Jump to: [中文](#zh-cn) | [English](#en)

<a id="zh-cn"></a>

## 中文

[Switch to English](#en)

### 这是什么

这是 Windows 下的第二种正式安装方式：

- 第一种：一键安装器
- 第二种：直接复制预改好的 `Zotero.dotm`

这份文档讲的就是第二种。

它的特点是：

- 不需要你自己手改 `customUI/customUI.xml`
- 不需要你手动导入 `ZoteroWordHyperlinks.bas`
- 本质上就是“备份原模板，再复制预改模板”

### 适合谁

- 想要最简单、最直接安装方式的人
- 不想运行动态 patch 脚本的人
- 接受“直接覆盖模板”这种方式的人

### 安装前提

- 已安装 `Microsoft Word`
- 已安装 `Zotero`
- 推荐使用 `Zotero 8.0`
- Word 中已经能看到官方 `Zotero` 选项卡
- 已关闭 `Word`

当前这套模板增强的日常使用与近期验证，主要基于 `Zotero 8.0`。

### 包内文件

这个包通常包含：

- `Zotero.dotm`
- `install_prebuilt_template.bat`
- `restore_prebuilt_template.bat`
- `WINDOWS_TEMPLATE_INSTALL.md`

### 安装方法

你可以用两种方式，本质上都是“使用预改模板”。

#### 方式 A：运行模板包里的安装脚本

1. 解压 `zotero-word-links-windows-template.zip`
2. 关闭 `Word`
3. 双击：

```text
install_prebuilt_template.bat
```

4. 重新打开 `Word`
5. 打开 `Zotero` 选项卡
6. 确认出现：
   - `Create Citation Links`
   - `Remove Citation Links`

安装完成后，链接外观由当前文档中的字符样式 `Zotero Citation Link` 控制。

#### 方式 B：自己手动复制覆盖

1. 关闭 `Word`
2. 备份当前模板：

```text
%APPDATA%\Microsoft\Word\STARTUP\Zotero.dotm
```

3. 将包内的：

```text
Zotero.dotm
```

复制并覆盖到：

```text
%APPDATA%\Microsoft\Word\STARTUP\Zotero.dotm
```

4. 重新打开 `Word`
5. 检查 `Zotero` 选项卡里的两个按钮

如果你想调整链接字体、字号、颜色或上下标，请在当前文档的样式窗格里编辑：

- `Zotero Citation Link`

更详细的样式说明见：

- 仓库文档：`docs/STYLE_GUIDE.md`
- 模板包内：`STYLE_GUIDE.md`

### 恢复方法

如果你想恢复原始状态：

#### 方法 A：运行恢复脚本

双击：

```text
restore_prebuilt_template.bat
```

#### 方法 B：手动恢复

把你之前备份的原始 `Zotero.dotm` 再复制回：

```text
%APPDATA%\Microsoft\Word\STARTUP\Zotero.dotm
```

### 这条路线的优点

- 简单
- 直观
- 不需要编辑 XML
- 不需要导入 VBA 模块

### 需要注意的地方

- 这是“直接覆盖模板”的方案，所以更依赖模板版本匹配
- 如果你的本机 `Zotero.dotm` 已经被其他工具改过，直接覆盖会把那些改动一起覆盖掉
- 如果 Zotero 官方模板将来变化较大，可能需要重新生成匹配版本的预改模板包

### 最终建议

Windows 现在只保留两种面向用户的安装方式：

1. 一键安装器
2. 预改模板直接复制覆盖

如果你只是想正常使用，优先顺序建议是：

1. 先试一键安装器
2. 如果你更喜欢直观覆盖模板，再用这个模板包

<a id="en"></a>

## English

[切换到中文](#zh-cn)

### What This Is

This is the second official Windows install path:

- first: the one-click installer
- second: direct replacement with a prebuilt `Zotero.dotm`

This guide explains the second one.

Its characteristics are:

- no need to manually edit `customUI/customUI.xml`
- no need to manually import `ZoteroWordHyperlinks.bas`
- in practice, it is just “back up the original template, then copy the prebuilt one”

### Who This Is For

- users who want the simplest and most direct setup
- users who do not want to run the dynamic patch installer
- users who are comfortable with direct template replacement

### Prerequisites

- `Microsoft Word` is installed
- `Zotero` is installed
- `Zotero 8.0` is recommended
- the standard Zotero tab is already visible in Word
- `Word` is closed

The current day-to-day workflow and recent verification for this template enhancement are primarily based on `Zotero 8.0`.

### Package Contents

This package usually contains:

- `Zotero.dotm`
- `install_prebuilt_template.bat`
- `restore_prebuilt_template.bat`
- `WINDOWS_TEMPLATE_INSTALL.md`

### Installation

You can use it in two ways, both based on the same prebuilt template.

#### Option A: Run the package installer script

1. Extract `zotero-word-links-windows-template.zip`
2. Close `Word`
3. Double-click:

```text
install_prebuilt_template.bat
```

4. Reopen `Word`
5. Open the `Zotero` tab
6. Confirm these buttons are visible:
- `Create Citation Links`
- `Remove Citation Links`

After installation, link appearance is controlled by the current document character style `Zotero Citation Link`.

#### Option B: Copy and replace the template yourself

1. Close `Word`
2. Back up the current template:

```text
%APPDATA%\Microsoft\Word\STARTUP\Zotero.dotm
```

3. Copy the packaged:

```text
Zotero.dotm
```

over:

```text
%APPDATA%\Microsoft\Word\STARTUP\Zotero.dotm
```

4. Reopen `Word`
5. Check the two buttons in the `Zotero` tab

If you want to change the link font, size, color, or superscript behavior, edit the current document style:

- `Zotero Citation Link`

For a more detailed style tutorial, see:

- repo document: `docs/STYLE_GUIDE.md`
- packaged copy: `STYLE_GUIDE.md`

### Restore

If you want to return to the original state:

#### Option A: Run the restore script

Double-click:

```text
restore_prebuilt_template.bat
```

#### Option B: Restore manually

Copy your backup `Zotero.dotm` back to:

```text
%APPDATA%\Microsoft\Word\STARTUP\Zotero.dotm
```

### Advantages

- simple
- direct
- no XML editing
- no manual VBA import

### Notes

- this is a direct template replacement workflow, so version matching matters more
- if your local `Zotero.dotm` already contains other custom changes, they will be overwritten
- if the upstream Zotero template changes significantly, a new prebuilt package may be needed

### Final Recommendation

Windows now keeps only two user-facing install methods:

1. the one-click installer
2. direct replacement with the prebuilt template package

For most users, the recommended order is:

1. try the one-click installer first
2. if you prefer a more direct template replacement workflow, use this package
