# Windows Manual Install Guide

Jump to: [中文](#zh-cn) | [English](#en)

<a id="zh-cn"></a>

## 中文

[Switch to English](#en)

### 适用场景

这份教程适合以下情况：

- 你不想运行安装器或脚本
- 你想完全理解这个项目在 Word 中改了什么
- 你准备自己研究模板结构，或者尝试后续移植到其他环境

### 你将手动完成什么

Windows 版的本质是对 Word 正在使用的 `Zotero.dotm` 做两类修改：

1. 给 `Zotero` 选项卡加入两个按钮
   - `Create Citation Links`
   - `Remove Citation Links`
2. 把宏模块 `ZoteroWordHyperlinks.bas` 导入模板

也就是说，手动安装的核心就是：

- 备份 `Zotero.dotm`
- 修改 `customUI/customUI.xml`
- 导入 `ZoteroWordHyperlinks.bas`

### 安装前准备

请先确认：

- 已安装 `Microsoft Word`
- 已安装 `Zotero`
- Word 中已经能看到官方 `Zotero` 选项卡
- 已关闭 `Word`

你需要用到这些文件：

- Word 模板：
  `%APPDATA%\Microsoft\Word\STARTUP\Zotero.dotm`
- 宏模块：
  `install/ZoteroWordHyperlinks.bas`

建议准备一个 Ribbon 编辑工具，最方便的是：

- `Office RibbonX Editor`

如果你不用 RibbonX Editor，也可以把 `.dotm` 当作压缩包处理，手动修改其中的 `customUI/customUI.xml`。

### 第一步：备份模板

先备份当前模板，避免出错后无法恢复。

建议复制一份：

```text
%APPDATA%\Microsoft\Word\STARTUP\Zotero.dotm
```

改名为：

```text
Zotero.backup.before-linking.dotm
```

### 第二步：修改 Ribbon 按钮

打开 `Zotero.dotm` 的 `customUI/customUI.xml`，找到 `ZoteroGroup`。

你需要确保最终顺序类似这样：

1. `Refresh`
2. `Unlink Citations`
3. 分隔符
4. `Create Citation Links`
5. `Remove Citation Links`

按钮 XML 可以参考下面这段：

```xml
<separator id="ZoteroCitationLinksSeparator" />
<button
    id="ZoteroCreateCitationLinksButton"
    label="Create Citation Links"
    imageMso="HyperlinkInsert"
    onAction="ZoteroWordHyperlinks.ZoteroCreateCitationLinks"
    supertip="Create clickable links from Zotero citations to bibliography entries"
    keytip="K" />
<button
    id="ZoteroRemoveCitationLinksButton"
    label="Remove Citation Links"
    imageMso="TableUnlinkExternalData"
    onAction="ZoteroWordHyperlinks.ZoteroRemoveCitationLinks"
    supertip="Remove citation links and bibliography bookmarks created by the hyperlink helper"
    keytip="L" />
```

如果你在手改时不确定放置位置，最重要的不是像素级顺序，而是：

- 这两个按钮必须在 `ZoteroGroup` 里
- `onAction` 名称必须和宏模块一致

### 第三步：允许 Word 访问 VBA 工程

导入 `.bas` 模块前，Word 需要允许访问 VBA 工程对象模型。

在 Word 中打开：

`文件 -> 选项 -> 信任中心 -> 信任中心设置 -> 宏设置`

勾选：

`信任对 VBA 项目对象模型的访问`

导入完成后，如果你比较谨慎，可以再关回去。

### 第四步：导入宏模块

1. 打开 Word
2. 打开 `Zotero.dotm`
3. 按 `Alt + F11` 打开 VBA 编辑器
4. 在左侧找到这个模板对应的 VBA 项目
5. 如果已经存在 `ZoteroWordHyperlinks` 模块，先删除旧版本
6. 选择 `File -> Import File...`
7. 导入：

```text
install/ZoteroWordHyperlinks.bas
```

导入后，你应该能看到这些入口宏：

- `ZoteroCreateCitationLinks`
- `ZoteroRemoveCitationLinks`

### 第五步：保存并重启 Word

1. 保存 `Zotero.dotm`
2. 关闭 Word
3. 重新打开 Word
4. 打开 `Zotero` 选项卡

此时你应该能看到：

- `Create Citation Links`
- `Remove Citation Links`

### 第六步：验证安装

最简单的测试方法：

1. 新建一个带 Zotero 引文和参考文献的 Word 文档
2. 点击 `Create Citation Links`
3. 点击正文中的引文，确认能跳到文末参考文献
4. 再点击 `Remove Citation Links`
5. 确认跳转被移除

### 恢复方法

如果你手动安装后想恢复原状：

1. 关闭 Word
2. 用备份文件覆盖当前的：
   `%APPDATA%\Microsoft\Word\STARTUP\Zotero.dotm`
3. 重新打开 Word

### 常见问题

#### 1. 为什么按钮出现了，但点了没反应？

通常有两种原因：

- `onAction` 名称写错
- `ZoteroWordHyperlinks.bas` 没有成功导入

#### 2. 为什么导入 `.bas` 时失败？

最常见原因是没有打开：

`信任对 VBA 项目对象模型的访问`

#### 3. 为什么不推荐普通用户直接手改？

因为手动安装比脚本安装更容易出错，尤其是在：

- `customUI.xml` 的位置
- 按钮插入位置
- VBA 模块命名
- 模板保存路径

如果你只是想正常使用，仍然推荐优先使用：

- `zotero-word-links-installer.exe`
- 或分享包里的 `install.bat`

<a id="en"></a>

## English

[切换到中文](#zh-cn)

### When to Use This Guide

This guide is for users who:

- do not want to run the installer or scripts
- want to fully understand what is being modified in Word
- want to study the template structure before adapting it elsewhere

### What You Are Manually Installing

The Windows version makes two kinds of changes to the active Word `Zotero.dotm` template:

1. it adds two buttons to the `Zotero` tab
   - `Create Citation Links`
   - `Remove Citation Links`
2. it imports the macro module `ZoteroWordHyperlinks.bas`

So the manual install flow is simply:

- back up `Zotero.dotm`
- edit `customUI/customUI.xml`
- import `ZoteroWordHyperlinks.bas`

### Prerequisites

Make sure:

- `Microsoft Word` is installed
- `Zotero` is installed
- the standard `Zotero` tab is already visible in Word
- `Word` is closed

You will need these files:

- Word template:
  `%APPDATA%\Microsoft\Word\STARTUP\Zotero.dotm`
- macro module:
  `install/ZoteroWordHyperlinks.bas`

The easiest tool for editing the Ribbon is:

- `Office RibbonX Editor`

If you do not use RibbonX Editor, you can also treat `.dotm` as a package and manually edit `customUI/customUI.xml`.

### Step 1: Back Up the Template

Back up the current template first:

```text
%APPDATA%\Microsoft\Word\STARTUP\Zotero.dotm
```

Suggested backup name:

```text
Zotero.backup.before-linking.dotm
```

### Step 2: Add the Ribbon Buttons

Open `customUI/customUI.xml` inside `Zotero.dotm` and find `ZoteroGroup`.

The final order should look like this:

1. `Refresh`
2. `Unlink Citations`
3. separator
4. `Create Citation Links`
5. `Remove Citation Links`

Use XML like this for the added controls:

```xml
<separator id="ZoteroCitationLinksSeparator" />
<button
    id="ZoteroCreateCitationLinksButton"
    label="Create Citation Links"
    imageMso="HyperlinkInsert"
    onAction="ZoteroWordHyperlinks.ZoteroCreateCitationLinks"
    supertip="Create clickable links from Zotero citations to bibliography entries"
    keytip="K" />
<button
    id="ZoteroRemoveCitationLinksButton"
    label="Remove Citation Links"
    imageMso="TableUnlinkExternalData"
    onAction="ZoteroWordHyperlinks.ZoteroRemoveCitationLinks"
    supertip="Remove citation links and bibliography bookmarks created by the hyperlink helper"
    keytip="L" />
```

If you are unsure about exact placement, the important parts are:

- the controls must be inside `ZoteroGroup`
- the `onAction` names must match the macro module

### Step 3: Allow Access to the VBA Project

Before importing the `.bas` file, Word must allow access to the VBA project object model.

In Word, open:

`File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings`

Enable:

`Trust access to the VBA project object model`

You can disable it again after import if you prefer.

### Step 4: Import the Macro Module

1. Open Word
2. Open `Zotero.dotm`
3. Press `Alt + F11` to open the VBA editor
4. Find the VBA project for this template
5. If a `ZoteroWordHyperlinks` module already exists, remove the old one first
6. Choose `File -> Import File...`
7. Import:

```text
install/ZoteroWordHyperlinks.bas
```

After import, you should see these public entry points:

- `ZoteroCreateCitationLinks`
- `ZoteroRemoveCitationLinks`

### Step 5: Save and Restart Word

1. Save `Zotero.dotm`
2. Close Word
3. Reopen Word
4. Open the `Zotero` tab

You should now see:

- `Create Citation Links`
- `Remove Citation Links`

### Step 6: Verify the Install

The simplest test:

1. open a document with Zotero citations and a bibliography
2. click `Create Citation Links`
3. click an in-text citation and confirm it jumps to the bibliography
4. click `Remove Citation Links`
5. confirm the generated links are removed

### Restore

To roll back:

1. close Word
2. copy your backup over:
   `%APPDATA%\Microsoft\Word\STARTUP\Zotero.dotm`
3. reopen Word

### Common Issues

#### 1. The buttons are visible, but clicking does nothing

Usually one of these:

- the `onAction` name is wrong
- `ZoteroWordHyperlinks.bas` was not imported successfully

#### 2. Importing the `.bas` file fails

The most common reason is that you did not enable:

`Trust access to the VBA project object model`

#### 3. Why is manual install not the default recommendation?

Because it is easier to make mistakes when editing:

- `customUI.xml`
- control placement
- VBA module naming
- template save location

If you just want to use the tool, the recommended paths are still:

- `zotero-word-links-installer.exe`
- or `install.bat` from the Windows share package
