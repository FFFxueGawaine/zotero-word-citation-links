# Windows Prebuilt Template Install Guide

Jump to: [中文](#zh-cn) | [English](#en)

<a id="zh-cn"></a>

## 中文

[Switch to English](#en)

### 这是什么

这是 Windows 下最简单的一条安装路线之一：

- 直接使用一个已经预先修改好的 `Zotero.dotm`
- 不需要你自己手改 `customUI/customUI.xml`
- 不需要你手动导入 `ZoteroWordHyperlinks.bas`

你可以把它理解成：

\[
\text{备份原模板} + \text{复制预改模板} = \text{完成安装}
\]

### 适合谁

- 想要尽量简单安装的人
- 不想运行“动态 patch 模板”脚本的人
- 接受“直接覆盖模板”这种方式的人

### 安装前提

- 已安装 `Microsoft Word`
- 已安装 `Zotero`
- Word 中已经能看到官方 `Zotero` 选项卡
- 已关闭 `Word`

### 包内文件

这个包通常包含：

- `Zotero.dotm`
- `install_prebuilt_template.bat`
- `restore_prebuilt_template.bat`
- `WINDOWS_TEMPLATE_INSTALL.md`

### 安装方法

#### 方法 A：推荐，直接运行安装脚本

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

#### 方法 B：纯复制覆盖

如果你想完全手动操作：

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
5. 检查 `Zotero` 选项卡中的两个新按钮

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

### 优点

- 简单
- 不需要编辑 XML
- 不需要导入 VBA 模块
- 对普通用户更直观

### 注意事项

- 这是“直接覆盖模板”的方案，所以更依赖模板版本匹配
- 如果你的本机 `Zotero.dotm` 已被其他工具改过，直接覆盖会把那些改动一起覆盖掉
- 如果 Zotero 官方模板将来发生明显变化，可能需要重新生成匹配版本的预改模板包

### 推荐理解

Windows 现在有三条路线：

1. 安装器自动 patch
   好处：最稳，适合大多数用户
2. 预改模板直接覆盖
   好处：最简单
3. 纯手动安装
   好处：最透明，适合研究和二次开发

<a id="en"></a>

## English

[切换到中文](#zh-cn)

### What This Is

This is one of the simplest Windows install paths:

- it uses a prebuilt `Zotero.dotm`
- you do not need to edit `customUI/customUI.xml` yourself
- you do not need to manually import `ZoteroWordHyperlinks.bas`

In practice, it is:

\[
\text{backup original template} + \text{copy prebuilt template} = \text{install complete}
\]

### Who This Is For

- users who want the simplest setup
- users who do not want to run the dynamic patch installer
- users who are comfortable with direct template replacement

### Prerequisites

- `Microsoft Word` is installed
- `Zotero` is installed
- the standard Zotero tab is already visible in Word
- `Word` is closed

### Package Contents

This package usually contains:

- `Zotero.dotm`
- `install_prebuilt_template.bat`
- `restore_prebuilt_template.bat`
- `WINDOWS_TEMPLATE_INSTALL.md`

### Installation

#### Method A: Recommended, run the installer script

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

#### Method B: Manual copy/replace

If you want a fully manual route:

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
5. Check the `Zotero` tab for the two added buttons

### Restore

If you want to return to the original state:

#### Method A: Run the restore script

Double-click:

```text
restore_prebuilt_template.bat
```

#### Method B: Restore manually

Copy your backup `Zotero.dotm` back to:

```text
%APPDATA%\Microsoft\Word\STARTUP\Zotero.dotm
```

### Advantages

- simple
- no XML editing
- no manual VBA import
- more intuitive for normal users

### Notes

- this is a direct template replacement workflow, so version matching matters more
- if your local `Zotero.dotm` already contains other custom changes, they will be overwritten
- if the upstream Zotero template changes significantly, a new prebuilt package may be needed

### Recommended Mental Model

There are now three Windows install paths:

1. installer-based patching
   benefit: most robust for most users
2. prebuilt template replacement
   benefit: simplest
3. fully manual install
   benefit: most transparent for research and adaptation
