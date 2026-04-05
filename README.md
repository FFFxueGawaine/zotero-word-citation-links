# Zotero Word Citation Links

Latest release: `v0.2.0`  
Downloads:
- [Latest Release Page](https://github.com/FFFxueGawaine/zotero-word-citation-links/releases/tag/v0.2.0)
- [Windows Installer](https://github.com/FFFxueGawaine/zotero-word-citation-links/releases/download/v0.2.0/zotero-word-links-installer.exe)
- [Windows Share Package](https://github.com/FFFxueGawaine/zotero-word-citation-links/releases/download/v0.2.0/zotero-word-links-share.zip)
- [Mac Template Package](https://github.com/FFFxueGawaine/zotero-word-citation-links/releases/download/v0.2.0/zotero-word-links-mac-template.zip)

Changelog: [CHANGELOG.md](./CHANGELOG.md)

Jump to: [中文](#中文) | [English](#english)

<a id="中文"></a>

## 中文

[Switch to English](#english)

### 项目简介

这是一个给 `Microsoft Word + Zotero` 使用的小型增强工具。  
它会在 Word 的 `Zotero` 选项卡中增加两个按钮：

- `Create Citation Links`
- `Remove Citation Links`

作用很直接：

- 让正文里的 Zotero 引文可以点击
- 点击后跳转到文末对应参考文献
- 尽量只改变引文颜色，不破坏原有字体、字号和段落格式

### 功能特点

- 支持数字编号引用
- 支持作者年份引用
- 保留 Zotero 原有写作流程
- Windows 提供一键安装
- Mac 提供实验性的手工安装方案
- 支持恢复原始 `Zotero.dotm`

### 支持情况

| 平台 | 状态 | 安装方式 |
| --- | --- | --- |
| Windows + Word | 正式支持 | 一键安装器 / 脚本安装 |
| Mac + Word | 实验性支持 | 手工安装预改 `Zotero.dotm` |

### 安装前提

- 已安装 `Zotero`
- 已安装 `Microsoft Word`
- Word 中已经可以看到官方 `Zotero` 选项卡

### Windows 安装

#### 方式一：一键安装器

推荐普通用户直接使用：

1. 关闭 `Word`
2. 下载并运行 [Windows Installer](https://github.com/FFFxueGawaine/zotero-word-citation-links/releases/download/v0.2.0/zotero-word-links-installer.exe)
3. 重新打开 `Word`
4. 打开 `Zotero` 选项卡，确认出现：
   - `Create Citation Links`
   - `Remove Citation Links`

#### 方式二：脚本安装

如果你更想看安装内容或自己控制流程：

1. 下载 [Windows Share Package](https://github.com/FFFxueGawaine/zotero-word-citation-links/releases/download/v0.2.0/zotero-word-links-share.zip)
2. 解压
3. 关闭 `Word`
4. 运行 `install/install.bat`

### Mac 安装

#### 实验性手工安装

Mac 当前不提供自动安装器，只提供实验性的手工安装方式。

1. 下载 [Mac Template Package](https://github.com/FFFxueGawaine/zotero-word-citation-links/releases/download/v0.2.0/zotero-word-links-mac-template.zip)
2. 关闭 `Word`
3. 备份你当前的 `Zotero.dotm`
4. 将压缩包中的预改 `Zotero.dotm` 复制到 Word Startup 模板目录
5. 重新打开 `Word`
6. 打开 `Zotero` 选项卡，确认出现：
   - `Create Citation Links`
   - `Remove Citation Links`

详细步骤请看：
[mac/MAC_INSTALL.md](./mac/MAC_INSTALL.md)

### 使用方法

1. 先像平时一样，用 Zotero 插入正文引用和参考文献
2. 点击 `Create Citation Links`
3. 点击正文中的引用，跳转到文末对应参考文献
4. 如需移除跳转，点击 `Remove Citation Links`
5. 如果你又执行了 `Zotero -> Refresh`，通常需要重新点击一次 `Create Citation Links`

### 仓库结构

- `install/`  
  Windows 安装脚本、恢复脚本、宏模块
- `mac/`  
  Mac 手工安装文档和 Release 说明
- `tools/`  
  Windows / Mac 资产构建脚本
- `dist/`  
  已构建好的发布资产

### 已知限制

- 目前只支持 Zotero，不支持 EndNote
- 当前数字模式默认链接数字本体，不是整个括号
- Mac 版本当前为实验性支持，未在所有 Mac / Word 版本上完整验证
- Zotero 更新后，可能需要重新安装匹配版本模板

<a id="english"></a>

## English

[切换到中文](#中文)

### Overview

This is a small enhancement tool for `Microsoft Word + Zotero`.

It adds two buttons to the `Zotero` tab in Word:

- `Create Citation Links`
- `Remove Citation Links`

Its purpose is simple:

- make Zotero citations in the document clickable
- jump from an in-text citation to the matching bibliography entry
- change citation color while preserving font, size, and paragraph formatting as much as possible

### Features

- Supports numeric citation styles
- Supports author-year citation styles
- Keeps the standard Zotero workflow
- One-click install on Windows
- Experimental manual install on Mac
- Can restore the original `Zotero.dotm`

### Support Matrix

| Platform | Status | Install Mode |
| --- | --- | --- |
| Windows + Word | Supported | One-click installer / script install |
| Mac + Word | Experimental | Manual install with a prebuilt `Zotero.dotm` |

### Prerequisites

- `Zotero` is installed
- `Microsoft Word` is installed
- The standard Zotero tab is already visible in Word

### Windows Installation

#### Option 1: One-click installer

Recommended for most users:

1. Close `Word`
2. Download and run the [Windows Installer](https://github.com/FFFxueGawaine/zotero-word-citation-links/releases/download/v0.2.0/zotero-word-links-installer.exe)
3. Reopen `Word`
4. Open the `Zotero` tab and confirm these buttons are visible:
   - `Create Citation Links`
   - `Remove Citation Links`

#### Option 2: Script install

If you prefer a more transparent/manual flow:

1. Download the [Windows Share Package](https://github.com/FFFxueGawaine/zotero-word-citation-links/releases/download/v0.2.0/zotero-word-links-share.zip)
2. Extract it
3. Close `Word`
4. Run `install/install.bat`

### Mac Installation

#### Experimental manual install

Mac currently does not include an automatic installer.  
It is provided as an experimental manual-install workflow.

1. Download the [Mac Template Package](https://github.com/FFFxueGawaine/zotero-word-citation-links/releases/download/v0.2.0/zotero-word-links-mac-template.zip)
2. Quit `Word`
3. Back up your current `Zotero.dotm`
4. Copy the prebuilt `Zotero.dotm` into the Word Startup template folder
5. Reopen `Word`
6. Open the `Zotero` tab and confirm these buttons are visible:
   - `Create Citation Links`
   - `Remove Citation Links`

Detailed instructions:
[mac/MAC_INSTALL.md](./mac/MAC_INSTALL.md)

### Usage

1. Insert citations and bibliography with Zotero as usual
2. Click `Create Citation Links`
3. Click an in-text citation to jump to the bibliography entry
4. Click `Remove Citation Links` if you want to remove the generated links
5. If you run `Zotero -> Refresh`, you will usually need to run `Create Citation Links` again

### Repository Layout

- `install/`  
  Windows install scripts, restore script, macro module
- `mac/`  
  Mac manual-install documentation and release notes
- `tools/`  
  Windows / Mac asset build scripts
- `dist/`  
  built release assets

### Known Limitations

- Zotero only, not EndNote
- Numeric mode currently links the visible number token instead of the full bracket
- Mac support is currently experimental and not fully validated across all Mac / Word versions
- Reinstallation may be needed after Zotero updates

## License

MIT
