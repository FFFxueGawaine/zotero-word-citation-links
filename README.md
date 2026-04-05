# Zotero Word Citation Links

Latest release: `v0.1.1`  
Changelog: [CHANGELOG.md](./CHANGELOG.md)

Jump to: [中文](#chinese) | [English](#english)

<a id="chinese"></a>

## 中文简介

[切换到 English](#english)

这是一个给 `Windows + Microsoft Word + Zotero` 使用的小型增强工具。

安装后，Word 的 `Zotero` 选项卡里会增加两个按钮：

- `Create Citation Links`
- `Remove Citation Links`

它的作用是让正文里的 Zotero 引用可以点击，并跳转到文末对应的参考文献。

### 功能特点

- 一键安装
- 普通用户默认不需要 Python
- 支持数字编号引用
- 支持作者年份引用
- 保留 Zotero 原有工作流
- 支持恢复原始 `Zotero.dotm`

### 适用环境

- Windows
- Microsoft Word 桌面版
- Zotero 已安装
- Word 中已经有 `Zotero` 选项卡

### 快速开始

最简单的方式：

1. 关闭 Word
2. 运行 `dist/zotero-word-links-installer.exe`
3. 重新打开 Word
4. 在 `Zotero` 选项卡里点击 `Create Citation Links`

如果你更喜欢脚本安装，也可以进入 `install/` 目录，运行：

- `install.bat`

### 使用方法

1. 先正常使用 Zotero 插入正文引用和参考文献
2. 点击 `Create Citation Links`
3. 点击正文中的引用，跳转到文末对应参考文献
4. 如果需要删除跳转，点击 `Remove Citation Links`
5. 如果你又执行了 `Zotero -> Refresh`，通常需要重新点击一次 `Create Citation Links`

### 仓库结构

- `install/`
  安装脚本、恢复脚本、PowerShell 安装器、宏模块
- `tools/`
  构建单文件安装器的脚本
- `dist/`
  已构建好的 `.exe` 安装器

### 已知限制

- 目前只支持 Zotero，不支持 EndNote
- 当前数字模式默认链接数字本体，不是整个括号
- Zotero 更新后，可能需要重新安装一次

<a id="english"></a>

## English

[Switch to 中文](#chinese)

This project is a small enhancement tool for `Windows + Microsoft Word + Zotero`.

After installation, it adds two buttons to the `Zotero` tab in Word:

- `Create Citation Links`
- `Remove Citation Links`

It makes Zotero citations in the document clickable and links them to the matching bibliography entries.

### Features

- One-click installation
- No Python required for normal end users
- Supports numeric citation styles
- Supports author-year citation styles
- Keeps the normal Zotero workflow
- Can restore the original `Zotero.dotm`

### Requirements

- Windows
- Microsoft Word desktop edition
- Zotero installed
- The `Zotero` tab already visible in Word

### Quick Start

The easiest way:

1. Close Word
2. Run `dist/zotero-word-links-installer.exe`
3. Open Word again
4. Click `Create Citation Links` in the `Zotero` tab

If you prefer a script-based install, go to the `install/` folder and run:

- `install.bat`

### Usage

1. Insert citations and bibliography with Zotero as usual
2. Click `Create Citation Links`
3. Click a citation in the document to jump to the bibliography entry
4. Click `Remove Citation Links` if you want to remove the generated links
5. If you run `Zotero -> Refresh`, you will usually need to run `Create Citation Links` again

### Repository Layout

- `install/`
  install scripts, restore script, PowerShell installer, macro module
- `tools/`
  script for building the standalone installer
- `dist/`
  prebuilt `.exe` installer

### Known Limitations

- Zotero only, not EndNote
- Numeric mode currently links the visible number token instead of the full bracket
- Reinstallation may be needed after Zotero updates

## License

MIT
