# Changelog

All notable changes to this project will be documented in this file.

## Unreleased

### 中文

- 新增 Mac 手工安装版文档，作为实验性支持。
- 新增 Mac 预改模板发布资产构建流程。
- README 新增平台支持矩阵和 Mac 限制说明。

### English

- Added experimental Mac manual-install documentation.
- Added a build flow for the prebuilt Mac template release asset.
- Updated the README with a platform support matrix and Mac-specific limitations.

## v0.1.1 - 2026-04-05

### 中文

- 修复了 `author-date` 引文创建链接后左右括号样式不一致的问题。
- 现在 `author-date` 模式只对括号内部正文创建链接，左右括号保持原样。
- 修复了 `author-date` 创建后正文局部出现下划线的问题。
- 修复了 `Remove Citation Links` 在 `author-date` 模式下颜色无法恢复的问题。
- 现在创建链接时会保存原始颜色，删除链接时优先恢复为创建前颜色。
- 进一步修复了“只改颜色、不改格式”的行为：
  现在无论是数字编号还是 `author-date`，创建链接都只改变颜色，不改变字体、字号、粗斜体、上下标和段落格式。
- 同步更新了安装器和分享包。

### English

- Fixed inconsistent bracket styling after creating links for `author-date` citations.
- `author-date` mode now links only the inner citation text and keeps the outer brackets unchanged.
- Fixed unwanted underline artifacts inside linked `author-date` citations.
- Fixed the color restore issue in `Remove Citation Links` for `author-date` citations.
- Original citation color is now stored on link creation and restored on link removal whenever possible.
- Improved the "change color only" behavior:
  both numeric and `author-date` modes now preserve font name, size, bold/italic, superscript/subscript, and paragraph formatting while changing only the link color.
- Updated the installer and share package.

## v0.1.0 - 2026-04-04

### 中文

- 首个公开版本。
- 支持在 Word 的 Zotero 选项卡中创建和移除引文跳转链接。
- 提供一键安装器和分享包。

### English

- First public release.
- Added create/remove citation link buttons to the Zotero tab in Word.
- Added a one-click installer and share package.
