# Changelog

All notable changes to this project will be documented in this file.

## Unreleased

### 中文

- 新增 Windows 纯手动安装教程，说明如何手动修改 `Zotero.dotm`、添加 Ribbon 按钮并导入宏模块。
- 补充了如何打开 `Zotero.dotm` 中 `customUI/customUI.xml` 的具体方法，包括 RibbonX Editor 和压缩包两种路径。
- 新增 Windows 预改模板包方案，支持直接覆盖 `Zotero.dotm` 或运行简单复制脚本完成安装。
- 新增 Windows 预改模板包构建脚本、安装脚本、恢复脚本和安装说明。
- 将 Windows 面向普通用户的安装方式收敛为两种：一键安装，或直接复制预改模板。

### English

- Added a Windows manual install guide describing how to modify `Zotero.dotm`, add the Ribbon buttons, and import the macro module by hand.
- Expanded the manual guide with concrete ways to open `customUI/customUI.xml`, including both RibbonX Editor and archive-based workflows.
- Added a Windows prebuilt template package path that supports direct `Zotero.dotm` replacement or a simple copy-based install script.
- Added the Windows prebuilt template package build script, install script, restore script, and install guide.
- Simplified the Windows user-facing install story to two methods only: one-click install, or direct replacement with the prebuilt template.

## v0.3.0 - 2026-04-05

### 中文

- 正式发布 `v0.3.0`，让外部按版本号或 Release 检测更新时能够正确识别新版本。
- 新增 Mac 一键安装脚本 `install_mac.command`。
- 新增 Mac 恢复脚本 `restore_mac.command`。
- README 重新整理为更清晰的中英双语结构。
- README 强化了安装说明和使用教程，重点突出 Windows / Mac 的安装入口与按钮使用流程。

### English

- Released `v0.3.0` as a proper detectable version for tag-based and release-based update checks.
- Added the one-click Mac installer script `install_mac.command`.
- Added the Mac restore script `restore_mac.command`.
- Reorganized the README into a clearer bilingual structure.
- Expanded the install and usage tutorial sections, with clearer Windows / Mac paths and button workflow guidance.

## v0.2.0 - 2026-04-05

### 中文

- 新增 Mac 手工安装版文档，作为实验性支持。
- 新增 Mac 预改模板发布资产构建流程。
- 新增 `install_mac.command` 一键安装脚本和 `restore_mac.command` 恢复脚本。
- README 新增平台支持矩阵和 Mac 限制说明。

### English

- Added experimental Mac manual-install documentation.
- Added a build flow for the prebuilt Mac template release asset.
- Added `install_mac.command` for one-click install and `restore_mac.command` for restore.
- Updated the README with a platform support matrix and Mac-specific limitations.

## v0.1.1 - 2026-04-05

### 中文

- 修复了 `author-date` 引文创建链接后左右括号样式不一致的问题。
- 现在 `author-date` 模式只对括号内部正文创建链接，左右括号保持原样。
- 修复了 `author-date` 创建后正文局部出现下划线的问题。
- 修复了 `Remove Citation Links` 在 `author-date` 模式下颜色无法恢复的问题。
- 现在创建链接时会保存原始颜色，删除链接时优先恢复为创建前颜色。
- 进一步修复了“只改颜色、不改格式”的行为。
- 同步更新了安装器和分享包。

### English

- Fixed inconsistent bracket styling after creating links for `author-date` citations.
- `author-date` mode now links only the inner citation text and keeps the outer brackets unchanged.
- Fixed unwanted underline artifacts inside linked `author-date` citations.
- Fixed the color restore issue in `Remove Citation Links` for `author-date` citations.
- Original citation color is now stored on link creation and restored on link removal whenever possible.
- Improved the "change color only" behavior.
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
