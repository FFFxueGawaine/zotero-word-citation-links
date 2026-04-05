# Zotero Word Citation Links

> Add clickable citation-to-bibliography links to Zotero citations in Microsoft Word, without disrupting the normal Zotero writing workflow.

Latest release: `v0.3.0`

Downloads:
- [Latest Release Page](https://github.com/FFFxueGawaine/zotero-word-citation-links/releases/latest)
- [Windows Installer](https://github.com/FFFxueGawaine/zotero-word-citation-links/releases/latest/download/zotero-word-links-installer.exe)
- [Windows Share Package](https://github.com/FFFxueGawaine/zotero-word-citation-links/releases/latest/download/zotero-word-links-share.zip)
- [Mac Template Package](https://github.com/FFFxueGawaine/zotero-word-citation-links/releases/latest/download/zotero-word-links-mac-template.zip)

Changelog: [CHANGELOG.md](./CHANGELOG.md)

Jump to: [中文](#zh-cn) | [English](#en)

<a id="zh-cn"></a>

## 中文

[Switch to English](#en)

### 项目简介

这是一个给 `Microsoft Word + Zotero` 使用的小型增强工具。

它会在 Word 的 `Zotero` 选项卡中增加两个按钮：

- `Create Citation Links`
- `Remove Citation Links`

它解决的问题很直接：

- 让正文中的 Zotero 引文可以点击
- 点击后跳转到文末对应参考文献
- 尽量只改变引文颜色，不破坏原有字体、字号、粗斜体、上下标和段落格式

### 适合谁

- 想让 Word 文档中的 Zotero 引文支持跳转的用户
- 正在写论文、综述、报告，希望提高文内定位效率的用户
- 不想改变原有 Zotero 使用方式，只想增加一个实用增强功能的用户

### 主要特点

| 特点 | 说明 |
| --- | --- |
| 支持数字编号格式 | 例如 `[1]`、`[2, 3]` |
| 支持作者-年份格式 | 例如 `(Smith, 2024)` |
| 使用方式直观 | 安装后直接在 Word 的 `Zotero` 选项卡中点击 |
| 尽量保留原格式 | 创建链接时主要只改变颜色，不改动排版结构 |
| 支持恢复 | 可以移除跳转链接，也可以恢复原始 `Zotero.dotm` |

### 支持情况

| 平台 | 状态 | 安装方式 |
| --- | --- | --- |
| Windows + Word | 正式支持 | 一键安装器 / 脚本安装 |
| Mac + Word | 实验性支持 | `.command` 一键安装 / 手工安装 |

### 安装前提

- 已安装 `Zotero`
- 已安装 `Microsoft Word`
- Word 中已经能看到官方的 `Zotero` 选项卡

### 安装方式

#### Windows

推荐普通用户直接使用一键安装器：

1. 关闭 `Word`
2. 下载并运行 [Windows Installer](https://github.com/FFFxueGawaine/zotero-word-citation-links/releases/latest/download/zotero-word-links-installer.exe)
3. 重新打开 `Word`
4. 打开 `Zotero` 选项卡，确认出现：
   - `Create Citation Links`
   - `Remove Citation Links`

如果你更喜欢看清楚安装内容，也可以使用脚本包：

1. 下载 [Windows Share Package](https://github.com/FFFxueGawaine/zotero-word-citation-links/releases/latest/download/zotero-word-links-share.zip)
2. 解压
3. 关闭 `Word`
4. 运行 `install.bat`

如果你想完全不依赖安装器和脚本，而是自己手动修改模板，请看：

[install/WINDOWS_MANUAL_INSTALL.md](./install/WINDOWS_MANUAL_INSTALL.md)

#### Mac

当前 Mac 版本为实验性支持，推荐先从模板包开始：

1. 下载 [Mac Template Package](https://github.com/FFFxueGawaine/zotero-word-citation-links/releases/latest/download/zotero-word-links-mac-template.zip)
2. 关闭 `Word`
3. 解压后双击：
   - `install_mac.command`
4. 如果 macOS 首次拦截，右键脚本并选择 `Open`
5. 等待脚本完成备份和安装
6. 重新打开 `Word`
7. 打开 `Zotero` 选项卡，确认出现：
   - `Create Citation Links`
   - `Remove Citation Links`

详细说明请看：
[mac/MAC_INSTALL.md](./mac/MAC_INSTALL.md)

### 使用教程

这是最重要的一部分。你可以把它理解成一个标准工作流。

#### 第一步：正常写作

先按平时的 Zotero 用法写作：

1. 用 Zotero 插入正文引文
2. 用 Zotero 生成参考文献
3. 正常修改、补充、刷新引文

#### 第二步：生成跳转

当你的文档已经有了正文引文和文末参考文献后：

1. 打开 Word 的 `Zotero` 选项卡
2. 点击 `Create Citation Links`
3. 文中引文会变成可点击状态
4. 点击引文，可跳转到对应参考文献

#### 第三步：需要时删除跳转

如果你想移除这次生成的跳转效果：

1. 点击 `Remove Citation Links`
2. 文中的跳转会被移除
3. 引文颜色会尽量恢复到创建前的状态

### 按钮说明

| 按钮 | 作用 | 什么时候用 |
| --- | --- | --- |
| `Create Citation Links` | 为文中 Zotero 引文创建跳转 | 当你已经插入好引文和参考文献，想添加点击跳转时 |
| `Remove Citation Links` | 移除本工具创建的跳转 | 当你想恢复普通显示，或准备重新生成跳转时 |

### 推荐使用节奏

如果你想要最稳的体验，建议这样用：

1. 先完成 Zotero 的正常引文编辑
2. 如果你刚点过 `Zotero -> Refresh`，先别急着测试跳转
3. 再点一次 `Create Citation Links`
4. 最后再检查点击跳转效果

原因很简单：

- `Zotero -> Refresh` 会重写 Word 中的引文结果
- 所以刷新之后，通常需要重新执行一次 `Create Citation Links`

### 典型效果

数字编号格式：

- `[1]`
- `[2, 3]`

作者-年份格式：

- `(Smith, 2024)`
- `(Kumar et al., 2026; Yu et al., 2025)`

当前设计目标是：

- 数字格式创建后只改变颜色
- 作者-年份格式创建后只让中间正文成为链接，括号保持普通样式
- 删除后尽量恢复原颜色，不保留下划线

### 恢复与回退

如果你想恢复安装前状态：

- Windows：运行分享包中的 `restore_original.bat`
- Mac：运行模板包中的 `restore_mac.command`

### 已知限制

- 当前只支持 `Zotero`，不支持 `EndNote`
- 数字模式默认链接数字本体，不是整个括号
- Mac 当前仍是实验性支持，尚未在所有 Mac / Word 版本上完成实机验证
- Zotero 更新后，可能需要重新安装匹配版本

### 仓库结构

- `install/`
  Windows 安装脚本、恢复脚本、宏模块、纯手动安装文档
- `mac/`
  Mac 安装文档和相关说明
- `tools/`
  构建脚本
- `dist/`
  发布资产

<a id="en"></a>

## English

[切换到中文](#zh-cn)

### Overview

This project is a lightweight enhancement for `Microsoft Word + Zotero`.

It adds two buttons to the `Zotero` tab in Word:

- `Create Citation Links`
- `Remove Citation Links`

Its goal is simple:

- make Zotero citations in the document clickable
- jump from an in-text citation to the matching bibliography entry
- change citation color while preserving font, size, style, superscript/subscript, and paragraph formatting as much as possible

### Who This Is For

- users who want clickable Zotero citations in Word
- researchers writing papers, reviews, reports, or theses
- users who want a useful enhancement without changing the normal Zotero workflow

### Key Features

| Feature | Description |
| --- | --- |
| Numeric styles | Supports citations like `[1]` and `[2, 3]` |
| Author-date styles | Supports citations like `(Smith, 2024)` |
| Simple workflow | Use it directly from the `Zotero` tab in Word |
| Format-preserving | Mostly changes citation color without altering layout |
| Reversible | You can remove generated links and restore the original template |

### Support Matrix

| Platform | Status | Install Mode |
| --- | --- | --- |
| Windows + Word | Supported | One-click installer / script install |
| Mac + Word | Experimental | One-click `.command` install / manual install |

### Prerequisites

- `Zotero` is installed
- `Microsoft Word` is installed
- the standard `Zotero` tab is already visible in Word

### Installation

#### Windows

Recommended for most users:

1. Close `Word`
2. Download and run the [Windows Installer](https://github.com/FFFxueGawaine/zotero-word-citation-links/releases/latest/download/zotero-word-links-installer.exe)
3. Reopen `Word`
4. Open the `Zotero` tab and confirm these buttons are visible:
   - `Create Citation Links`
   - `Remove Citation Links`

If you prefer a more transparent/manual flow:

1. Download the [Windows Share Package](https://github.com/FFFxueGawaine/zotero-word-citation-links/releases/latest/download/zotero-word-links-share.zip)
2. Extract it
3. Close `Word`
4. Run `install.bat`

If you want a fully manual path with no installer and no install script, see:

[install/WINDOWS_MANUAL_INSTALL.md](./install/WINDOWS_MANUAL_INSTALL.md)

#### Mac

Mac support is currently experimental. The recommended path is the template package:

1. Download the [Mac Template Package](https://github.com/FFFxueGawaine/zotero-word-citation-links/releases/latest/download/zotero-word-links-mac-template.zip)
2. Quit `Word`
3. Extract the package and double-click:
   - `install_mac.command`
4. If macOS blocks it the first time, right-click the script and choose `Open`
5. Wait for the script to finish backup and install
6. Reopen `Word`
7. Open the `Zotero` tab and confirm these buttons are visible:
   - `Create Citation Links`
   - `Remove Citation Links`

Detailed guide:
[mac/MAC_INSTALL.md](./mac/MAC_INSTALL.md)

### Usage Tutorial

This is the core workflow.

#### Step 1: Write normally with Zotero

Use Zotero as you normally would:

1. insert in-text citations
2. generate the bibliography
3. edit, add, or refresh citations as needed

#### Step 2: Create jump links

Once your document already contains in-text citations and a bibliography:

1. open the `Zotero` tab in Word
2. click `Create Citation Links`
3. the citations become clickable
4. click a citation to jump to the matching bibliography entry

#### Step 3: Remove jump links when needed

If you want to remove the generated links:

1. click `Remove Citation Links`
2. the generated jumps are removed
3. citation color is restored as closely as possible to the pre-link state

### Button Guide

| Button | What It Does | When to Use It |
| --- | --- | --- |
| `Create Citation Links` | Creates clickable links for Zotero citations | After your citations and bibliography are already in place |
| `Remove Citation Links` | Removes the links created by this tool | When you want to restore normal display or recreate links |

### Recommended Workflow

For the most stable experience:

1. finish your normal Zotero editing first
2. if you just used `Zotero -> Refresh`, do not test links yet
3. run `Create Citation Links` again
4. then check the jump behavior

Why:

- `Zotero -> Refresh` rewrites citation results in Word
- so after a refresh, you will usually need to run `Create Citation Links` again

### Typical Output

Numeric styles:

- `[1]`
- `[2, 3]`

Author-date styles:

- `(Smith, 2024)`
- `(Kumar et al., 2026; Yu et al., 2025)`

Current design goals:

- numeric citations change color only
- author-date citations link only the inner text, while keeping the outer brackets normal
- removed links should not leave underline artifacts

### Restore / Rollback

If you want to go back to the pre-install state:

- Windows: run `restore_original.bat` from the share package
- Mac: run `restore_mac.command` from the Mac template package

### Known Limitations

- Zotero only, not EndNote
- numeric mode links the visible number token rather than the full bracket
- Mac support is still experimental and not fully validated across all Mac / Word versions
- reinstallation may be needed after Zotero updates

### Repository Layout

- `install/`
  Windows install scripts, restore script, macro module, manual install guide
- `mac/`
  Mac install documentation and support notes
- `tools/`
  build scripts
- `dist/`
  release assets

## License

MIT
