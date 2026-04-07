# Release Process

## 中文

从现在开始，GitHub Release 正文的唯一正式来源是：

- [CHANGELOG.md](../CHANGELOG.md)

这样做的原因很简单：

- `CHANGELOG.md` 是 UTF-8 文件
- [tools/sync_github_release_notes.py](../tools/sync_github_release_notes.py) 会按 UTF-8 读取正文
- GitHub API 提交正文时也会按 UTF-8 JSON 发送

这样可以避免中文在发布时被错误转换成 `?`。

### 标准发版步骤

1. 先在 [CHANGELOG.md](../CHANGELOG.md) 中补全新版本节  
   例如：

   ```md
   ## v0.4.2 - 2026-04-07
   ```

2. 正常打 tag、推送代码、创建 GitHub Release

3. 运行同步脚本，把对应版本节同步到 GitHub Release 正文：

   ```powershell
   python .\tools\sync_github_release_notes.py --repo FFFxueGawaine/zotero-word-citation-links --tag v0.4.2 --token-file D:\Claude\ZoteroWork\.github_token.txt
   ```

### 发布前推荐检查

先做一次预览：

```powershell
python .\tools\sync_github_release_notes.py --repo FFFxueGawaine/zotero-word-citation-links --tag v0.4.2 --dry-run
```

如果终端里提取出的版本节内容正常，再正式同步到 GitHub。

### 禁止做法

不要再用下面这种方式直接构造中文 Release 正文：

```powershell
@'
中文正文
'@ | python -
```

原因是这条链路可能会在文本进入 Python 之前就把中文转换成 `?`。

## English

From now on, the single source of truth for GitHub release notes is:

- [CHANGELOG.md](../CHANGELOG.md)

Why:

- `CHANGELOG.md` is stored as UTF-8
- [tools/sync_github_release_notes.py](../tools/sync_github_release_notes.py) reads it as UTF-8
- the GitHub API body is sent as UTF-8 JSON

This avoids the release-note corruption caused by piping Chinese text through a PowerShell here-string into `python -`.

### Standard Release Steps

1. Add the new version section to [CHANGELOG.md](../CHANGELOG.md)

   ```md
   ## v0.4.2 - 2026-04-07
   ```

2. Create the tag, push the code, and create the GitHub release as usual

3. Sync the matching changelog section into the GitHub release body:

   ```powershell
   python .\tools\sync_github_release_notes.py --repo FFFxueGawaine/zotero-word-citation-links --tag v0.4.2 --token-file D:\Claude\ZoteroWork\.github_token.txt
   ```

### Recommended Preview Step

Preview the extracted body first:

```powershell
python .\tools\sync_github_release_notes.py --repo FFFxueGawaine/zotero-word-citation-links --tag v0.4.2 --dry-run
```

If the extracted text looks correct, run the real sync command.

### Deprecated Path

Do not publish Chinese release notes through:

```powershell
@'
Chinese release body
'@ | python -
```

That path can turn Chinese text into `?` before the content even reaches Python or GitHub.
