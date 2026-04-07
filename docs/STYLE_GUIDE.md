# Zotero Citation Link Style Guide

Jump to: [中文](#zh-cn) | [English](#en)

<a id="zh-cn"></a>

## 中文

[Switch to English](#en)

### 这份文档是干什么的

从 `v0.4.2` 开始，这个项目不再单独用按钮控制链接颜色，而是把链接外观交给当前文档中的字符样式：

- `Zotero Citation Link`

这意味着：

1. 链接看起来是什么样，主要由这个样式决定
2. 你改的是 **当前文档** 的样式，不是全局 Word 默认值
3. 改完样式后，重新执行一次 `Create Citation Links`，新建链接就会按新样式显示

### 默认效果

第一次在某篇文档中执行 `Create Citation Links` 时，如果文档里还没有这个样式，工具会自动创建它。

当前默认外观是：

- 蓝色
- 下划线

这样做的好处是最直观：

- 用户一眼就能看出哪里是可点击链接
- 看起来更像普通超链接
- 对大多数论文草稿阶段更容易检查跳转是否生成成功

### 这个样式会影响哪里

#### 数字编号格式

例如：

- `[1]`
- `[2, 3]`
- `[4]-[7]`

当前规则是：

- 样式主要作用在数字本体
- 方括号本身通常保持普通正文外观

#### 作者-年份格式

例如：

- `(Smith, 2024)`
- `(Kumar et al., 2026; Yu et al., 2025)`

当前规则是：

- 样式只作用在括号内部正文
- 左右括号保持普通正文样式

### 怎么打开样式窗格

最稳的方法有两种。

#### 方法一：从“开始”选项卡打开

1. 打开 Word
2. 切到 `开始`
3. 在“样式”区域右下角点击小箭头
4. 右侧会出现样式窗格

#### 方法二：快捷键

直接按：

```text
Alt + Ctrl + Shift + S
```

### 怎么找到 `Zotero Citation Link`

如果你在样式窗格里没有第一眼看到它，不要着急，按下面顺序来：

1. 先确认你已经至少执行过一次 `Create Citation Links`
2. 打开样式窗格
3. 在样式列表中搜索：

```text
Zotero Citation Link
```

如果还是看不到，通常是因为：

- 当前文档还没创建过链接
- 样式列表筛选太严格

这时可以先点一次 `Create Citation Links`，再回来看样式。

### 推荐怎么改

右键 `Zotero Citation Link`，然后点：

- `Modify...`

你最常改的通常是这些：

1. 字体
   好处：和正文或期刊模板统一

2. 字号
   好处：避免链接看起来过大或过小

3. 颜色
   好处：可以改成黑色、深蓝、灰色，符合不同投稿风格

4. 下划线
   好处：如果你不想太像网页链接，可以去掉下划线

5. 粗体 / 斜体
   好处：一般不推荐乱改，但有些模板会需要

6. 上标 / 下标
   好处：某些数字引用风格可能会用到

### 最推荐的改法

如果你刚开始用，我建议按这 3 个等级来。

#### 一级：最稳

- 保持蓝色
- 保留下划线

好处：

- 最容易识别
- 最容易检查有没有生成成功

#### 二级：更论文化

- 改成深蓝或黑色
- 去掉下划线

好处：

- 看起来更像正式论文
- 不会太像网页超链接

#### 三级：完全跟随模板

- 让字体、字号、颜色都尽量贴近你的正文模板

好处：

- 视觉最统一
- 最适合定稿阶段

### 改完后什么时候生效

这里要分两种情况。

#### 情况一：你只是刚改完样式

这时建议再点一次：

- `Create Citation Links`

好处：

- 让当前文档中的链接统一按新样式重建

#### 情况二：你之后又点了 Zotero 的 `Refresh`

现在项目已经做成：

- `Refresh` 完成后会自动重建链接

所以如果你已经改好样式，后面正常使用 `Refresh` 时，新建出来的链接也会继续跟这个样式走。

### 删除链接时会怎样

点 `Remove Citation Links` 后：

1. 超链接会被移除
2. 工具会尽量恢复到创建前的原始字符格式
3. 不会强行把 `Zotero Citation Link` 样式残留在普通正文上

这一步的好处是：

- 你可以安全重建
- 不会越点越乱

### 常见问题

#### 1. 我改了样式，但当前文档里的旧链接没立刻变

先再点一次：

- `Create Citation Links`

因为当前流程是“安全重建”，重建后效果最稳定。

#### 2. 我想让所有新文档都默认用这个样式

当前不是全局样式方案，而是：

- **当前文档级字符样式**

好处：

- 不会污染你其他 Word 文档
- 每篇论文可以按自己的模板单独调

#### 3. 我删掉链接后，样式还在文档里吗

字符样式本身一般还会留在文档中，但普通正文不会继续被强行套用它。  
这属于正常现象，不是错误。

### 一句话建议

如果你现在还在写作阶段，先保留默认蓝色加下划线最省心。  
如果你已经进入定稿阶段，再把 `Zotero Citation Link` 改成更接近论文模板的样子。

<a id="en"></a>

## English

[切换到中文](#zh-cn)

### What This Guide Is For

Starting with `v0.4.2`, this project no longer uses a dedicated button to control link color. Link appearance is now controlled by a current-document character style:

- `Zotero Citation Link`

That means:

1. link appearance is primarily controlled by this style
2. the style belongs to the **current document**, not to Word globally
3. after changing the style, run `Create Citation Links` again so rebuilt links follow the updated style

### Default Appearance

When you run `Create Citation Links` for the first time in a document, the tool will create the style automatically if it does not already exist.

The current default is:

- blue text
- underline

This is intentional because it makes the link state obvious:

- users can immediately see what is clickable
- it looks like a normal hyperlink
- it is easier to verify during drafting

### What the Style Affects

#### Numeric citations

Examples:

- `[1]`
- `[2, 3]`
- `[4]-[7]`

Current behavior:

- the style mainly applies to the numeric part
- the square brackets usually remain normal body text

#### Author-date citations

Examples:

- `(Smith, 2024)`
- `(Kumar et al., 2026; Yu et al., 2025)`

Current behavior:

- the style applies only to the inner citation text
- the outer parentheses remain normal body text

### How to Open the Styles Pane

Two reliable ways:

#### Option 1: From the Home tab

1. Open Word
2. Go to `Home`
3. In the Styles area, click the small launcher arrow
4. The Styles pane opens on the right

#### Option 2: Keyboard shortcut

Press:

```text
Alt + Ctrl + Shift + S
```

### How to Find `Zotero Citation Link`

If you do not see it immediately:

1. make sure you have already run `Create Citation Links` at least once
2. open the Styles pane
3. search for:

```text
Zotero Citation Link
```

If it still does not show up, the most common reasons are:

- links have not been created in this document yet
- the Styles pane is filtered too aggressively

Run `Create Citation Links` once first, then look again.

### What to Edit

Right-click `Zotero Citation Link`, then choose:

- `Modify...`

The most common edits are:

1. Font
2. Size
3. Color
4. Underline
5. Bold / italic
6. Superscript / subscript

### Recommended Levels

#### Level 1: safest

- keep blue
- keep underline

Benefit:

- easiest to inspect
- easiest to confirm that links were created successfully

#### Level 2: more paper-like

- change to dark blue or black
- remove the underline

Benefit:

- looks closer to a formal manuscript
- less like a web hyperlink

#### Level 3: fully template-matched

- make font, size, and color follow your manuscript style exactly

Benefit:

- visually consistent
- ideal for final polishing

### When Changes Take Effect

#### Case 1: you just changed the style

Run:

- `Create Citation Links`

This rebuilds links so they all pick up the new style reliably.

#### Case 2: you later click Zotero `Refresh`

The project now rebuilds citation links automatically after `Refresh`, so the updated style continues to apply in the normal workflow.

### What Happens on Remove

When you click `Remove Citation Links`:

1. hyperlinks are removed
2. the tool tries to restore the original character formatting from before link creation
3. the link style is not intentionally left behind on plain citation text

### Common Questions

#### 1. I changed the style, but existing links did not visibly update right away

Run:

- `Create Citation Links`

The rebuild path is the most reliable way to make the style change visible everywhere.

#### 2. Can I make this the default for every new document?

Not currently. The design is intentionally:

- **current-document character style**

Benefit:

- it does not pollute unrelated Word documents
- each paper can keep its own formatting

#### 3. Does the style stay in the document after I remove links?

Usually yes. The style definition may remain in the document, but normal text will not keep being forced into that style. That is expected behavior.

### One-line Recommendation

Keep the default blue + underline while drafting.  
Switch `Zotero Citation Link` to something closer to your manuscript template during final polishing.
