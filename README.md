# Formula Canvas

把 LaTeX / KaTeX 风格公式渲染为纯 SVG，并直接插入 Microsoft Word 或 PowerPoint。支持 SVG 的 Office 会得到可无限缩放的矢量图；不支持 SVG 的版本会自动回退到最高 8 倍分辨率的透明 PNG。也可以在“插入格式”中手动选择输出方式。

在 **Word** 中还可以选择“可编辑公式（OMML）”，把公式插入为 Word 原生的公式对象，插入后可以直接用鼠标点进去编辑。此模式通过 `LaTeX → MathML → OMML` 转换实现，仅在 Word 宿主下出现该选项（PowerPoint 无原生公式引擎，继续使用 SVG / PNG 图片）。

如果希望避开 Word 原生公式固定使用的 Cambria Math，可以选择 **“MathType 风格（独立公式自动编号）”**。该模式会严格遵循“排版”设置：选择“独立公式”时插入一个无边框的三栏版式，公式在整行严格居中，编号在右侧显示为 `(1)`、`(2)`……；插入前会读取当前节的实际页宽、页边距、装订线、分栏宽度及所在表格单元格宽度，两侧只保留固定编号槽，中间公式使用全部剩余宽度，超宽时按原始宽高比整体缩小而不会裁切。选择“行内公式”时只在当前光标处紧凑插入公式，不创建表格、不添加编号，也不会在公式前后主动加入空格。行内公式还会读取 MathJax 为每条公式计算的真实数学基线，并按下标、分式等结构的下深度自动下移，避免带下标公式相对正文整体抬高；该校正随字号线性缩放，不使用固定偏移。独立公式的编号是 Word 的 `SEQ` 域而不是写死的文本，增删公式后可用 `Command/Ctrl + A`、`F9` 统一刷新；每个编号还带有独立书签，可在 Word 中用于交叉引用。旧版 Word 会自动使用内嵌的高清 PNG 作为显示后备。

“数学字体”目前提供 **TeX 数学字体** 和 **Times New Roman（MathType 风格）** 两种选择。Times 预设使用 Times New Roman 排印英文字母和数字，同时保留专用数学字形来绘制积分号、根号、伸展括号等结构；预览、SVG、PNG 和 Word 自动编号模式使用同一套字体设置。由于 Times New Roman 本身不是完整的 OpenType 数学字体，这种混合方式也更接近传统 MathType 的排版方法。

Times New Roman 的斜体 `f`、大写字母和上下标可能带有超出 MathJax 原始 TeX 字符框的墨迹。插件会在渲染后测量整条公式的实际边界并只扩展溢出的 SVG 方向，避免插入 Word 后笔画被裁切；这项画布校正不会在公式前后加入文本空格。

“字体样式”提供 **数学默认（变量斜体）**、**正体**、**斜体**、**粗体** 和 **粗斜体**。样式会应用到预览以及 SVG、PNG、Word 自动编号等固定排版输出；Word 原生可编辑公式仍由 Word 的 OMML 引擎决定字体和样式。

字号采用 MathType 的 typesize 定义：主字号显示为 **完整字号（Full）**，默认 12 pt；脚标为 Full 的 58%，二级脚标为 42%。积分、求和、根号等大符号仍根据公式结构自动伸展。插件会把 MathJax 的 `ex` 标尺换算为真实 point，并以 point 向 Word 传递图片尺寸，因此 10 pt、12 pt、18 pt 会线性改变公式的物理尺寸，不改变字重。字号输入支持 0.5 pt 步进，并提供常用预设。

> 说明：可编辑公式跟随 Word 文档的公式字体与字号，颜色、字号、白底等自定义样式在该模式下不生效；极少数非标准 LaTeX 命令可能转换不完整，此时会自动回退到图片插入。需要像素级精确排版时请使用 SVG / PNG。

> “MathType 风格”模式优先保证字体外观、居中和编号版式，公式以矢量图保存，不能像 OMML 一样在 Word 公式编辑器中逐字符修改。需要修改时，在 Word 中选中由新版 Formula Canvas 插入的公式，点击“编辑选中公式”即可恢复原 LaTeX 和排版设置；更新后会在原位置替换公式。旧版本插入、未携带源码元数据的公式仍需重新插入一次。

## macOS 安装包

[下载 Formula Canvas for macOS v0.1.0](https://github.com/CalsRain408/math-rendering/releases/download/v0.1.0/formula-canvas-mac.zip)

> 此一键安装包仅适用于 macOS。下载并解压后，请按照包内的 `README-安装说明.txt` 完成安装。

## 本地运行

需要 Node.js 20.19+。

```bash
npm install
npm run certs
npm run dev
```

首次执行 `npm run certs` 时，系统会请求允许安装 Microsoft Office Add-in 本地开发证书。随后在 Word 或 PowerPoint 中旁加载项目根目录的 `manifest.xml`。

### macOS 旁加载

1. 保持 `npm run dev` 运行。
2. 将 `manifest.xml` 复制到：
   - Word：`~/Library/Containers/com.microsoft.Word/Data/Documents/wef/`
   - PowerPoint：`~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/`
3. 完全退出并重新打开对应 Office 应用。
4. 在“开始”→“加载项”中打开 Formula Canvas。

### Windows 旁加载

通过 Office 的共享文件夹加载项目根目录的 `manifest.xml`，或使用 Microsoft 365 管理中心部署。详细步骤见 Microsoft 官方的 Office Add-in 旁加载文档。

## 使用

输入公式后实时预览，设置字号、颜色、行内/独立排版和背景，然后把光标放到 Word 的目标位置或打开目标幻灯片，点击“插入为矢量公式”。可以只输入公式主体；粘贴带 `$...$`、`$$...$$`、`\(...\)` 或 `\[...\]` 分隔符的内容时，插件也会自动移除最外层分隔符。下标应写为 `_i`，而 `\_i` 会按照 LaTeX 语义显示一个普通下划线后再显示 `i`。插入快捷键为 `Control + Shift + I`；清空输入框可点击“清空”，或按 `Control + Shift + K`。

公式渲染完全在本机任务窗格内完成，不上传内容。
