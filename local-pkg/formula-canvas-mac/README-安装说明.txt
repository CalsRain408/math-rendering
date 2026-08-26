Formula Canvas —— Mac 本地安装说明
==================================================

这是把 LaTeX 公式渲染成高清矢量图、直接插入 Word / PowerPoint 的 Office 插件。
公式在你本机完成渲染，不上传任何内容。本版本完全在本地运行，无需登录。

【安装步骤】

1. 把整个 formula-canvas-mac 文件夹拖到一个固定位置
   （例如「文稿」或「应用程序」下，之后不要移动或删除它）。

2. 双击文件夹里的  install.command
   - 如果被拦截（提示「已阻止 install.command 以保护 Mac」
     或「Apple 无法验证 install.command…」）：
     打开「系统设置」→「隐私与安全性」，向下滚动到「安全性」区域，
     找到 已阻止 "install.command" 一行，点右侧的「仍要打开」，
     然后在弹窗里再点一次「打开」。
   - 安装过程中会要求输入你的 Mac 登录密码（用于信任本机证书），
     输入时屏幕不显示字符属正常，输完回车即可。
   - 若双击没反应，或终端报「no such file or directory」：
     打开「终端」App，输入  bash  和一个空格，
     然后把 install.command 文件直接拖进终端窗口（会自动填入完整路径），
     按回车即可。

3. 看到「安装完成！」后，完全退出 Word 和 PowerPoint（Cmd+Q）。

4. 重新打开 Word 或 PowerPoint，
   点菜单「开始」→「加载项」→ 选择 Formula Canvas。

【使用】

输入公式后实时预览，设置字号、颜色、行内/独立排版和背景，
把光标放到目标位置，点「插入」即可。
在 Word 的「插入格式」中选择「MathType 风格（独立公式自动编号）」：
「独立公式」会居中插入，并在右侧自动生成 (1)、(2)… 编号；
「行内公式」只在光标处紧凑插入，不添加前后空格或编号。
「数学字体」可选择 TeX 或 Times New Roman（MathType 风格）；Times 模式会用
Times New Roman 排印英文字母和数字，并保留专用数学符号来绘制复杂结构。
「字体样式」可选择数学默认、正体、斜体、粗体或粗斜体。
「完整字号（Full）」按 MathType 定义控制公式主字号，默认 12 pt；脚标为 58%，
二级脚标为 42%。积分、求和、根号等大符号会按结构自动伸展。
插入 Word 时使用真实 point 尺寸。
增删公式后，可在 Word 中按 Command+A，再按 F9 统一刷新编号。
插入快捷键：Control + Shift + I
清空输入框：Control + Shift + K

【卸载】

双击文件夹里的 uninstall.command，按提示操作即可。

【常见问题】

- 插件打不开 / 空白：
  确认 formula-canvas-mac 文件夹没有被移动或删除；
  重新双击 install.command 再试。

- 首次运行提示要安装「命令行开发者工具」：
  点「安装」，装完后重新双击 install.command。

- 服务是否在运行：
  浏览器打开 https://localhost:3000/ 能看到插件界面即正常。
