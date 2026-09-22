ExcelRouter · Excel 智能拆分工具   —— 麒麟 / UOS 版使用说明
====================================================================

怎么用
--------------------------------------------------------------------
1. 把 tar.gz 解压到你自己的目录（例如 ~/工具/）
       tar -xzf ExcelRouter-vX.Y.Z-linux-*.tar.gz
2. 进到解压出来的文件夹，运行：
       ./启动ExcelRouter.sh
   （也可以在文件管理器里双击它。如果双击没反应，右键 → 属性 →
     勾上「允许作为程序执行文件」，或者在终端里跑上面那行命令。）
3. 按界面上的三步走：① 选表格 → ② 选字段 → ③ 开始拆分

数据全程在你这台机器上处理，不联网、不上传。


常见问题
--------------------------------------------------------------------
Q: 报 "GLIBC_2.xx not found" 起不来？
A: 这个包是在另一台 glibc 更新的机器上打的。在你自己这台机器上用源码里的
   build_linux.sh 重新打一份即可（见项目 README 的「从源码打包」一节）。

Q: 界面里的中文是方块 / 豆腐块？
A: 系统缺中文字体：  sudo apt install -y fonts-wqy-zenhei

Q: PDF 水印里的中文变成了 ??? ？
A: 同样是缺中文字体。装上上面那个包即可。如果装了还不行，可以手动指定：
       ER_CJK_FONT=/usr/share/fonts/你的字体.ttf ./启动ExcelRouter.sh

Q: 拖拽文件进窗口没反应？
A: 个别机型的 tkdnd 原生库加载不了，日志里会有一行提示。不影响使用，
   点「📄 选一个 Excel 文件」/「📁 选整个文件夹」按钮即可。

Q: 点「打开输出文件夹」没反应？
A: 系统里没有 xdg-open：  sudo apt install -y xdg-utils

Q: 输入框里打不出中文？
A: 用 ./启动ExcelRouter.sh 启动（不要直接运行 ExcelRouter 这个文件）——
   启动脚本会设置输入法需要的 XMODIFIERS 环境变量。

更多问题见在线 FAQ：
   https://gitee.com/Marsandsea/Excelrouter/blob/main/docs/FAQ.md


--------------------------------------------------------------------
作者：AbeLin     开源协议：MIT     项目主页：
   https://gitee.com/Marsandsea/Excelrouter
   https://github.com/MarsandSea/excel-router
