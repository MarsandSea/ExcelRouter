#!/bin/sh
# ============================================================
# ExcelRouter 注册到开始菜单 / 桌面（银河麒麟 / 统信 UOS）
# Copyright (c) 2026 AbeLin · MIT License
#
# 跑一次就够。之后就能在「开始菜单 → 办公」里找到 ExcelRouter，
# 也可以右键固定到任务栏，跟普通软件一样用，不必再回到这个文件夹。
#
# 为什么需要这个脚本，而不能在 tar.gz 里直接放一个 .desktop：
#   .desktop 的 Exec= 必须写**绝对路径**，而用户把包解压到哪儿是未知的
#   （~/下载、~/工具、U 盘……）。所以只能在用户机器上、解压之后，
#   由脚本把当前真实路径写进去。
#   另外多数桌面环境不信任「放在随便某个文件夹里」的 .desktop，
#   只认 ~/.local/share/applications 下的——这也只能运行时才能装。
# ============================================================
set -u

# 双击运行时 Terminal=false，没有终端，echo 出去的话谁也看不见——
# 用户会以为「点了没反应」。所以结果一律再用图形对话框说一遍。
say() {
  echo "$1"
  if command -v zenity >/dev/null 2>&1; then
    zenity --info --no-wrap --title="ExcelRouter" --text="$1" >/dev/null 2>&1
  elif command -v kdialog >/dev/null 2>&1; then
    kdialog --title "ExcelRouter" --msgbox "$1" >/dev/null 2>&1
  elif command -v yad >/dev/null 2>&1; then
    yad --info --title="ExcelRouter" --text="$1" >/dev/null 2>&1
  else
    # 连对话框工具都没有：退回终端等待，至少从终端里跑时看得到
    printf "  按回车关闭..."
    read -r _ 2>/dev/null || true
  fi
}

DIR=$(cd "$(dirname "$0")" && pwd)
APPS="$HOME/.local/share/applications"
DESKTOP_FILE="$APPS/excelrouter.desktop"

# 解压工具经常把可执行位丢掉，顺手补回来（否则菜单项点了没反应）
chmod +x "$DIR/ExcelRouter" "$DIR/启动ExcelRouter.sh" 2>/dev/null

mkdir -p "$APPS" || {
  say "❌ 建不了 $APPS，没能注册到开始菜单。

仍然可以直接双击这个文件夹里的「启动ExcelRouter.sh」使用。"
  exit 1
}

# StartupWMClass 必须是 ExcelRouter：UKUI/GNOME 靠它把已经打开的窗口
# 和菜单项对应起来，缺了会在任务栏上多出一个「未知程序」的重复图标。
cat > "$DESKTOP_FILE" <<EOF
[Desktop Entry]
Type=Application
Version=1.0
Name=ExcelRouter
Name[zh_CN]=ExcelRouter
GenericName=Excel 智能拆分工具
GenericName[zh_CN]=Excel 智能拆分工具
Comment=按部门、区域、工号等字段批量拆分 Excel，打包分发
Comment[zh_CN]=按部门、区域、工号等字段批量拆分 Excel，打包分发
Exec="$DIR/启动ExcelRouter.sh" %f
Icon=$DIR/_internal/app.png
Terminal=false
Categories=Office;Spreadsheet;Utility;
Keywords=excel;split;拆分;表格;分发;
StartupWMClass=ExcelRouter
StartupNotify=true
EOF

chmod +x "$DESKTOP_FILE" 2>/dev/null
update-desktop-database "$APPS" >/dev/null 2>&1 || true

# 顺带在桌面上放一份，最接近 Windows 的双击体验。
# 桌面目录名随语言环境变化，两种都试。
# metadata::trusted 是关键：多数桌面默认不信任 .desktop，不标记的话
# 图标会显示成一个问号、双击弹「不受信任的应用程序启动器」。
ON_DESKTOP=""
for d in "$HOME/桌面" "$HOME/Desktop"; do
  if [ -d "$d" ]; then
    if cp -f "$DESKTOP_FILE" "$d/ExcelRouter.desktop" 2>/dev/null; then
      chmod +x "$d/ExcelRouter.desktop" 2>/dev/null
      gio set "$d/ExcelRouter.desktop" metadata::trusted true 2>/dev/null
      ON_DESKTOP="
✅ 桌面上也放了一个图标，直接双击就能用"
    fi
    break
  fi
done

say "✅ ExcelRouter 已添加到开始菜单

在「开始菜单」里搜「ExcelRouter」就能打开，
也可以在菜单项上右键 →「固定到任务栏」。$ON_DESKTOP

以后不用再回到这个文件夹了。
（想卸载：删掉本文件夹即可，程序没有装进系统）"
