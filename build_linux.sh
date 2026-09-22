#!/usr/bin/env bash
# ============================================================
# ExcelRouter - Linux 打包脚本（银河麒麟 Kylin V10 / 统信 UOS）
# Copyright (c) 2026 AbeLin · MIT License
#
# ★ 必须在【目标机器上】运行。
#   PyInstaller 产物的 glibc 下限 == 打包机的 glibc，这是唯一能保证 ABI 对得上的
#   办法，也是信创适配的标准做法。在 Ubuntu 22.04+（glibc 2.35+）上打出来的包，
#   拿到麒麟 V10（glibc 2.28~2.31）上会直接报 "GLIBC_2.34 not found"。
#   如果你有多台不同 SP 的麒麟，请在【最旧】的那台上打包。
#
# 用法：
#   ./build_linux.sh                    # 用 python3
#   PYTHON=python3.11 ./build_linux.sh  # 指定解释器
#
# 产物：dist/ExcelRouter-vX.Y.Z-linux-$(uname -m).tar.gz
#
# 与 build.bat 的差异（每一条都有原因，别照搬 Windows 的写法）：
#   --add-data 分隔符      ';' → ':'      PyInstaller 的平台约定
#   --icon app.ico         去掉           只对 .exe/.app 有意义，ELF 上没用；
#                                         Linux 窗口图标靠运行时 iconphoto(app.png)
#   --version-file         去掉           那是 Windows PE 版本资源，非 Windows 不适用
#   --noconsole            去掉           Linux 上是空操作，留着只会让人误以为它做了事
#   --onefile              ★ 绝不做       onefile 每次启动都解包到 /tmp/_MEIxxxx，
#                                         很多信创镜像把 /tmp 挂成 noexec → 直接起不来
#   --noupx                加上           麒麟一般没装 UPX，显式关掉保证可重现
# ============================================================
set -euo pipefail
cd "$(dirname "$0")"

APP=ExcelRouter
PY=${PYTHON:-python3}
ARCH=$(uname -m)                    # x86_64 / aarch64

echo "[0/5] 环境自检 ..."

# --- Python 版本：requirements.txt 钉了 pandas>=2.0.0，需要 3.9+。
#     麒麟 V10 SP1 自带的常常是 3.7，不先拦住的话后面 pip 会吐一屏天书。
"$PY" - <<'PYEOF' || exit 1
import sys
if sys.version_info < (3, 9):
    print(f"❌ {sys.executable} 是 Python {sys.version.split()[0]}，pandas 2.x 需要 3.9+")
    print("   解决办法（任选其一）：")
    print("     1) sudo apt install -y python3.9 python3.9-tk python3.9-dev")
    print("        然后：PYTHON=python3.9 ./build_linux.sh")
    print("     2) 装 miniforge/conda 建一个 3.11 环境，激活后重跑本脚本")
    print("   注意：这只影响【打包机】。目标机不需要装 Python，解释器打进包里了。")
    raise SystemExit(1)
PYEOF

# --- tkinter：麒麟默认不装 python3-tk，缺了会打出一个跑不起来的包
"$PY" -c "import tkinter" 2>/dev/null || {
  echo "❌ 缺少 tkinter（customtkinter 的底座）。"
  echo "   deb 系（麒麟桌面 / UOS）: sudo apt install -y python3-tk python3-dev"
  echo "   rpm 系（麒麟服务器）    : sudo yum install -y python3-tkinter python3-devel"
  exit 1; }

# --- 中文字体：不致命，但 PDF 水印会变 '?'，必须提前说
if ! fc-list :lang=zh 2>/dev/null | grep -q .; then
  echo "⚠  本机没有中文字体，PDF 水印里的中文会显示成 '?'"
  echo "   建议先装： sudo apt install -y fonts-wqy-zenhei   （或 fonts-noto-cjk）"
fi

VER=$("$PY" - <<'PYEOF'
import io, re
# 版本号唯一真源是 gui/app.py 的 APP_VERSION —— 不为 Linux 另造第四处版本定义
src = io.open("gui/app.py", encoding="utf-8").read()
print(re.search(r'APP_VERSION\s*=\s*"([^"]+)"', src).group(1))
PYEOF
)
OUT="${APP}-v${VER}-linux-${ARCH}"
GLIBC=$(getconf GNU_LIBC_VERSION | awk '{print $2}')

echo "   版本   : v$VER"
echo "   Python : $("$PY" -V 2>&1) / tk $("$PY" -c 'import tkinter; print(tkinter.TkVersion)')"
echo "   架构   : $ARCH     打包机 glibc: $GLIBC"

echo "[1/5] 安装依赖 ..."
"$PY" -m pip install --upgrade pip
"$PY" -m pip install -r requirements.txt
"$PY" -m pip install pyinstaller

echo "[2/5] 跑自测（不过就不打包）..."
"$PY" -m pytest -q

echo "[3/5] 清理旧产物 ..."
rm -rf build dist "${APP}.spec"

echo "[4/5] PyInstaller onedir ..."
"$PY" -m PyInstaller --onedir --noupx \
  --name "$APP" \
  --collect-all customtkinter \
  --collect-all tkinterdnd2 \
  --add-data "config:config" \
  --add-data "app.png:." \
  main.py

echo "[5/5] 生成启动脚本 + 打 tar.gz ..."
# 把打包机的 glibc 版本写进启动脚本，启动失败时好给用户一句人话
sed "s/@GLIBC@/${GLIBC}/g" packaging/launcher.sh.in > "dist/${APP}/启动ExcelRouter.sh"
chmod +x "dist/${APP}/启动ExcelRouter.sh" "dist/${APP}/${APP}"
cp packaging/README-Linux.txt "dist/${APP}/使用说明.txt"
mv "dist/${APP}" "dist/${OUT}"
tar -C dist -czf "dist/${OUT}.tar.gz" "${OUT}"

# ★ 整个脚本最有价值的一行：产物真实的 glibc 下限。
#   只有 glibc 不低于这个版本的机器才跑得起来。
FLOOR=$(objdump -T "dist/${OUT}/${APP}" "dist/${OUT}/_internal/"*.so* 2>/dev/null \
        | grep -oE 'GLIBC_2\.[0-9]+' | sort -uV | tail -1 || true)

echo "============================================================"
echo " ✅ 产物: dist/${OUT}.tar.gz"
echo " glibc 下限: ${FLOOR:-未检出}  （低于此版本的机器跑不起来）"
echo " 打包机 glibc: ${GLIBC}"
echo "------------------------------------------------------------"
echo " 冒烟自测: cd dist/${OUT} && ./启动ExcelRouter.sh"
echo "============================================================"
