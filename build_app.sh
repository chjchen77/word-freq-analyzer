#!/bin/bash
# ============================================================
# 中文文本词频统计分析工具 — macOS / Windows 打包脚本
# ============================================================
# 使用 PyInstaller 将 Python 脚本打包为独立可执行程序。
# 打包后的 app 不需要安装 Python 即可运行。
#
# 用法：
#   chmod +x build_app.sh
#   ./build_app.sh
#
# 前置条件：
#   pip install -r requirements-build.txt
# ============================================================

set -e

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

echo "================================================"
echo "  中文文本词频统计分析工具 — 打包开始"
echo "================================================"

# 检查 PyInstaller
if ! python3 -c 'import PyInstaller' >/dev/null 2>&1; then
    echo "正在安装 PyInstaller..."
    python3 -m pip install -r requirements-build.txt
fi

# 检查依赖
echo "检查依赖..."
python3 -m pip install -q -r requirements-build.txt

# 清理旧构建
rm -rf build dist
rm -f *.spec 2>/dev/null || true

echo "开始打包..."

# 打包为目录包（onedir）：启动时无需解压，速度远快于 onefile
python3 -m PyInstaller \
    --onedir \
    --windowed \
    --name "词频统计分析工具" \
    --hidden-import llm_sentence_analyzer \
    --hidden-import jieba \
    --hidden-import openpyxl \
    --hidden-import xlrd \
    --hidden-import openai \
    --hidden-import xlsxwriter \
    --hidden-import jieba.finalseg \
    --hidden-import jieba.posseg \
    --hidden-import jieba.analyse \
    --collect-data jieba \
    --collect-data openpyxl \
    --collect-submodules keyring.backends \
    --exclude-module torch \
    --exclude-module tensorflow \
    --exclude-module matplotlib \
    --exclude-module scipy \
    --exclude-module numpy.distutils \
    --noupx \
    word_freq_analyzer.py

echo ""
echo "================================================"
echo "  打包完成！"
echo "================================================"

if [ -d "dist/词频统计分析工具.app" ]; then
    echo "macOS 应用: dist/词频统计分析工具.app"
    echo ""
    echo "使用方式："
    echo "  1. 双击 dist/词频统计分析工具.app 即可运行"
    echo "  2. 或将 .app 拖入「应用程序」文件夹"
    echo ""
    echo "分发方式："
    echo "  将 dist/词频统计分析工具.app 压缩为 .zip 发送给他人"
elif [ -f "dist/词频统计分析工具" ]; then
    echo "可执行文件: dist/词频统计分析工具"
    echo ""
    echo "使用方式："
    echo "  ./dist/词频统计分析工具"
fi

# 清理中间文件
rm -rf build
rm -f *.spec 2>/dev/null || true

echo ""
echo "Done!"
