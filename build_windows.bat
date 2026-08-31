@echo off
chcp 65001 >nul 2>&1
REM ============================================================
REM 中文文本词频统计分析工具 - Windows 打包脚本
REM ============================================================
REM 使用 PyInstaller 将 Python 脚本打包为独立可执行程序。
REM 打包后的 .exe 不需要安装 Python 即可运行。
REM
REM 前置条件：
REM   1. 安装 Python 3.10 或更高版本
REM   2. 本脚本会自动安装 requirements.txt 和 PyInstaller
REM
REM 用法：双击此文件 或 在命令行中运行 build_windows.bat
REM ============================================================

echo ================================================
echo   中文文本词频统计分析工具 - Windows 打包
echo ================================================

REM 检查 Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到 Python，请先安装 Python 3.10+
    echo 下载地址: https://www.python.org/downloads/
    pause
    exit /b 1
)

REM 检查并安装依赖
echo 检查依赖...
python -m pip install -q -r requirements.txt pyinstaller

REM 清理旧构建
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist *.spec del /f /q *.spec

echo 开始打包...

python -m PyInstaller ^
    --noconfirm ^
    --onedir ^
    --windowed ^
    --name "词频统计分析工具" ^
    --hidden-import llm_sentence_analyzer ^
    --hidden-import jieba ^
    --hidden-import openpyxl ^
    --hidden-import xlrd ^
    --hidden-import openai ^
    --hidden-import xlsxwriter ^
    --hidden-import jieba.finalseg ^
    --hidden-import jieba.posseg ^
    --hidden-import jieba.analyse ^
    --collect-data jieba ^
    --collect-data openpyxl ^
    --exclude-module torch ^
    --exclude-module tensorflow ^
    --exclude-module matplotlib ^
    --exclude-module scipy ^
    --exclude-module numpy.distutils ^
    --noupx ^
    word_freq_analyzer.py

if errorlevel 1 (
    echo.
    echo [错误] 打包失败，请检查上方错误信息。
    pause
    exit /b 1
)

REM 清理中间文件
if exist build rmdir /s /q build
if exist *.spec del /f /q *.spec

echo.
echo ================================================
echo   打包完成！
echo ================================================
echo.
echo 程序目录: dist\词频统计分析工具\
echo.
echo 使用方式:
echo   双击 dist\词频统计分析工具\词频统计分析工具.exe 即可运行
echo.
echo 分发方式:
echo   将 dist\词频统计分析工具\ 整个文件夹压缩为 .zip 发送给他人
echo.
pause
