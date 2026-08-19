@echo off
chcp 65001 >nul
title QCP转换工具
echo ========================================
echo   QCP → CNPE 转换工具
echo ========================================
echo.
echo 正在检查Python环境...
python --version >nul 2>&1
if errorlevel 1 (
    echo [错误] 未找到Python，请先安装Python 3
    echo 下载地址: https://www.python.org/downloads/windows/
    pause
    exit /b 1
)

echo 检查依赖库...
python -c "import pdfplumber" >nul 2>&1
if errorlevel 1 (
    echo 正在安装 pdfplumber...
    pip install pdfplumber -q
)

python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo 正在安装 openpyxl...
    pip install openpyxl -q
)

echo 依赖检查完成，正在启动...
echo.
python "%~dp0qcp_converter.py"
pause
