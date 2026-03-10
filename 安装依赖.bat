@echo off
chcp 65001 >nul
title 安装依赖
echo ========================================
echo    一键实证分析 - 安装依赖
echo ========================================
echo.
echo 正在安装所需依赖，请稍候...
echo.
pip install -r requirements.txt
echo.
echo ========================================
if %errorlevel% equ 0 (
    echo 安装成功！
) else (
    echo 安装失败，请检查是否安装了 Python
)
echo ========================================
echo.
pause