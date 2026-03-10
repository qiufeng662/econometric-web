@echo off
chcp 65001 >nul
title 一键实证分析
echo ========================================
echo    一键实证分析 正在启动...
echo ========================================
echo.
echo 请在浏览器中访问: http://localhost:8501
echo.
echo 按 Ctrl+C 停止服务
echo ========================================
echo.
streamlit run app.py --server.headless true --browser.gatherUsageStats false
pause