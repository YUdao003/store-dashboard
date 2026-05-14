@echo off
chcp 65001 >nul
title 门店运营看板

:: 切换到脚本所在目录（dashboard.py 必须和本文件在同一文件夹）
cd /d "%~dp0"

echo.
echo  ========================================
echo   门店运营看板 正在启动...
echo   浏览器会自动打开，请稍候
echo   关闭此窗口 = 关闭看板
echo  ========================================
echo.

:: 检查 streamlit 是否已安装
streamlit --version >nul 2>&1
if errorlevel 1 (
    echo  [!] 未检测到 streamlit，正在安装依赖...
    pip install -r requirements.txt
    echo.
)

:: 启动看板（--server.headless false 确保自动弹出浏览器）
streamlit run dashboard.py --server.headless false --browser.gatherUsageStats false

pause
