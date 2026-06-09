@echo off
chcp 65001 >nul
title 作业批改系统

echo ========================================
echo    📚 作业批改系统 正在启动...
echo ========================================
echo.

:: 检查 Python
where python >nul 2>nul
if %ERRORLEVEL% NEQ 0 (
    where python3 >nul 2>nul
    if %ERRORLEVEL% NEQ 0 (
        echo ❌ 未找到 Python，请先安装
        echo    下载地址：https://www.python.org/downloads/
        echo.
        pause
        exit /b
    )
)

:: 检测 Python 命令
set PYTHON_CMD=python
where python3 >nul 2>nul && set PYTHON_CMD=python3

:: 检查虚拟环境
if not exist ".venv" (
    echo 📦 正在创建虚拟环境...
    %PYTHON_CMD% -m venv .venv
)

:: 激活虚拟环境并安装依赖
call .venv\Scripts\activate.bat
echo 📦 正在安装依赖（仅首次运行较慢）...
%PYTHON_CMD% -m pip install -r requirements.txt -q

echo.
echo ✅ 启动成功！请在浏览器中打开：
echo.
echo    👉  http://localhost:8900
echo.
echo ========================================

:: 启动服务器
%PYTHON_CMD% app.py

pause
