@echo off
chcp 65001 >nul
title 论文格式检测工具 - CLI模式

echo ========================================
echo    论文格式检测工具 (命令行模式)
echo ========================================
echo.

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo [错误] 未找到虚拟环境
    pause
    exit /b 1
)

if "%~1"=="" (
    echo 用法: run_cli.bat "论文路径.docx" [配置文件路径]
    echo.
    echo 示例:
    echo   run_cli.bat "论文.docx"
    echo   run_cli.bat "论文.docx" "config\config_default.json"
    echo.
    pause
    exit /b 1
)

set DOC_PATH=%~1
set CONFIG_PATH=%~2

if "%CONFIG_PATH%"=="" (
    .venv\Scripts\python.exe main.py -i "%DOC_PATH%" --cli
) else (
    .venv\Scripts\python.exe main.py -i "%DOC_PATH%" -c "%CONFIG_PATH%" --cli
)

echo.
pause
