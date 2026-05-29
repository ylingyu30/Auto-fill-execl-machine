@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

REM =====================================================
REM Windows 一键打包脚本 (PyInstaller + venv)
REM 用法: build_exe.bat [你的Python脚本.py]
REM 示例: build_exe.bat fill_excel.py
REM =====================================================

set SCRIPT_NAME=%1
if "%SCRIPT_NAME%"=="" (
    echo 错误: 请指定要打包的 Python 脚本文件名
    echo 用法: %~nx0 your_script.py
    exit /b 1
)

if not exist "%SCRIPT_NAME%" (
    echo 错误: 文件 "%SCRIPT_NAME%" 不存在
    exit /b 1
)

echo ========================================
echo 开始打包: %SCRIPT_NAME%
echo ========================================

REM 1. 清理旧的构建产物 (可选)
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "*.spec" del /q *.spec

REM 2. 创建虚拟环境
if not exist "venv_pack" (
    echo 创建虚拟环境 venv_pack ...
    python -m venv venv_pack
    if errorlevel 1 (
        echo 创建虚拟环境失败，请检查 Python 是否安装
        exit /b 1
    )
) else (
    echo 虚拟环境已存在，跳过创建
)

REM 3. 激活虚拟环境并安装依赖
echo 激活虚拟环境并安装 PyInstaller + openpyxl ...
call venv_pack\Scripts\activate.bat
if errorlevel 1 (
    echo 激活虚拟环境失败
    exit /b 1
)

python -m pip install --upgrade pip
pip install pyinstaller openpyxl

REM 4. 执行打包命令
echo 正在使用 PyInstaller 打包...
pyinstaller --onefile ^
    --hidden-import=openpyxl ^
    --name="%~n1" ^
    --console ^
    "%SCRIPT_NAME%"

if errorlevel 1 (
    echo 打包失败！
    deactivate
    exit /b 1
)

REM 5. 清理临时文件（可选）
echo 打包完成！可执行文件位于: dist\%~n1.exe

REM 6. 退出虚拟环境
deactivate

echo ========================================
echo 成功生成: dist\%~n1.exe
echo ========================================
pause