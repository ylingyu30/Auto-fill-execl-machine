#!/bin/bash

set -e  # 遇到错误即退出

SCRIPT_NAME="$1"
if [ -z "$SCRIPT_NAME" ]; then
    echo "错误: 请指定要打包的 Python 脚本文件名"
    echo "用法: $0 your_script.py"
    exit 1
fi

if [ ! -f "$SCRIPT_NAME" ]; then
    echo "错误: 文件 '$SCRIPT_NAME' 不存在"
    exit 1
fi

echo "========================================"
echo "开始打包: $SCRIPT_NAME"
echo "========================================"

# 1. 清理旧的构建产物
rm -rf build dist *.spec 2>/dev/null || true

# 2. 创建虚拟环境
if [ ! -d "venv_pack" ]; then
    echo "创建虚拟环境 venv_pack ..."
    python3 -m venv venv_pack
else
    echo "虚拟环境已存在，跳过创建"
fi

# 3. 激活虚拟环境并安装依赖
echo "激活虚拟环境并安装 PyInstaller + openpyxl ..."
source venv_pack/bin/activate

pip install --upgrade pip
pip install pyinstaller openpyxl

# 4. 执行打包命令
echo "正在使用 PyInstaller 打包..."
BASENAME=$(basename "$SCRIPT_NAME" .py)
pyinstaller --onefile \
    --hidden-import=openpyxl \
    --name="$BASENAME" \
    --console \
    "$SCRIPT_NAME"

# 5. 完成
deactivate

echo "========================================"
echo "成功生成: dist/$BASENAME"
echo "========================================"