#!/bin/bash
cd "$(dirname "$0")"
echo "正在检查环境..."
if ! python3 -c "import pandas" 2>/dev/null; then
    echo "首次运行，正在安装组件（约需1-2分钟）..."
    pip3 install pandas openpyxl msoffcrypto-tool pikepdf -q
    echo "✅ 组件安装完成"
fi
echo "正在启动程序..."
python3 main.py
echo "程序已退出，按回车关闭..."
read
