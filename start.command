#!/bin/bash
# =========================================
#  作业批改系统 - Mac 一键启动脚本
#  双击此文件即可启动
# =========================================

# 获取脚本所在目录
cd "$(dirname "$0")"

echo "========================================"
echo "   📚 作业批改系统 正在启动..."
echo "========================================"
echo ""

# 检查 Python
if ! command -v python3 &> /dev/null; then
    echo "❌ 未找到 Python3，请先安装 Python"
    echo "   下载地址：https://www.python.org/downloads/"
    echo ""
    read -p "按回车键退出..."
    exit 1
fi

# 检查虚拟环境，没有则创建
if [ ! -d ".venv" ]; then
    echo "📦 正在创建虚拟环境..."
    python3 -m venv .venv
fi

# 激活虚拟环境并安装依赖
source .venv/bin/activate
echo "📦 正在安装依赖（仅首次运行较慢）..."
pip install -r requirements.txt -q

echo ""
echo "✅ 启动成功！请在浏览器中打开："
echo ""
echo "   👉  http://localhost:8900"
echo ""
echo "========================================"

# 启动服务器
python3 app.py

echo ""
read -p "按回车键退出..."
