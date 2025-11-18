#!/bin/bash
# PDF验证工具 - 打包脚本（Linux/Mac）

echo "======================================"
echo "PDF工资单验证工具 - 打包脚本"
echo "======================================"
echo ""

# 检查Python
if ! command -v python3 &> /dev/null; then
    echo "错误: 未找到Python3"
    exit 1
fi

echo "1. 创建虚拟环境..."
python3 -m venv venv
source venv/bin/activate

echo ""
echo "2. 安装依赖..."
pip install --upgrade pip
pip install -r requirements.txt

echo ""
echo "3. 使用PyInstaller打包..."
pyinstaller --onefile \
    --windowed \
    --name "PDF验证工具" \
    --icon=NONE \
    --add-data "requirements.txt:." \
    pdf_validator.py

echo ""
echo "======================================"
echo "打包完成！"
echo "可执行文件位于: dist/PDF验证工具"
echo "======================================"

# 清理
deactivate
