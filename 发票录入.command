#!/bin/zsh
# 发票登记表自动录入 —— 双击运行即可
# 读取「发票待整理」文件夹里的电子发票 PDF，抽取信息写入
# 发票登记表模板，生成 发票登记表-YYYYMMDD.xlsx
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
python3 "$SCRIPT_DIR/提取发票.py"
echo ""
read "?按回车键关闭窗口..."
