#!/bin/bash
# 双击即可运行（只用 Mac 自带的 Python 3，无需安装任何库）
# Double-click to run. Uses the Python 3 that ships with macOS; nothing to install.
cd "$(dirname "$0")"
echo "=== 品牌城市月销量 抓取工具 (Mac) ==="
python3 crawler_mac.py
echo ""
read -p "完成，按回车键关闭。 Finished - press Enter to close."
