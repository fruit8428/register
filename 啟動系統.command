#!/bin/bash
DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" >/dev/null 2>&1 && pwd )"
cd "$DIR"

echo "=========================================================="
echo "  國際扶輪 3523 地區 2026-27 保齡球比賽 管理系統啟動器"
echo "=========================================================="
echo "正在啟動本地後端服務伺服器 (Port 5001)..."

# 確保虛擬環境 Python 存在
if [ -f "$DIR/.venv/bin/python" ]; then
    PYTHON_CMD="$DIR/.venv/bin/python"
else
    PYTHON_CMD="python3"
fi

# 延遲 1.5 秒自動在預設瀏覽器開啟管理頁面
(sleep 1.5 && open "http://127.0.0.1:5001") &

# 執行後端伺服器 (包含每 3 小時自動定時收信排程與即時對帳)
"$PYTHON_CMD" web_app.py
