#!/bin/bash

# エラー時に停止
set -e

echo "🚀 QC7つ道具システムを起動しています..."
echo ""

# カレントディレクトリを確認
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# 必要なライブラリがインストールされているか確認
echo "📦 ライブラリをチェック中..."
python -c "import streamlit" 2>/dev/null || {
    echo "❌ Streamlitがインストールされていません"
    echo "インストール中..."
    pip install streamlit pandas numpy plotly matplotlib seaborn scipy openpyxl xlrd python-pptx reportlab Pillow kaleido
}

echo ""
echo "✅ 起動準備完了！"
echo ""
echo "ブラウザで http://localhost:8501 にアクセスしてください"
echo ""
echo "⚠️  注意: このウィンドウを閉じるとシステムが停止します"
echo ""

# Streamlitを起動
streamlit run app.py --server.port=8501 --server.headless=true
