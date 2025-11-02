@echo off
chcp 65001 >nul
echo QC7つ道具システムを起動しています...
echo.

cd /d "%~dp0"

echo 📦 必要なライブラリを確認中...
pip install -q -r requirements.txt
if errorlevel 1 (
    echo ❌ ライブラリのインストールに失敗しました
    echo エラーを確認してください
    pause
    exit /b 1
)

echo.
echo ✅ 起動準備完了！
echo.
echo 🌐 ブラウザで http://localhost:8501 にアクセスしてください
echo.
echo ⚠️  このウィンドウを閉じるとシステムが停止します
echo.

streamlit run app.py --server.headless=true

pause
