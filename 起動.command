#!/bin/bash

# QC7つ道具システムの起動スクリプト

# ターミナルウィンドウを表示
osascript -e 'tell application "Terminal" to activate'

# スクリプトのディレクトリに移動
cd "$(dirname "$0")"

# ターミナルで起動スクリプトを実行
./起動.sh

