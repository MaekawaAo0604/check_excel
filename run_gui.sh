#!/bin/bash
# WSL用エンコーディング環境設定
export LANG=C.UTF-8
export LC_ALL=C.UTF-8
export PYTHONIOENCODING=utf-8
export PYTHONLEGACYWINDOWSSTDIO=utf-8

# Tkinter用設定
export TK_LIBRARY=""
export TCL_LIBRARY=""

# X11転送設定（WSLGまたはX410等を使用する場合）
if [[ -n "$WSL_DISTRO_NAME" ]]; then
    export DISPLAY=:0.0
fi

# GUI版を実行
python3 excel_validator.py