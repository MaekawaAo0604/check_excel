#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
文字化け診断スクリプト
このスクリプトは文字化けの原因を特定するためのものです。
"""

import sys
import os
import locale
import platform

def main():
    print("=" * 60)
    print("文字化け診断スクリプト")
    print("=" * 60)
    
    # システム情報
    print(f"OS: {platform.system()} {platform.release()}")
    print(f"Python: {sys.version}")
    print()
    
    # エンコーディング情報
    print("エンコーディング設定:")
    print(f"  sys.stdout.encoding: {sys.stdout.encoding}")
    print(f"  sys.stderr.encoding: {sys.stderr.encoding}")
    print(f"  sys.getdefaultencoding(): {sys.getdefaultencoding()}")
    print(f"  sys.getfilesystemencoding(): {sys.getfilesystemencoding()}")
    print()
    
    # ロケール情報
    print("ロケール設定:")
    try:
        current_locale = locale.getlocale()
        print(f"  locale.getlocale(): {current_locale}")
        print(f"  locale.getpreferredencoding(): {locale.getpreferredencoding()}")
    except Exception as e:
        print(f"  ロケール取得エラー: {e}")
    print()
    
    # 環境変数
    print("関連する環境変数:")
    env_vars = ['LANG', 'LC_ALL', 'LC_CTYPE', 'PYTHONIOENCODING', 'TERM']
    for var in env_vars:
        value = os.environ.get(var, '(未設定)')
        print(f"  {var}: {value}")
    print()
    
    # 日本語テスト
    print("日本語表示テスト:")
    test_strings = [
        "ひらがな: あいうえお",
        "カタカナ: アイウエオ",
        "漢字: 日本語表示確認",
        "記号: ！＠＃＄％",
        "混合: 生徒現状報告書チェッカー"
    ]
    
    for test_str in test_strings:
        print(f"  {test_str}")
    print()
    
    # 推奨設定
    print("推奨設定:")
    print("以下のコマンドを実行してから再度テストしてください:")
    print()
    print("  export LANG=ja_JP.UTF-8")
    print("  # または")
    print("  export LANG=C.UTF-8")
    print()
    print("  export PYTHONIOENCODING=utf-8")
    print()
    print("文字化けが発生している場合は、上記の設定を確認してください。")
    print("=" * 60)

if __name__ == "__main__":
    main()