"""
MCPサーバーのエントリーポイントとユーティリティコマンド
"""

import subprocess

# 品質チェック対象ディレクトリの定数
QUALITY_CHECK_DIRS = ["src"]


def lint():
    """
    ruffでコードの静的解析を実行する
    """
    cmd = ["ruff", "check"] + QUALITY_CHECK_DIRS
    subprocess.run(cmd)


def format():
    """
    ruffでコードフォーマットを実行する
    """
    cmd = ["ruff", "format"] + QUALITY_CHECK_DIRS
    subprocess.run(cmd)


def fix():
    """
    ruffで自動修正とフォーマットを一括実行する
    """
    print("🔧 コードの自動修正とフォーマットを実行中...")

    # 自動修正（安全な修正 + 危険な修正も含む）
    fix_cmd = ["ruff", "check", "--fix", "--unsafe-fixes"] + QUALITY_CHECK_DIRS
    fix_result = subprocess.run(fix_cmd)

    # フォーマット
    format_cmd = ["ruff", "format"] + QUALITY_CHECK_DIRS
    format_result = subprocess.run(format_cmd)

    if fix_result.returncode == 0 and format_result.returncode == 0:
        print("✅ 自動修正とフォーマットが完了しました")
    else:
        print("❌ 自動修正またはフォーマットでエラーが発生しました")

    exit(fix_result.returncode or format_result.returncode)


def type_check():
    """
    メイン型チェックを実行する (ty - 高速)
    """
    cmd = ["ty", "check"] + QUALITY_CHECK_DIRS
    subprocess.run(cmd)




def check():
    """
    型チェック、Lintを同時に実行する（修正はしない）
    """
    print("🔍 型チェックを実行中...")
    type_result = subprocess.run(["ty", "check"] + QUALITY_CHECK_DIRS, capture_output=True)

    print("📝 Lintを実行中...")
    lint_result = subprocess.run(["ruff", "check"] + QUALITY_CHECK_DIRS, capture_output=True)

    # 結果をまとめて表示
    print("\n" + "=" * 50)
    print("📊 実行結果サマリー")
    print("=" * 50)

    type_status = "✅ PASS" if type_result.returncode == 0 else "❌ FAIL"
    lint_status = "✅ PASS" if lint_result.returncode == 0 else "❌ FAIL"

    print(f"型チェック: {type_status}")
    print(f"Lint:       {lint_status}")

    # エラーがある場合は詳細を表示
    if type_result.returncode != 0:
        print("\n🔍 型チェックエラー:")
        print(type_result.stdout.decode())
        print(type_result.stderr.decode())

    if lint_result.returncode != 0:
        print("\n📝 Lintエラー:")
        print(lint_result.stdout.decode())
        print(lint_result.stderr.decode())

    # いずれかが失敗した場合は非ゼロで終了
    if any(result.returncode != 0 for result in [type_result, lint_result]):
        exit(1)
    else:
        print("\n🎉 すべてのチェックが成功しました！")
        exit(0)