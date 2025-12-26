# リポジトリ運用ガイドライン

## コミュニケーション
- ユーザーへの回答は丁寧な日本語で行ってください。

## プロジェクト概要
- レガシーのExcel VBAツール（`図面作成TOOL.xlam`）をPythonにリファクタリングするプロジェクトです。
- CODE V出力（.seq）を基に、光学設計データから図面を自動生成します。
- CADは図脳RAPIDをCOM経由で制御する想定です（`pywin32` 相当）。

## プロジェクト構成とモジュール配置
- `src/` にアプリケーション本体を配置します（エントリポイント: `src/main.py`）。
- `notebooks/` は検証・分析用ノートブックの置き場です。
- `data/` はローカル入力データ用です。大容量・機密データはコミットしないでください。
- `docs/` と `references/` は設計・調査ドキュメントの置き場です。
- `.github/` にコーディング規約とコミット指針があります。

## ビルド・テスト・開発コマンド
- `uv sync` 依存関係を `pyproject.toml` と `uv.lock` から同期します。
- `uv run python src/main.py` ローカル実行の基本コマンドです。
- `uv run pytest` テストを実行します。
- `uv run ruff check src tests` 静的チェックを実行します（`--fix` で自動修正）。
- `uv run ruff format src tests` フォーマットを統一します。

## コーディングスタイルと命名規則
- インデントは4スペース、標準的なPythonフォーマットを採用します。
- 命名は `PascalCase`（クラス）、`snake_case`（関数・変数）、`UPPER_CASE_SNAKE_CASE`（定数）。
- 公開モジュール/クラス/関数/メソッドには日本語のGoogleスタイルDocstringが必須です。
- Docstring/コメント/ドキュメントは日本語で記述します。
- すべての引数と戻り値に型ヒントを付与します（`|` のUnion記法を推奨）。
- デバッグ出力は `print` ではなく `logging` を使用します。
- 辞書より `@dataclass` を優先し、マジックナンバーは定数化します。
- 共有インターフェースは `abc.ABC` を用いて抽象化します。
- 不正な入力には適切な例外（`ValueError` / `TypeError` など）を送出します。

## テストガイドライン
- フレームワークは `pytest` を使用します。
- テストは `tests/` 配下に `test_*.py` 形式で追加します。
- 共有セットアップには `@pytest.fixture` を活用してください。
- VBA由来の計算ロジックは、VBAの出力と一致する数値検証テストを用意します。

## コミット・PRガイドライン
- コミットはConventional Commits準拠、本文は日本語で記述します。
- 例: `docs(readme): セットアップ手順を追記`
- PRには概要、実行したテスト、関連Issueを記載します。
- ノートブックや分析結果の変更はスクリーンショットや出力添付を推奨します。

## 設定・セキュリティ
- Pythonは `.python-version` で 3.12 系に固定されています。
- 秘密情報は環境変数やローカル設定に保存し、リポジトリに含めないでください。

## 重要資料
- `.github/copilot-instructions.md` に詳細なコーディング規約があります。
- `docs/design/VBA_Module_Reference.md` はVBAの全体構成を把握するための必読資料です。
