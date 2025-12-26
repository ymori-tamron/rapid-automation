# 実装Plan（Issue 1-1: Sag関数）

## 目的
`calculate_sag()` と `AsphericCoefficients` を実装し、VBA `basUDF.Sag` と相対誤差 1e-7 以下で一致させる。

## 前提
- 仕様: `docs/design/Issue1-1-sag-function/requirements.md`
- 参照: `references/VBA/basUDF.bas`

## 実装ステップ
1. 参照コード確認
   - `basUDF.Sag` の引数・分岐・数式を確認する。
2. データクラス実装
   - `src/optics/calculations.py` に `AsphericCoefficients` を追加する。
3. 関数実装
   - `calculate_sag()` を実装し、平面判定・平方根の負値チェックを含める。
4. 型ヒント・Docstring整備
   - 日本語のGoogleスタイルDocstringを付与する。
5. 単体テスト実装
   - `tests/test_optics_calculations.py` に7観点以上のテストを追加する。
6. 数値一致検証
   - VBA値と比較し、相対誤差 1e-7 以下を満たす。
7. 静的チェック・フォーマット
   - `uv run ruff check src tests`
   - `uv run ruff format src tests`

## 完了条件
- `calculate_sag()` が仕様通りに動作する。
- 7観点以上のテストがパスする。
- VBA版と相対誤差 1e-7 以下で一致する。
