# 実装Plan（Issue 1-2: FocalLength関数）

## 目的
`calculate_focal_length()` を実装し、VBA `basUDF.FocalLength` と相対誤差 1e-7 以下で一致させる。

## 前提
- 仕様: `docs/design/Issue1-2-focal-length-function/requirements.md`
- 参照: `references/VBA/basUDF.bas`

## 実装ステップ
1. 参照コード確認
   - `basUDF.FocalLength` の引数・入力正規化・数式を確認する。
2. 定数定義
   - `src/optics/calculations.py` に `PLANE_RADIUS = 1e10` を定義する。
3. 関数実装
   - `calculate_focal_length()` を実装し、入力正規化（平面・屈折率下限）と無限大判定を含める。
   - **VBA互換性のため、例外は送出せず、入力値を正規化する。**
4. 型ヒント・Docstring整備
   - 日本語のGoogleスタイルDocstringを付与する。
5. 単体テスト実装
   - `tests/test_optics_calculations.py` に5観点以上のテストを追加する。
6. 数値一致検証
   - VBA値と比較し、相対誤差 1e-7 以下を満たす。
7. 静的チェック・フォーマット
   - `uv run ruff check src tests`
   - `uv run ruff format src tests`

## 完了条件
- `calculate_focal_length()` が仕様通りに動作する。
- 5観点以上のテストがパスする。
- VBA版と相対誤差 1e-7 以下で一致する。
