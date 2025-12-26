# 実装Plan（Issue 1-3: GlassWeight関数）

## 目的
`calculate_glass_weight()` を実装し、VBA `basUDF.GlassWeight` と相対誤差 1e-6 以下で一致させる。

## 前提
- 仕様: `docs/design/Issue-1-3-glass-weight-function/requirements.md`
- 参照: `references/VBA/basUDF.bas`
- 依存: `calculate_sag()` が利用可能であること

## 実装ステップ
1. 参照コード確認
   - `basUDF.GlassWeight` の引数、符号、積分式、外周補正を確認する。
2. サグ計算の前提整理
   - `calculate_sag()` の入出力と`None`返却条件を確認する。
3. 関数実装
   - `calculate_glass_weight()` を実装し、iMax=100固定の台形積分を再現する。
   - R1は正、R2は負の符号でサグ配列を構成する。
4. 異常系対応
   - サグ計算が`None`の場合は`None`を返す。
5. 型ヒント・Docstring整備
   - 日本語のGoogleスタイルDocstringを付与する。
6. 単体テスト実装
   - `tests/test_optics_calculations.py` に6観点以上のテストを追加する。
7. 数値一致検証
   - VBA値と比較し、相対誤差 1e-6 以下を満たす。
8. 静的チェック・フォーマット
   - `uv run ruff check src tests`
   - `uv run ruff format src tests`

## 完了条件
- `calculate_glass_weight()` が仕様通りに動作する。
- 6観点以上のテストがパスする。
- VBA版と相対誤差 1e-6 以下で一致する。
