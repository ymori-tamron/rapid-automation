# Taskリスト（Issue 1-3: GlassWeight関数）

## 実装タスク
- [x] `references/VBA/basUDF.bas` の `GlassWeight` 実装詳細を確認する
- [x] `docs/design/Issue-1-3-glass-weight-function/requirements.md` を確認する
- [x] `calculate_glass_weight()` を実装する
- [x] サグ計算の符号処理（R1正、R2負）を反映する
- [x] 日本語GoogleスタイルDocstringを整備する

## テストタスク
- [x] 球面レンズの重量計算テストを追加する
- [x] 非球面（R1のみ）の重量計算テストを追加する
- [x] 非球面（両面）の重量計算テストを追加する
- [x] 平面を含むレンズの重量計算テストを追加する
- [x] サグ計算失敗時の`None`返却テストを追加する
- [x] エッジケース（比重極小/極大、厚み0近傍）のテストを追加する

## 検証タスク
- [ ] VBA版との数値一致（相対誤差 1e-6 以下）を確認する
- [ ] `uv run ruff check src tests` を実行する
- [ ] `uv run ruff format src tests` を実行する
- [ ] `uv run pytest tests/test_optics_calculations.py -v` を実行する
