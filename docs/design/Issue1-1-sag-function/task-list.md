# Taskリスト（Issue 1-1: Sag関数）

## 実装タスク
- [x] `references/VBA/basUDF.bas` の `Sag` 実装詳細を確認する
- [x] `AsphericCoefficients` データクラスを追加する
- [x] `calculate_sag()` を実装する
- [x] 日本語GoogleスタイルDocstringを整備する

## テストタスク
- [x] 球面凸のサグ量計算テストを追加する
- [x] 球面凹のサグ量計算テストを追加する
- [x] 非球面（コーニック定数）のテストを追加する
- [x] 非球面（A4係数）のテストを追加する
- [x] 平面（`radius=None`）のテストを追加する
- [x] 平方根の中が負になるケースのテストを追加する
- [x] 直径0のエッジケーステストを追加する

## 検証タスク
- [x] CODE V版との数値一致（相対誤差 1e-7 以下）を確認する
- [x] `uv run ruff check src tests` を実行する
- [x] `uv run ruff format src tests` を実行する
- [x] `uv run pytest tests/test_optics_calculations.py -v` を実行する
