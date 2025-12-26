# Taskリスト（Issue 1-1: Sag関数）

## 実装タスク
- [ ] `references/VBA/basUDF.bas` の `Sag` 実装詳細を確認する
- [ ] `AsphericCoefficients` データクラスを追加する
- [ ] `calculate_sag()` を実装する
- [ ] 日本語GoogleスタイルDocstringを整備する

## テストタスク
- [ ] 球面凸のサグ量計算テストを追加する
- [ ] 球面凹のサグ量計算テストを追加する
- [ ] 非球面（コーニック定数）のテストを追加する
- [ ] 非球面（A4係数）のテストを追加する
- [ ] 平面（`radius=None`）のテストを追加する
- [ ] 平方根の中が負になるケースのテストを追加する
- [ ] 直径0のエッジケーステストを追加する

## 検証タスク
- [ ] VBA版との数値一致（相対誤差 1e-7 以下）を確認する
- [ ] `uv run ruff check src tests` を実行する
- [ ] `uv run ruff format src tests` を実行する
- [ ] `uv run pytest tests/test_optics_calculations.py -v` を実行する
