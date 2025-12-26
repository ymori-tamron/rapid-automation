# Taskリスト（Issue 1-2: FocalLength関数）

## 実装タスク
- [x] `references/VBA/basUDF.bas` の `FocalLength` 実装詳細を確認する
- [x] `src/optics/calculations.py` に `PLANE_RADIUS` 定数を定義する
- [x] `calculate_focal_length()` を実装する
- [x] 入力正規化（平面の半径、屈折率の下限）を実装する
- [x] レンズメーカー公式によるパワー計算を実装する
- [x] `PW == 0` の無限大判定を実装する
- [x] **VBA互換性のため、例外を送出せず入力値を正規化する処理を実装する**
- [x] 日本語GoogleスタイルDocstringを整備する

## テストタスク
- [x] 両凸レンズの焦点距離計算テストを追加する
- [x] 両凹レンズの焦点距離計算テストを追加する
- [x] 平凸レンズの焦点距離計算テストを追加する
- [x] 焦点距離が無限大になるケースのテストを追加する
- [x] エッジケース（`refractive_index=1`, `thickness=0`）のテストを追加する

## 検証タスク
- [x] CODEVとの数値一致（相対誤差 1e-7 以下）を確認する
- [x] `uv run ruff check src tests` を実行する
- [x] `uv run ruff format src tests` を実行する
- [x] `uv run pytest tests/test_optics_calculations.py::test_focal_length* -v` を実行する
