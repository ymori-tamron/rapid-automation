## 実装Plan
1. 要求仕様の確認とVBA実装の差分確認（`references/VBA/basUDF.bas`）
2. `src/optics/calculations.py` に `calculate_focal_length()` を追加
3. 入力正規化（平面/屈折率下限）とパワー計算の実装
4. 無限大判定（`PW == 0`）と戻り値仕様の実装
5. 日本語GoogleスタイルDocstringと型ヒントの整備
6. `tests/test_optics_calculations.py` にテストケース追加
7. VBA結果との数値一致を確認し、必要に応じて精度調整
