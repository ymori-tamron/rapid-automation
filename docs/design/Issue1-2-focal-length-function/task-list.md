## タスクリスト
- [ ] `references/VBA/basUDF.bas` のFocalLength実装を確認
- [ ] `src/optics/calculations.py` に関数スケルトンを追加
- [ ] 入力正規化（平面の半径、屈折率の下限）を実装
- [ ] レンズメーカー公式によるパワー計算を実装
- [ ] `PW == 0` の無限大判定を実装
- [ ] Docstring（日本語）と型ヒントを整備
- [ ] `tests/test_optics_calculations.py` に5ケース以上のテストを追加
- [ ] `uv run pytest tests/test_optics_calculations.py::test_focal_length* -v` を実行
- [ ] VBA出力と数値一致を確認（相対誤差1e-7目安）
