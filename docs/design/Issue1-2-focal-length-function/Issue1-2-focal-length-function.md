# Issue 1-2 要求仕様: FocalLength関数（焦点距離計算）

## 1. 目的
単レンズの焦点距離を薄肉レンズ近似（レンズメーカー公式）で算出する関数をPythonに移植し、VBA版と数値的に一致させる。

## 2. 背景・位置づけ
- 移植計画: `docs/scratch/migration_plan.md` の Phase 1 / Issue 1-2
- 対象Issue: `docs/scratch/issues/issue-1-2.md`
- ※ サグ計算は Issue 1-1 にて扱う（本書の対象外）

## 3. 対象範囲
### 3.1 対象
- `calculate_focal_length()` の仕様
- 入力値の正規化（平面の扱い、屈折率の下限）
- 無限大焦点距離の判定

### 3.2 対象外
- サグ計算（Issue 1-1）
- 多レンズ系の合成焦点距離計算
- 材料データベースやCODE Vパーサーとの連携

## 4. 入力仕様
| 引数名 | 型 | 単位 | 説明 |
|---|---|---|---|
| `radius1` | `float | None` | mm | R1面の曲率半径。`None`/`0` は平面扱い |
| `radius2` | `float | None` | mm | R2面の曲率半径。`None`/`0` は平面扱い |
| `thickness` | `float` | mm | 中心厚 |
| `refractive_index` | `float` | - | 硝材の屈折率 |

### 4.1 入力正規化
- `radius1` / `radius2` が `None` または `0` の場合は、平面として扱う。
- 平面は非常に大きい値 `PLANE_RADIUS = 10000000000`（1e10）に置換する。
- `refractive_index < 1` の場合は `n = 1` として扱う。

## 5. 計算仕様
### 5.1 レンズメーカー公式（薄肉レンズ近似）
```
PW = (n - 1) * (1/R1 - 1/R2 + t * (n - 1) / (n * R1 * R2))
f = 1 / PW
```

### 5.2 無限大の判定
- `PW == 0` の場合は焦点距離を `"Inf"` で返す。

## 6. 出力仕様
- 戻り値は `float | str`。
- 有限焦点距離は `float`（単位: mm）で返す。
- 無限大の場合は文字列 `"Inf"` を返す。

## 7. 例外・エラー処理
- 入力型が仕様外の場合は `TypeError` を送出する。
- `thickness` が数値でない場合は `TypeError` を送出する。
- 数値計算で `ZeroDivisionError` が発生し得る場合は、事前に `PW == 0` を判定する。

## 8. 精度要件
- VBA `basUDF.FocalLength` と数値一致（相対誤差 1e-7 以下を目安）。

## 9. テスト要件（抜粋）
- 両凸: `radius1=100.0, radius2=-100.0, thickness=5.0, refractive_index=1.5168`
- 両凹: `radius1=-100.0, radius2=100.0, thickness=5.0, refractive_index=1.5168`
- 平凸: `radius1=50.0, radius2=None, thickness=3.0, refractive_index=1.5`
- 無限大: `radius1=None, radius2=None, thickness=3.0, refractive_index=1.5`
- エッジケース: `refractive_index=1`, `thickness=0`

## 10. 完了条件（Definition of Done）
- `calculate_focal_length()` が仕様どおりに実装されている
- Docstring（日本語・Googleスタイル）と型ヒントが完備
- `tests/test_optics_calculations.py` で最低5ケースがパス
- VBA実装と数値的に一致

## 11. 参照
- `docs/scratch/issues/issue-1-2.md`
- `docs/scratch/migration_plan.md`
- `references/VBA/basUDF.bas`（FocalLength）
