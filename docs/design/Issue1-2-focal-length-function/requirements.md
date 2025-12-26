# FocalLength関数 要求仕様

## 1. 目的
単レンズの焦点距離を薄肉レンズ近似（レンズメーカー公式）で算出する`FocalLength`関数の要求仕様を定義し、VBA実装と数値的に一致するPython実装の基準とします。

## 2. 参照
- `references/VBA/basUDF.bas` の `FocalLength` 関数
- `docs/scratch/migration_plan.md` の Issue 1-2
- `docs/scratch/issues/issue-1-2.md`
- 実装ファイル: `src/optics/calculations.py`
- テストファイル: `tests/test_optics_calculations.py`

## 3. 用語・記号
- `R1`: レンズ第1面の曲率半径[mm]
- `R2`: レンズ第2面の曲率半径[mm]
- `t`: レンズ中心厚[mm]
- `n`: 硝材の屈折率（無次元）
- `PW`: レンズパワー（光学的な屈折力）
- `f`: 焦点距離[mm]

## 4. 入力仕様
### 4.1 関数シグネチャ（想定）
```python
def calculate_focal_length(
    radius1: float | None,
    radius2: float | None,
    thickness: float,
    refractive_index: float,
) -> float | str:
    ...
```

### 4.2 引数の意味と制約
- `radius1`:
  - R1面の曲率半径[mm]。`None`または`0`の場合は平面として扱う。
  - 有限の実数を想定する。
- `radius2`:
  - R2面の曲率半径[mm]。`None`または`0`の場合は平面として扱う。
  - 有限の実数を想定する。
- `thickness`:
  - レンズ中心厚[mm]。
  - 有限の実数を想定する。
- `refractive_index`:
  - 硝材の屈折率（無次元）。
  - `1`未満の場合は`n = 1`として扱う。

### 4.3 定数定義
```python
PLANE_RADIUS = 1e10  # 平面を表す非常に大きい曲率半径値[mm]
```

### 4.4 入力正規化
- `radius1` / `radius2` が数値でない、`None`、または `0` の場合は、平面として `PLANE_RADIUS = 1e10` に置換する。
- `refractive_index` が数値でない、または `< 1` の場合は `n = 1` として扱う。
- `thickness` が数値でない場合は `0` として扱う。

## 5. 出力仕様
- 戻り値は `float | str`。
- 有限焦点距離は `float`（単位: mm）で返す。
- 無限大の場合は文字列 `"Inf"` を返す。

## 6. 計算仕様
### 6.1 レンズメーカー公式（薄肉レンズ近似）
```
PW = (n - 1) × (1/R1 - 1/R2 + t × (n - 1) / (n × R1 × R2))
f = 1 / PW
```

### 6.2 計算手順
1. `radius1`が数値でない、`None`、または`0`の場合は`r1 = PLANE_RADIUS`とする。それ以外は`r1 = radius1`。
2. `radius2`が数値でない、`None`、または`0`の場合は`r2 = PLANE_RADIUS`とする。それ以外は`r2 = radius2`。
3. `refractive_index`が数値でない、または`< 1`の場合は`n = 1`とする。それ以外は`n = refractive_index`。
4. `thickness`が数値でない場合は`t = 0`とする。それ以外は`t = thickness`。
5. `PW = (n - 1) × (1/r1 - 1/r2 + t × (n - 1) / (n × r1 × r2))` を計算する。
6. `PW == 0` の場合は`"Inf"`を返す。それ以外は`1 / PW`を返す。

## 7. 異常系・入力検証
**VBA互換性のため、例外は送出しない。**

VBA実装では、不正な入力値に対して例外を送出せず、入力値を正規化してデフォルト値を使用します。Python実装もこの動作に合わせます：
- 入力値が数値でない場合は、上記4.4の正規化ルールに従う。
- エラーは発生させず、静かに正規化して計算を続行する。

## 8. 数値精度・互換性
- VBA `basUDF.FocalLength` と相対誤差`1e-7`以下で一致すること。
- 倍精度浮動小数点（Pythonの`float`）で計算する。

## 9. テスト要件
最低限、以下の観点を満たすこと。
- 両凸レンズの焦点距離計算: `radius1=100.0, radius2=-100.0, thickness=5.0, refractive_index=1.5168`
- 両凹レンズの焦点距離計算: `radius1=-100.0, radius2=100.0, thickness=5.0, refractive_index=1.5168`
- 平凸レンズの焦点距離計算: `radius1=50.0, radius2=None, thickness=3.0, refractive_index=1.5`
- 焦点距離が無限大の場合: `radius1=None, radius2=None, thickness=3.0, refractive_index=1.5`
- エッジケース: `refractive_index=1`, `thickness=0`

## 10. 完了条件（Definition of Done）
- `calculate_focal_length`の実装が仕様通りであること。
- VBA版との数値一致が検証できること（相対誤差`1e-7`以下）。
- 単体テストで上記の観点を満たすこと。
- Docstring（日本語・Googleスタイル）と型ヒントが完備されていること。
