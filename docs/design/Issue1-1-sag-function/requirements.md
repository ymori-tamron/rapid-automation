# Sag関数 要求仕様

## 1. 目的
光学レンズの表面形状からサグ量（深さ）を算出する`Sag`関数の要求仕様を定義し、VBA実装と数値的に一致するPython実装の基準とします。

## 2. 参照
- `references/VBA/basUDF.bas` の `Sag` 関数
- `docs/scratch/migration_plan.md` の Issue 1-1
- `docs/scratch/issues/issue-1-1.md`

## 3. 用語・記号
- `R`: 曲率半径[mm]
- `c`: 曲率（`c = 1 / R`）
- `h`: サグ量計算に用いる半径（`h = diameter / 2`）
- `K`: コーニック定数（`conic`）
- `A4, A6, A8, A10, A12, A14`: 非球面係数

## 4. 入力仕様
### 4.1 関数シグネチャ（想定）
```python
def calculate_sag(
    radius: float | None,
    diameter: float,
    coefficients: AsphericCoefficients | None = None,
) -> float | None:
    ...
```

### 4.2 引数の意味と制約
- `radius`:
  - 曲率半径[mm]。`None`または`0`の場合は平面として扱う。
  - 有限の実数を想定する。
- `diameter`:
  - サグ量を計算する直径[mm]。
  - 有限の実数を想定する。
- `coefficients`:
  - 非球面係数のデータクラス。省略時はすべて0として扱う。

### 4.3 非球面係数データクラス（想定）
```python
from dataclasses import dataclass

@dataclass
class AsphericCoefficients:
    """非球面係数を保持するデータクラス。"""
    conic: float = 0.0
    a4: float = 0.0
    a6: float = 0.0
    a8: float = 0.0
    a10: float = 0.0
    a12: float = 0.0
    a14: float = 0.0
```

## 5. 出力仕様
- 戻り値はサグ量[mm]（`float`）。
- 計算式の平方根の中が負になる場合は`None`を返す。

## 6. 計算仕様
### 6.1 基本式
非球面のサグ量は以下で定義する。

```
z = (c·h^2) / (1 + sqrt(1 - (1 + K)·c^2·h^2))
    + A4·h^4 + A6·h^6 + A8·h^8 + A10·h^10 + A12·h^12 + A14·h^14
```

### 6.2 計算手順
1. `radius`が`None`または`0`の場合は`c = 0`とする。それ以外は`c = 1 / radius`。
2. `h = diameter / 2` を計算する。
3. `arg = 1 - (1 + conic) * c^2 * h^2` を計算する。
4. `arg < 0` の場合は`None`を返す。
5. 球面項と非球面補正項を合算してサグ量を返す。

## 7. 異常系・入力検証
- `radius`が`None`以外で数値でない場合は`TypeError`を送出する。
- `diameter`が数値でない場合は`TypeError`を送出する。
- `coefficients`が指定された場合、各係数は数値であることを前提とする。

## 8. 数値精度・互換性
- VBA `basUDF.Sag` と相対誤差`1e-7`以下で一致すること。
- 倍精度浮動小数点（Pythonの`float`）で計算する。

## 9. テスト要件
最低限、以下の観点を満たすこと。
- 球面凸レンズのサグ量計算
- 球面凹レンズのサグ量計算
- 非球面（コーニック定数のみ）のサグ量計算
- 非球面（A4係数のみ）のサグ量計算
- 平面（`radius=None`）のサグ量計算
- 平方根の中が負となるケースの`None`返却
- 直径0のエッジケース

## 10. 完了条件（Definition of Done）
- `calculate_sag`の実装が仕様通りであること。
- VBA版との数値一致が検証できること（相対誤差`1e-7`以下）。
- 単体テストで上記の観点を満たすこと。
