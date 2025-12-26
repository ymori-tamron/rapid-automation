# GlassWeight関数 要求仕様

## 1. 目的
単レンズの重量を数値積分で算出する`GlassWeight`関数の要求仕様を定義し、VBA実装と数値的に一致するPython実装の基準とします。サグ計算はIssue 1-1の`calculate_sag()`に準拠します。

## 2. 参照
- `references/VBA/basUDF.bas` の `GlassWeight` 関数
- `docs/scratch/migration_plan.md` の Issue 1-3
- `docs/scratch/issues/issue-1-3.md`
- `docs/design/Issue1-1-sag-function/requirements.md`
- 実装ファイル: `src/optics/calculations.py`
- テストファイル: `tests/test_optics_calculations.py`

## 3. 用語・記号
- `R1`, `R2`: レンズ第1面/第2面の曲率半径[mm]
- `t`: レンズ中心厚[mm]
- `SpG`: 硝材の比重[g/cm^3]
- `D1`, `D2`: R1/R2面の有効径[mm]
- `Dmax`: 最大外径[mm]
- `iMax`: 数値積分の分割数（固定値100）
- `dh1`, `dh2`: 半径方向の刻み幅[mm]
- `z(i, j)`: 半径位置`i`のサグ量配列（R1=1, R2=2）
- `v1`, `v2`: R1/R2面の体積補正項[mm^3]

## 4. 入力仕様
### 4.1 関数シグネチャ（想定）
```python
def calculate_glass_weight(
    radius1: float | None,
    radius2: float | None,
    thickness: float,
    specific_gravity: float,
    diameter1: float,
    diameter2: float,
    max_diameter: float,
    coefficients1: AsphericCoefficients | None = None,
    coefficients2: AsphericCoefficients | None = None,
) -> float | None:
    ...
```

### 4.2 引数の意味と制約
- `radius1`, `radius2`:
  - 曲率半径[mm]。`None`または`0`の場合は平面として扱う。
  - `calculate_sag()`の入力仕様に従う。
- `thickness`:
  - 中心厚[mm]。正の実数を想定する（0や負値の扱いは実装で検討）。
  - VBAでは型チェックなしで処理されるが、Pythonでは妥当性検証を考慮。
- `specific_gravity`:
  - 比重[g/cm^3]。正の実数を想定する（0や負値の扱いは実装で検討）。
  - VBAでは型チェックなしで処理されるが、Pythonでは妥当性検証を考慮。
- `diameter1`, `diameter2`:
  - R1/R2面の有効径[mm]。正の実数を想定する。
- `max_diameter`:
  - 最大外径[mm]。正の実数を想定する。
- `coefficients1`, `coefficients2`:
  - 非球面係数（`AsphericCoefficients`）。省略時は全て0として扱う。

## 5. 出力仕様
- 戻り値は重量[g]（`float`）。
- サグ計算が失敗した場合は`None`を返す。

## 6. 計算仕様
### 6.1 参照するサグ計算
- サグ量の算出は`calculate_sag()`に準拠し、直径`d`を入力として使用する。
- R1は正符号、R2は負符号として扱う（VBAの`(j * (-2) + 3)`に対応）。

### 6.2 数値積分の前提
- 分割数は固定で`iMax = 100`。
- 半径方向刻み幅は次で定義する。
  - `dh1 = (D1 / 2) / iMax`
  - `dh2 = (D2 / 2) / iMax`

### 6.3 サグ配列の算出
- `j = 1`（R1）、`j = 2`（R2）について、`i = 0..iMax`で以下を計算する。
  - `d = (dhj * i) * 2`（直径）
  - `sag = calculate_sag(rj, d, coefficientsj)`
  - `z(i, j) = sag`（R1）
  - `z(i, j) = -sag`（R2）
- R1は正符号、R2は負符号として扱う（VBAの`(j * (-2) + 3)`に対応）。
  - j=1のとき: 1×(-2)+3 = 1 → 正の符号（R1面は正方向のサグ）
  - j=2のとき: 2×(-2)+3 = -1 → 負の符号（R2面は負方向のサグ）
- `sag`が`None`の場合は`None`を返す。

### 6.4 体積補正項の台形積分
- `v(j)`は以下の差分台形積分により算出する（VBA式に一致）。

```
for i in 1..iMax:
  v(j) -= (
    ((dhj * iMax)^2) * 2
    - (dhj * (i - 1))^2
    - (dhj * i)^2
  ) * pi * (z(i, j) - z(i - 1, j)) / 2
```

- さらに外周部の補正を加える。

```
v(j) -= (((Dmax / 2)^2) - ((dhj * iMax)^2)) * pi * z(iMax, j)
```

### 6.5 最終重量
- 円柱体積と補正体積の合計から重量を算出する。

```
weight = SpG * (
  pi * ((Dmax / 2)^2) * thickness + v1 + v2
) / 1000
```

- mm^3からcm^3への換算のため`/ 1000`を適用する。

### 6.6 単位系の注意点
- 入力: すべて[mm]単位（曲率半径、厚み、直径）
- 比重: [g/cm³]
- 体積計算: mm³で実施
- 最終変換: mm³ → cm³ (÷1000) → 重量[g] (×比重)

## 7. 異常系・入力検証
- `calculate_sag()`が`None`を返した場合は`None`を返す。
- 数値でない入力は`TypeError`の対象とする（VBA互換ではなくPythonとしての安全性を優先）。

## 8. 数値精度・互換性
- VBA `basUDF.GlassWeight` と相対誤差`1e-6`以下で一致すること。
- 倍精度浮動小数点（Pythonの`float`）で計算する。

## 9. テスト要件
最低限、以下の観点を満たすこと。
- 球面レンズの重量計算（両面球面、標準寸法）
- 非球面レンズ（R1面のみ非球面）の重量計算
- 非球面レンズ（両面非球面）の重量計算
- 平面を含むレンズ（`radius=None`）の重量計算
- サグ計算失敗時の`None`返却
- エッジケース（以下のような条件での動作確認）:
  - 比重が極小: SpG = 0.01
  - 比重が極大: SpG = 20.0 (重金属ガラス想定)
  - 厚みが極小: t = 0.1mm
  - 平面レンズ: R1=None, R2=None, t=5.0mm

## 10. 完了条件（Definition of Done）
- `calculate_glass_weight`の実装が仕様通りであること。
- VBA版との数値一致が検証できること（相対誤差`1e-6`以下）。
- 単体テストで上記の観点を満たすこと。
- Docstring（日本語・Googleスタイル）と型ヒントが完備されていること。

## 11. VBA互換性に関する注記
- VBAの`#VALUE!`エラーはPythonでは`None`返却で代替
- VBAの`Variant`型はPythonでは`float | None`で表現
- 数値でない入力の扱い: VBAは暗黙的変換、Pythonは明示的な型検証を推奨
