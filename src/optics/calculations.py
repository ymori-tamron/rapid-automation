from __future__ import annotations

import math
from dataclasses import dataclass

PLANE_RADIUS = 1e10


@dataclass
class AsphericCoefficients:
    """非球面係数を保持するデータクラス。

    Attributes:
        conic (float): コーニック定数。
        a4 (float): 4次の非球面係数。
        a6 (float): 6次の非球面係数。
        a8 (float): 8次の非球面係数。
        a10 (float): 10次の非球面係数。
        a12 (float): 12次の非球面係数。
        a14 (float): 14次の非球面係数。
    """

    conic: float = 0.0
    a4: float = 0.0
    a6: float = 0.0
    a8: float = 0.0
    a10: float = 0.0
    a12: float = 0.0
    a14: float = 0.0


def _validate_number(value: object, name: str) -> None:
    if not isinstance(value, (int, float)) or isinstance(value, bool):
        raise TypeError(f"{name}は数値である必要があります。")


def _is_number(value: object) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def calculate_sag(
    radius: float | None,
    diameter: float,
    coefficients: AsphericCoefficients | None = None,
) -> float | None:
    """レンズ表面のサグ量（深さ）を計算する。

    球面および非球面形状のサグ量を計算します。非球面の場合は
    コーニック定数と非球面係数（A4～A14）を考慮します。

    Args:
        radius (float | None): 曲率半径[mm]。Noneまたは0の場合は平面として扱う。
        diameter (float): サグ量を計算する直径[mm]。
        coefficients (AsphericCoefficients | None): 非球面係数。省略時は球面として扱う。

    Returns:
        float | None: サグ量[mm]。計算不可能な場合はNoneを返す。
    """

    if radius is None or radius == 0:
        c = 0.0
    else:
        _validate_number(radius, "radius")
        c = 1.0 / float(radius)

    _validate_number(diameter, "diameter")
    h = float(diameter) / 2.0

    if coefficients is None:
        coefficients = AsphericCoefficients()

    arg = 1.0 - (1.0 + coefficients.conic) * (c**2) * (h**2)
    if arg < 0:
        return None

    base = (c * (h**2)) / (1.0 + math.sqrt(arg))
    aspheric = (
        coefficients.a4 * (h**4)
        + coefficients.a6 * (h**6)
        + coefficients.a8 * (h**8)
        + coefficients.a10 * (h**10)
        + coefficients.a12 * (h**12)
        + coefficients.a14 * (h**14)
    )
    return base + aspheric


def calculate_focal_length(
    radius1: float | None,
    radius2: float | None,
    thickness: float,
    refractive_index: float,
) -> float | str:
    """単レンズの焦点距離を計算する（薄肉レンズ近似）。

    レンズメーカー公式を使用して焦点距離を算出します。
    VBA互換性のため、入力が不正な場合は例外を送出せずに正規化します。

    Args:
        radius1: R1面の曲率半径[mm]。Noneまたは0、数値以外の場合は平面扱い。
        radius2: R2面の曲率半径[mm]。Noneまたは0、数値以外の場合は平面扱い。
        thickness: 中心厚[mm]。数値以外の場合は0扱い。
        refractive_index: 硝材の屈折率。数値以外や1未満の場合は1扱い。

    Returns:
        焦点距離[mm]。無限大の場合は文字列"Inf"を返す。
    """

    r1_is_plane = (not _is_number(radius1)) or radius1 == 0
    r2_is_plane = (not _is_number(radius2)) or radius2 == 0

    r1 = PLANE_RADIUS if r1_is_plane else float(radius1)
    r2 = PLANE_RADIUS if r2_is_plane else float(radius2)

    if not _is_number(refractive_index):
        n = 1.0
    else:
        n = float(refractive_index)
        if n < 1:
            n = 1.0

    t = 0.0 if not _is_number(thickness) else float(thickness)

    if r1_is_plane and r2_is_plane:
        return "Inf"

    pw = (n - 1.0) * (1.0 / r1 - 1.0 / r2 + (t * (n - 1.0)) / (n * r1 * r2))
    if pw == 0:
        return "Inf"
    return 1.0 / pw


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
    """単レンズの重量を数値積分で算出する。

    球面・非球面の両方に対応し、VBA版と同様に100ステップの台形積分で体積を求める。

    Args:
        radius1: R1面の曲率半径[mm]。Noneまたは0の場合は平面として扱う。
        radius2: R2面の曲率半径[mm]。Noneまたは0の場合は平面として扱う。
        thickness: 中心厚[mm]。
        specific_gravity: 硝材の比重[g/cm^3]。
        diameter1: R1面の有効径[mm]。
        diameter2: R2面の有効径[mm]。
        max_diameter: 最大外径[mm]。
        coefficients1: R1面の非球面係数。
        coefficients2: R2面の非球面係数。

    Returns:
        重量[g]。サグ計算に失敗した場合はNoneを返す。
    """

    _validate_number(thickness, "thickness")
    _validate_number(specific_gravity, "specific_gravity")
    _validate_number(diameter1, "diameter1")
    _validate_number(diameter2, "diameter2")
    _validate_number(max_diameter, "max_diameter")

    if coefficients1 is None:
        coefficients1 = AsphericCoefficients()
    if coefficients2 is None:
        coefficients2 = AsphericCoefficients()

    i_max = 100
    pi = math.pi

    dh1 = (float(diameter1) / 2.0) / i_max
    dh2 = (float(diameter2) / 2.0) / i_max

    def calculate_volume(
        radius: float | None,
        coefficients: AsphericCoefficients,
        dh: float,
        sign: float,
    ) -> float | None:
        z_values = [0.0] * (i_max + 1)
        volume = 0.0
        for i in range(i_max + 1):
            diameter = (dh * i) * 2.0
            sag = calculate_sag(radius, diameter, coefficients)
            if sag is None:
                return None
            z_values[i] = sign * sag
            if i > 0:
                volume -= (
                    ((dh * i_max) ** 2) * 2.0
                    - (dh * (i - 1)) ** 2
                    - (dh * i) ** 2
                ) * pi * (z_values[i] - z_values[i - 1]) / 2.0

        volume -= (((float(max_diameter) / 2.0) ** 2) - ((dh * i_max) ** 2)) * pi * z_values[
            i_max
        ]
        return volume

    volume1 = calculate_volume(radius1, coefficients1, dh1, 1.0)
    if volume1 is None:
        return None

    volume2 = calculate_volume(radius2, coefficients2, dh2, -1.0)
    if volume2 is None:
        return None

    return (
        float(specific_gravity)
        * (pi * ((float(max_diameter) / 2.0) ** 2) * float(thickness) + volume1 + volume2)
        / 1000.0
    )
