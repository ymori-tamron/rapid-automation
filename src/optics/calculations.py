from __future__ import annotations

import math
from dataclasses import dataclass


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
