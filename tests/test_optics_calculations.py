import math

import pytest

from src.optics.calculations import (
    AsphericCoefficients,
    calculate_focal_length,
    calculate_glass_weight,
    calculate_sag,
)

# =============================================================================
# テストデータ: CODE Vから取得した期待値
# =============================================================================
# フォーマット: (radius, diameter, coefficients, expected_sag, test_description)
#
# 期待値の取得方法:
# 1. CODE Vで各テストケースのパラメータを設定
# 2. SAGFコマンドで指定半径（diameter/2）のサグ量を計算
# 3. 15桁精度で期待値を記録
#
# =============================================================================

SAG_TEST_CASES = [
    # 球面レンズ（凸）
    # R=50mm, φ=20mm の球面凸レンズのサグ量
    (50.0, 20.0, None, 1.010205144336438, "球面凸レンズ (R=50mm, φ=20mm)"),
    # 球面レンズ（凹）
    # R=-50mm, φ=20mm の球面凹レンズのサグ量（負の値）
    (-50.0, 20.0, None, -1.010205144336438, "球面凹レンズ (R=-50mm, φ=20mm)"),
    # 非球面レンズ（コーニック定数のみ）
    # R=50mm, φ=20mm, K=-0.5（双曲面）の非球面レンズ
    (
        50.0,
        20.0,
        AsphericCoefficients(conic=-0.5),
        1.005050633883346,
        "非球面レンズ（コーニック定数=-0.5）",
    ),
    # 非球面レンズ（A4係数のみ）
    # R=50mm, φ=20mm, A4=1e-6 の非球面レンズ
    (
        50.0,
        20.0,
        AsphericCoefficients(a4=1e-6),
        1.020205144336438,
        "非球面レンズ（A4係数=1e-6）",
    ),
    # 非球面レンズ（複数係数の組み合わせ）
    # R=50mm, φ=20mm, K=-0.5, A4=1e-6, A6=-2e-9 の複合非球面
    (
        50.0,
        20.0,
        AsphericCoefficients(conic=-0.5, a4=1e-6, a6=-2e-9),
        1.013050633883346,
        "非球面レンズ（複合: K=-0.5, A4=1e-6, A6=-2e-9）",
    ),
    # 非球面レンズ（A6係数のみ）
    # R=50mm, φ=20mm, A6=1e-9 の非球面レンズ
    (
        50.0,
        20.0,
        AsphericCoefficients(a6=1e-9),
        1.011205144336438,
        "非球面レンズ（A6係数=1e-9）",
    ),
    # 非球面レンズ（A8係数のみ）
    # R=50mm, φ=20mm, A8=1e-12 の非球面レンズ
    (
        50.0,
        20.0,
        AsphericCoefficients(a8=1e-12),
        1.010305144336438,
        "非球面レンズ（A8係数=1e-12）",
    ),
    # 非球面レンズ（A10係数のみ）
    # R=50mm, φ=20mm, A10=1e-15 の非球面レンズ
    (
        50.0,
        20.0,
        AsphericCoefficients(a10=1e-15),
        1.010215144336438,
        "非球面レンズ（A10係数=1e-15）",
    ),
    # 非球面レンズ（A12係数のみ）
    # R=50mm, φ=20mm, A12=1e-18 の非球面レンズ
    (
        50.0,
        20.0,
        AsphericCoefficients(a12=1e-18),
        1.010206144336438,
        "非球面レンズ（A12係数=1e-18）",
    ),
    # 非球面レンズ（A14係数のみ）
    # R=50mm, φ=20mm, A14=1e-21 の非球面レンズ
    (
        50.0,
        20.0,
        AsphericCoefficients(a14=1e-21),
        1.010205244336438,
        "非球面レンズ（A14係数=1e-21）",
    ),
    # 非球面レンズ（全係数使用）
    # R=50mm, φ=20mm, 全ての非球面係数を含む複雑な非球面形状
    # 実際の光学設計で使用される典型的な係数値を設定
    (
        50.0,
        20.0,
        AsphericCoefficients(
            conic=-0.8,
            a4=2.5e-6,
            a6=-1.2e-9,
            a8=5.3e-13,
            a10=-8.7e-17,
            a12=3.2e-20,
            a14=-6.1e-24,
        ),
        1.025860201615352,
        "非球面レンズ（全係数使用: K=-0.8, A4～A14）",
    ),
    # 平面（曲率半径=無限大）
    # radius=None は平面を表し、サグ量は常に0
    (None, 20.0, None, 0.0, "平面（曲率半径=無限大）"),
    # 平面（曲率半径=0）
    # radius=0 も平面として扱い、サグ量は常に0
    (0.0, 20.0, None, 0.0, "平面（曲率半径=0）"),
    # エッジケース：直径0
    # 直径が0の場合、サグ量は0になる
    (50.0, 0.0, None, 0.0, "エッジケース: 直径0"),
    # 小径レンズ
    # R=2mm, φ=1.0mm の小径球面レンズ
    (2.0, 1.0, None, 0.063508326896292, "小径レンズ (R=2mm, φ=1mm)"),
    # 大口径レンズ
    # R=200mm, φ=100mm の大口径球面レンズ
    (200.0, 100.0, None, 6.350832689629156, "大口径レンズ (R=200mm, φ=100mm)"),
]

FOCAL_LENGTH_TEST_CASES = [
    (
        100.0,
        -100.0,
        5.0,
        1.5168,
        97.58040934528817,
        "両凸レンズ (R1=100, R2=-100, t=5.0, n=1.5168)",
    ),
    (
        -100.0,
        100.0,
        5.0,
        1.5168,
        -95.93208299962866,
        "両凹レンズ (R1=-100, R2=100, t=5.0, n=1.5168)",
    ),
    (
        50.0,
        None,
        3.0,
        1.5,
        100.00000049,
        "平凸レンズ (R1=50, R2=平面, t=3.0, n=1.5)",
    ),
    (
        None,
        None,
        3.0,
        1.5,
        "Inf",
        "両面平面 (R1=平面, R2=平面, t=3.0, n=1.5)",
    ),
    (
        50.0,
        -50.0,
        0.0,
        1.5,
        50.0,
        "エッジケース: 厚み0 (R1=50, R2=-50, t=0, n=1.5)",
    ),
    (
        50.0,
        -50.0,
        5.0,
        1.0,
        "Inf",
        "エッジケース: 屈折率1 (R1=50, R2=-50, t=5.0, n=1.0)",
    ),
]


@pytest.mark.parametrize(
    "radius,diameter,coefficients,expected,description",
    SAG_TEST_CASES,
)
def test_calculate_sag_parametrized(
    radius: float | None,
    diameter: float,
    coefficients: AsphericCoefficients | None,
    expected: float,
    description: str,
) -> None:
    """CODE VまたはVBA版から取得した期待値と比較してサグ量の正確性を検証する。

    Args:
        radius: 曲率半径[mm]
        diameter: レンズ直径[mm]
        coefficients: 非球面係数
        expected: 期待されるサグ量[mm]（CODE VまたはVBAから取得）
        description: テストケースの説明
    """
    result = calculate_sag(radius, diameter, coefficients)
    assert result == pytest.approx(expected, rel=1e-7), f"Failed: {description}"


@pytest.mark.parametrize(
    "radius1,radius2,thickness,refractive_index,expected,description",
    FOCAL_LENGTH_TEST_CASES,
)
def test_calculate_focal_length_parametrized(
    radius1: float | None,
    radius2: float | None,
    thickness: float,
    refractive_index: float,
    expected: float | str,
    description: str,
) -> None:
    """焦点距離の計算結果が期待値と一致することを検証する。"""
    result = calculate_focal_length(radius1, radius2, thickness, refractive_index)
    if expected == "Inf":
        assert result == "Inf", f"Failed: {description}"
    else:
        assert result == pytest.approx(expected, rel=1e-7), f"Failed: {description}"


# =============================================================================
# 異常系テスト
# =============================================================================


def test_calculate_sag_returns_none_when_sqrt_argument_negative() -> None:
    """平方根の引数が負になる場合はNoneを返すことを検証する。

    極端に大きい直径と小さい曲率半径、および負のコーニック定数の組み合わせで
    式 arg = 1 - (1+K)*c²*h² が負になるケース。
    """
    result = calculate_sag(10.0, 100.0)
    assert result is None


def test_calculate_sag_returns_none_when_sqrt_argument_negative_with_conic() -> None:
    """極端なコーニック定数で平方根の引数が負になる場合を検証する。

    R=5mm, φ=20mm, K=-20.0 のような極端な組み合わせ。
    arg = 1 - (1+K)*c²*h² = 1 - (1-20)*(1/5)²*(10)²
        = 1 - (-19)*0.04*100 = 1 - (-76) = 77 > 0 (まだ正)

    より確実に負にするため、K=-100, R=5mm, φ=25mm を使用。
    arg = 1 - (1-100)*(1/5)²*(12.5)² = 1 - (-99)*0.04*156.25
        = 1 - (-618.75) = 619.75 > 0 (まだ正)

    注: 負のコーニック定数では arg は増加するため、正のコーニック定数を使用。
    R=5mm, φ=25mm, K=10.0
    arg = 1 - (1+10)*(1/5)²*(12.5)² = 1 - 11*0.04*156.25
        = 1 - 68.75 = -67.75 < 0 (負になる!)
    """
    coefficients = AsphericCoefficients(conic=10.0)
    result = calculate_sag(5.0, 25.0, coefficients)
    assert result is None


# =============================================================================
# 型検証テスト
# =============================================================================


def test_calculate_sag_invalid_radius_type() -> None:
    """radiusに無効な型を渡した場合にTypeErrorを送出することを検証する。"""
    with pytest.raises(TypeError, match="radiusは数値である必要があります"):
        calculate_sag("invalid", 20.0)  # type: ignore


def test_calculate_sag_invalid_diameter_type() -> None:
    """diameterに無効な型を渡した場合にTypeErrorを送出することを検証する。"""
    with pytest.raises(TypeError, match="diameterは数値である必要があります"):
        calculate_sag(50.0, "invalid")  # type: ignore


def test_calculate_sag_boolean_not_accepted_as_number() -> None:
    """真偽値が数値として受け入れられないことを検証する。

    Pythonではboolがintのサブクラスだが、この関数では明示的に拒否する。
    """
    with pytest.raises(TypeError):
        calculate_sag(True, 20.0)  # type: ignore

    with pytest.raises(TypeError):
        calculate_sag(50.0, False)  # type: ignore


# =============================================================================
# GlassWeight関数テスト
# =============================================================================


def _reference_glass_weight(
    radius1: float | None,
    radius2: float | None,
    thickness: float,
    specific_gravity: float,
    diameter1: float,
    diameter2: float,
    max_diameter: float,
    coefficients1: AsphericCoefficients | None = None,
    coefficients2: AsphericCoefficients | None = None,
    steps: int = 2000,
) -> float | None:
    """GlassWeightの参照値を高分割数で計算する。"""
    if coefficients1 is None:
        coefficients1 = AsphericCoefficients()
    if coefficients2 is None:
        coefficients2 = AsphericCoefficients()

    dh1 = (diameter1 / 2.0) / steps
    dh2 = (diameter2 / 2.0) / steps
    pi = math.pi

    def calculate_volume(
        radius: float | None,
        coefficients: AsphericCoefficients,
        dh: float,
        sign: float,
    ) -> float | None:
        volume = 0.0
        z_prev = 0.0
        for i in range(steps + 1):
            diameter = (dh * i) * 2.0
            sag = calculate_sag(radius, diameter, coefficients)
            if sag is None:
                return None
            z = sign * sag
            if i > 0:
                volume -= (
                    ((dh * steps) ** 2) * 2.0 - (dh * (i - 1)) ** 2 - (dh * i) ** 2
                ) * pi * (z - z_prev) / 2.0
            z_prev = z

        volume -= (((max_diameter / 2.0) ** 2) - ((dh * steps) ** 2)) * pi * z_prev
        return volume

    volume1 = calculate_volume(radius1, coefficients1, dh1, 1.0)
    if volume1 is None:
        return None
    volume2 = calculate_volume(radius2, coefficients2, dh2, -1.0)
    if volume2 is None:
        return None

    return (
        specific_gravity
        * (pi * ((max_diameter / 2.0) ** 2) * thickness + volume1 + volume2)
        / 1000.0
    )


def test_calculate_glass_weight_spherical_matches_reference() -> None:
    """球面レンズの重量が参照計算と一致することを検証する。"""
    weight = calculate_glass_weight(
        radius1=50.0,
        radius2=-50.0,
        thickness=5.0,
        specific_gravity=2.5,
        diameter1=40.0,
        diameter2=40.0,
        max_diameter=40.0,
    )
    expected = _reference_glass_weight(
        radius1=50.0,
        radius2=-50.0,
        thickness=5.0,
        specific_gravity=2.5,
        diameter1=40.0,
        diameter2=40.0,
        max_diameter=40.0,
    )
    assert expected is not None
    assert weight == pytest.approx(expected, rel=1e-6)


def test_calculate_glass_weight_aspheric_r1_matches_reference() -> None:
    """R1面のみ非球面の重量が参照計算と一致することを検証する。"""
    coefficients1 = AsphericCoefficients(conic=-0.5, a4=1e-6)
    weight = calculate_glass_weight(
        radius1=50.0,
        radius2=-50.0,
        thickness=5.0,
        specific_gravity=2.5,
        diameter1=40.0,
        diameter2=40.0,
        max_diameter=40.0,
        coefficients1=coefficients1,
    )
    expected = _reference_glass_weight(
        radius1=50.0,
        radius2=-50.0,
        thickness=5.0,
        specific_gravity=2.5,
        diameter1=40.0,
        diameter2=40.0,
        max_diameter=40.0,
        coefficients1=coefficients1,
    )
    assert expected is not None
    assert weight == pytest.approx(expected, rel=1e-6)


def test_calculate_glass_weight_aspheric_both_matches_reference() -> None:
    """両面非球面の重量が参照計算と一致することを検証する。"""
    coefficients1 = AsphericCoefficients(conic=-0.5, a4=1e-6)
    coefficients2 = AsphericCoefficients(conic=-0.8, a4=-2e-6)
    weight = calculate_glass_weight(
        radius1=50.0,
        radius2=-60.0,
        thickness=6.0,
        specific_gravity=2.6,
        diameter1=45.0,
        diameter2=45.0,
        max_diameter=45.0,
        coefficients1=coefficients1,
        coefficients2=coefficients2,
    )
    expected = _reference_glass_weight(
        radius1=50.0,
        radius2=-60.0,
        thickness=6.0,
        specific_gravity=2.6,
        diameter1=45.0,
        diameter2=45.0,
        max_diameter=45.0,
        coefficients1=coefficients1,
        coefficients2=coefficients2,
    )
    assert expected is not None
    assert weight == pytest.approx(expected, rel=1e-6)


def test_calculate_glass_weight_plane_and_sphere_matches_reference() -> None:
    """平面を含むレンズの重量が参照計算と一致することを検証する。"""
    weight = calculate_glass_weight(
        radius1=None,
        radius2=-80.0,
        thickness=4.0,
        specific_gravity=2.5,
        diameter1=35.0,
        diameter2=35.0,
        max_diameter=35.0,
    )
    expected = _reference_glass_weight(
        radius1=None,
        radius2=-80.0,
        thickness=4.0,
        specific_gravity=2.5,
        diameter1=35.0,
        diameter2=35.0,
        max_diameter=35.0,
    )
    assert expected is not None
    assert weight == pytest.approx(expected, rel=1e-6)


def test_calculate_glass_weight_flat_plate() -> None:
    """平行平板の重量が円柱体積と一致することを検証する。"""
    weight = calculate_glass_weight(
        radius1=None,
        radius2=None,
        thickness=5.0,
        specific_gravity=2.5,
        diameter1=40.0,
        diameter2=40.0,
        max_diameter=40.0,
    )
    expected = 2.5 * math.pi * (20.0**2) * 5.0 / 1000.0
    assert weight == pytest.approx(expected, rel=1e-7)


def test_calculate_glass_weight_scales_with_specific_gravity() -> None:
    """比重に比例して重量が変化することを検証する。"""
    base = calculate_glass_weight(
        radius1=None,
        radius2=None,
        thickness=4.0,
        specific_gravity=2.0,
        diameter1=30.0,
        diameter2=30.0,
        max_diameter=30.0,
    )
    doubled = calculate_glass_weight(
        radius1=None,
        radius2=None,
        thickness=4.0,
        specific_gravity=4.0,
        diameter1=30.0,
        diameter2=30.0,
        max_diameter=30.0,
    )
    assert base is not None
    assert doubled == pytest.approx(base * 2.0, rel=1e-7)


def test_calculate_glass_weight_extreme_specific_gravity() -> None:
    """比重の極小・極大値でも比例関係が維持されることを検証する。"""
    small = calculate_glass_weight(
        radius1=None,
        radius2=None,
        thickness=4.0,
        specific_gravity=0.001,
        diameter1=30.0,
        diameter2=30.0,
        max_diameter=30.0,
    )
    large = calculate_glass_weight(
        radius1=None,
        radius2=None,
        thickness=4.0,
        specific_gravity=20.0,
        diameter1=30.0,
        diameter2=30.0,
        max_diameter=30.0,
    )
    assert small is not None
    assert large == pytest.approx(small * 20000.0, rel=1e-7)


def test_calculate_glass_weight_increases_with_max_diameter() -> None:
    """最大外径が大きいほど重量が増加することを確認する。"""
    small = calculate_glass_weight(
        radius1=None,
        radius2=None,
        thickness=3.0,
        specific_gravity=2.5,
        diameter1=20.0,
        diameter2=20.0,
        max_diameter=20.0,
    )
    large = calculate_glass_weight(
        radius1=None,
        radius2=None,
        thickness=3.0,
        specific_gravity=2.5,
        diameter1=20.0,
        diameter2=20.0,
        max_diameter=25.0,
    )
    assert small is not None
    assert large is not None
    assert large > small


def test_calculate_glass_weight_returns_none_when_sag_invalid_r1() -> None:
    """R1面のサグ計算が失敗する場合はNoneを返すことを検証する。"""
    coefficients = AsphericCoefficients(conic=10.0)
    weight = calculate_glass_weight(
        radius1=5.0,
        radius2=None,
        thickness=3.0,
        specific_gravity=2.5,
        diameter1=25.0,
        diameter2=20.0,
        max_diameter=25.0,
        coefficients1=coefficients,
    )
    assert weight is None


def test_calculate_glass_weight_returns_none_when_sag_invalid_r2() -> None:
    """R2面のサグ計算が失敗する場合はNoneを返すことを検証する。"""
    coefficients = AsphericCoefficients(conic=10.0)
    weight = calculate_glass_weight(
        radius1=None,
        radius2=5.0,
        thickness=3.0,
        specific_gravity=2.5,
        diameter1=20.0,
        diameter2=25.0,
        max_diameter=25.0,
        coefficients2=coefficients,
    )
    assert weight is None


def test_calculate_glass_weight_zero_thickness_flat_plate() -> None:
    """厚み0の平行平板では重量が0になることを検証する。"""
    weight = calculate_glass_weight(
        radius1=None,
        radius2=None,
        thickness=0.0,
        specific_gravity=2.5,
        diameter1=40.0,
        diameter2=40.0,
        max_diameter=40.0,
    )
    assert weight == pytest.approx(0.0, rel=1e-7)
