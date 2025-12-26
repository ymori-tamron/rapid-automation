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

# VBA計算値との比較テスト
# 注: expected_weight の値はVBAマクロ (basUDF.GlassWeight) で計算した値を入力してください


@pytest.mark.parametrize(
    "radius1,radius2,thickness,specific_gravity,diameter1,diameter2,max_diameter,"
    "coefficients1,coefficients2,expected_weight,test_id",
    [
        # 球面レンズ（両面球面）
        (
            50.0,
            -50.0,
            5.0,
            2.5,
            40.0,
            40.0,
            40.0,
            None,
            None,
            2.78456744399565,
            "spherical_both",
        ),
        # 非球面レンズ（R1面のみ非球面）
        (
            50.0,
            -50.0,
            5.0,
            2.5,
            40.0,
            40.0,
            40.0,
            AsphericCoefficients(conic=-0.5, a4=1e-6),
            None,
            2.70910802866498,
            "aspheric_r1_only",
        ),
        # 非球面レンズ（両面非球面）
        (
            50.0,
            -60.0,
            6.0,
            2.6,
            45.0,
            45.0,
            45.0,
            AsphericCoefficients(conic=-0.5, a4=1e-6),
            AsphericCoefficients(conic=-0.8, a4=-2e-6),
            4.33624166550092,
            "aspheric_both",
        ),
        # 平面+球面レンズ
        (
            None,
            -80.0,
            4.0,
            2.5,
            35.0,
            35.0,
            35.0,
            None,
            None,
            7.30049943952618,
            "plane_and_sphere",
        ),
        # 平行平板（両面平面）
        (
            None,
            None,
            5.0,
            2.5,
            40.0,
            40.0,
            40.0,
            None,
            None,
            15.707963267949,
            "flat_plate",
        ),
        # エッジケース: 比重が極小
        (
            None,
            None,
            4.0,
            0.01,
            30.0,
            30.0,
            30.0,
            None,
            None,
            0.0282743338823081,
            "edge_specific_gravity_min",
        ),
        # エッジケース: 比重が極大
        (
            None,
            None,
            4.0,
            20.0,
            30.0,
            30.0,
            30.0,
            None,
            None,
            56.5486677646163,
            "edge_specific_gravity_max",
        ),
        # エッジケース: 厚みが極小
        (
            None,
            None,
            0.1,
            2.5,
            40.0,
            40.0,
            40.0,
            None,
            None,
            0.314159265358979,
            "edge_thickness_min",
        ),
        # 有効径が異なるケース: R1とR2で有効径が異なる両面球面
        (
            50.0,
            -80.0,
            5.0,
            2.5,
            40.0,
            35.0,
            40.0,
            None,
            None,
            5.49901692279008,
            "different_effective_diameters",
        ),
        # 最大外径が有効径より大きいケース: コバ部分がある
        (
            50.0,
            -50.0,
            5.0,
            2.5,
            30.0,
            30.0,
            35.0,
            None,
            None,
            5.04926952118464,
            "with_edge_margin",
        ),
        # すべてのパラメータが異なるケース
        (
            50.0,
            -60.0,
            6.0,
            2.5,
            40.0,
            35.0,
            42.0,
            None,
            None,
            7.10061027912313,
            "all_parameters_different",
        ),
        # 非球面で有効径が異なるケース
        (
            50.0,
            -60.0,
            5.5,
            2.6,
            38.0,
            34.0,
            40.0,
            AsphericCoefficients(conic=-0.5, a4=1e-6),
            None,
            6.12950112280577,
            "aspheric_different_diameters",
        ),
    ],
    ids=lambda params: params[-1] if isinstance(params, tuple) else None,
)
def test_calculate_glass_weight_matches_vba(
    radius1: float | None,
    radius2: float | None,
    thickness: float,
    specific_gravity: float,
    diameter1: float,
    diameter2: float,
    max_diameter: float,
    coefficients1: AsphericCoefficients | None,
    coefficients2: AsphericCoefficients | None,
    expected_weight: float,
    test_id: str,
) -> None:
    """VBAで計算した重量値と一致することを検証する（相対誤差1e-6以下）。"""
    weight = calculate_glass_weight(
        radius1=radius1,
        radius2=radius2,
        thickness=thickness,
        specific_gravity=specific_gravity,
        diameter1=diameter1,
        diameter2=diameter2,
        max_diameter=max_diameter,
        coefficients1=coefficients1,
        coefficients2=coefficients2,
    )
    assert weight is not None
    assert weight == pytest.approx(expected_weight, rel=1e-6)


def test_calculate_glass_weight_flat_plate_matches_cylinder_volume() -> None:
    """平行平板の重量が円柱体積の計算式と一致することを検証する。"""
    weight = calculate_glass_weight(
        radius1=None,
        radius2=None,
        thickness=5.0,
        specific_gravity=2.5,
        diameter1=40.0,
        diameter2=40.0,
        max_diameter=40.0,
    )
    # 円柱体積 = π * r^2 * h, 重量 = 比重 * 体積[cm^3]
    expected = 2.5 * math.pi * (20.0**2) * 5.0 / 1000.0
    assert weight == pytest.approx(expected, rel=1e-7)


def test_calculate_glass_weight_scales_with_specific_gravity() -> None:
    """比重に比例して重量が変化することを検証する（物理的性質の確認）。"""
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


def test_calculate_glass_weight_increases_with_max_diameter() -> None:
    """最大外径が大きいほど重量が増加することを確認する（物理的傾向の検証）。"""
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
    """R1面のサグ計算が失敗する場合はNoneを返すことを検証する（異常系）。"""
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
    """R2面のサグ計算が失敗する場合はNoneを返すことを検証する（異常系）。"""
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
    """厚み0の平行平板では重量が0になることを検証する（境界値テスト）。"""
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
