import pytest

from src.optics.calculations import AsphericCoefficients, calculate_sag

# =============================================================================
# テストデータ: CODE Vまたはレガシー実装から取得した期待値
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
