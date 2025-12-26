import math

import pytest

from src.optics.calculations import AsphericCoefficients, calculate_sag


def _expected_sag(
    radius: float | None,
    diameter: float,
    coefficients: AsphericCoefficients | None = None,
) -> float | None:
    if radius is None or radius == 0:
        c = 0.0
    else:
        c = 1.0 / radius

    h = diameter / 2.0

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


def test_calculate_sag_spherical_convex() -> None:
    result = calculate_sag(50.0, 20.0)
    expected = _expected_sag(50.0, 20.0)
    assert result == pytest.approx(expected)


def test_calculate_sag_spherical_concave() -> None:
    result = calculate_sag(-50.0, 20.0)
    expected = _expected_sag(-50.0, 20.0)
    assert result == pytest.approx(expected)


def test_calculate_sag_aspheric_conic() -> None:
    coefficients = AsphericCoefficients(conic=-0.5)
    result = calculate_sag(50.0, 20.0, coefficients)
    expected = _expected_sag(50.0, 20.0, coefficients)
    assert result == pytest.approx(expected)


def test_calculate_sag_aspheric_a4() -> None:
    coefficients = AsphericCoefficients(a4=1e-6)
    result = calculate_sag(50.0, 20.0, coefficients)
    expected = _expected_sag(50.0, 20.0, coefficients)
    assert result == pytest.approx(expected)


def test_calculate_sag_plane() -> None:
    result = calculate_sag(None, 20.0)
    assert result == 0.0


def test_calculate_sag_returns_none_when_argument_negative() -> None:
    result = calculate_sag(10.0, 100.0)
    assert result is None


def test_calculate_sag_zero_diameter() -> None:
    result = calculate_sag(50.0, 0.0)
    assert result == 0.0


def test_calculate_sag_invalid_diameter_type() -> None:
    with pytest.raises(TypeError):
        calculate_sag(50.0, "20")
