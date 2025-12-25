# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Communication Guidelines

**IMPORTANT**: When interacting with users in this repository, always respond in polite Japanese (丁寧な日本語).

## Project Overview

This is a Python-based refactoring project aimed at replacing a legacy Excel VBA tool (`図面作成TOOL.xlam`) used for optical lens manufacturing drawing automation. The original VBA tool controls a 2D CAD software "図脳RAPID" to automatically generate technical drawings from optical design data (CODE V output). This project seeks to modernize the system while preserving complex optical calculation logic and CAD automation capabilities.

**Business Context**: The original VBA tool is heavily person-dependent and contains critical engineering knowledge (lens shape calculations, tolerance settings) that must be preserved. This refactoring addresses business continuity risks.

## Development Commands

### Environment Setup
- `uv sync` - Install and synchronize dependencies from pyproject.toml and uv.lock

### Running the Application
- `uv run python src/main.py` - Execute the main application

### Testing
- `uv run pytest` - Run all tests
- `uv run pytest tests/test_specific.py` - Run a specific test file
- `uv run pytest tests/test_specific.py::test_function_name` - Run a single test

### Code Quality
- `uv run ruff check src tests` - Run static analysis/linting
- `uv run ruff check src tests --fix` - Auto-fix linting issues
- `uv run ruff format src tests` - Format code to standard style

## Architecture and Structure

### Current System Analysis

The legacy VBA system (`references/VBA/`) contains approximately 40+ modules organized into:

1. **Data Processing Layer** (`clsOpticalData`, `clsGlassData`, `clsToleranceData`)
   - Parses CODE V sequence files (.seq) using regex
   - Converts optical design data to part sheets
   - Auto-calculates tolerances (Newton rings, astigmatism, thickness, diameter)

2. **Optical Calculation Layer** (`basUDF.bas`)
   - **Critical**: Contains core optical physics formulas
   - `Sag`: Calculates lens surface sag (depth) for spherical and aspherical surfaces
   - `GlassWeight`: Computes lens weight using integral calculation
   - `FocalLength`: Calculates single lens focal length
   - `PressValue`: Calculates press lens centering allowance
   - These functions represent proprietary manufacturing knowledge

3. **CAD Control Layer** (`clsRapid*` classes - 15+ modules)
   - Controls 図脳RAPID via COM API (pywin32 equivalent needed)
   - Draws lens profiles (arcs for spherical, polylines for aspherical)
   - Adds dimensions, tolerances, geometric tolerances (GD&T)
   - Generates assembly drawings, part drawings, and press drawings

4. **Material Database Integration** (`clsGlaDB`, `clsMaterialCode`)
   - Manages glass material properties database
   - Handles material substitution and code mapping

### Migration Strategy

**Phase 1 (Months 1-2)**: Core Logic Migration
- Port optical calculation functions from VBA to Python
- Implement unit tests with numerical validation against VBA output
- Parse .seq file format

**Phase 2 (Months 2-3)**: CAD Integration PoC
- Use `pywin32` to control 図脳RAPID COM interface
- Validate drawing primitives (lines, arcs, text)

**Phase 3 (Months 1-2)**: System Integration
- Build CLI or simple GUI
- Parallel operation with legacy system for validation

**Phase 4**: Full Migration
- Complete transition from VBA system
- Documentation and training

### Dependencies

Key Python libraries used:
- `pandas`: Data manipulation and processing
- `openpyxl`: Excel file handling
- `pywin32`: Windows COM API for CAD control (図脳RAPID)
- `jupyterlab`: Analysis and prototyping notebooks
- `pytest`: Testing framework
- `ruff`: Linting and formatting

## Coding Standards

### Language and Documentation
- **All docstrings, comments, and documentation MUST be in Japanese**
- Use Google-style docstrings
- Type hints are mandatory for all function arguments and return values

### Naming Conventions
- Classes: `PascalCase`
- Functions, methods, variables: `snake_case`
- Constants: `UPPER_CASE_SNAKE_CASE`

### Design Principles
- **Prefer `@dataclass` over dictionaries** for type safety, IDE support, and maintainability
- Use `abc.ABC` for shared interfaces/abstract base classes
- Avoid magic numbers - use named constants (e.g., `EPSILON = 1e-12`)
- Use `logging` module instead of `print()` for debug output
- Raise appropriate exceptions (`ValueError`, `TypeError`) for invalid inputs

### Example Code Style
```python
def refract_ray(
    incident_ray_direction: list[float],
    surface_normal: list[float],
    n_incident: float,
    n_refracted: float,
) -> list[float] | None:
    """屈折光線の方向を計算する。

    Args:
        incident_ray_direction (list[float]): 入射光線の方向ベクトル。
        surface_normal (list[float]): 表面の法線ベクトル。
        n_incident (float): 入射側の屈折率。
        n_refracted (float): 屈折側の屈折率。

    Returns:
        list[float] | None: 屈折光線の方向ベクトル。全反射の場合はNoneを返す。
    """
    pass
```

## Testing Guidelines

- Framework: `pytest`
- Test files: `test_*.py` format in `tests/` directory
- Use `@pytest.fixture` for shared test data setup
- When porting VBA calculations, create test cases that validate Python output matches VBA output exactly

## Version Control and Commits

### Commit Message Format
Follow Conventional Commits with Japanese body:
- Format: `<type>(<scope>): <subject in Japanese>`
- Examples:
  - `docs(readme): セットアップ手順を追記`
  - `feat(optics): サグ量計算関数を実装`
  - `fix(parser): .seqファイルのパース処理を修正`

### Pull Request Requirements
Include in PR description:
- Summary of changes
- Tests performed
- Related issue links
- For notebook/analysis changes: attach screenshots or output

## Important Notes

### Python Version
- Fixed to Python 3.12 (see `.python-version`)

### Security
- Never commit secrets, API keys, or sensitive data
- Use environment variables or local config files for sensitive information
- `data/` folder is for local input data - avoid committing large or confidential files

### Directory Structure
- `src/` - Application source code (entry point: `src/main.py`)
- `notebooks/` - Jupyter notebooks for analysis and validation
- `data/` - Local input data (not committed)
- `docs/` - Design documents and specifications
  - `docs/design/VBA_Module_Reference.md` - Comprehensive VBA module documentation
  - `docs/planning/project_proposal.md` - Project background and roadmap
- `references/VBA/` - Legacy VBA source code for reference
- `tests/` - Test suite

### Critical Reference Documents
- `.github/copilot-instructions.md` - Detailed coding standards
- `AGENTS.md` - Repository operation guidelines
- `docs/design/VBA_Module_Reference.md` - Essential reading for understanding the VBA system architecture and logic to be migrated
