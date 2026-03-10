"""pytest 配置文件和共享 fixtures。

本模块提供测试所需的共享配置和 fixture 对象，包括：
- 测试文档文件路径
- 临时目录管理
"""

import tempfile
from pathlib import Path
from typing import Generator

import pytest

PROJECT_ROOT = Path(__file__).parent.parent
FIXTURES_DIR = PROJECT_ROOT / "tests" / "fixtures"


@pytest.fixture
def fixtures_dir() -> Path:
    """获取测试 fixtures 目录路径。

    Returns:
        Path: fixtures 目录的 Path 对象
    """
    return FIXTURES_DIR


@pytest.fixture
def sample_txt_file(fixtures_dir: Path) -> Path:
    """获取示例 TXT 文件路径。

    Args:
        fixtures_dir: fixtures 目录路径

    Returns:
        Path: 示例 TXT 文件路径
    """
    return fixtures_dir / "sample.txt"


@pytest.fixture
def sample_txt_gbk_file(fixtures_dir: Path) -> Path:
    """获取 GBK 编码的示例 TXT 文件路径。

    Args:
        fixtures_dir: fixtures 目录路径

    Returns:
        Path: GBK 编码的示例 TXT 文件路径
    """
    return fixtures_dir / "sample_gbk.txt"


@pytest.fixture
def sample_docx_file(fixtures_dir: Path) -> Path:
    """获取示例 DOCX 文件路径。

    Args:
        fixtures_dir: fixtures 目录路径

    Returns:
        Path: 示例 DOCX 文件路径
    """
    return fixtures_dir / "sample.docx"


@pytest.fixture
def sample_pdf_file(fixtures_dir: Path) -> Path:
    """获取示例 PDF 文件路径。

    Args:
        fixtures_dir: fixtures 目录路径

    Returns:
        Path: 示例 PDF 文件路径
    """
    return fixtures_dir / "sample.pdf"


@pytest.fixture
def sample_excel_file(fixtures_dir: Path) -> Path:
    """获取示例 Excel 文件路径。

    Args:
        fixtures_dir: fixtures 目录路径

    Returns:
        Path: 示例 Excel 文件路径
    """
    return fixtures_dir / "sample.xlsx"


@pytest.fixture
def empty_txt_file(fixtures_dir: Path) -> Path:
    """获取空 TXT 文件路径。

    Args:
        fixtures_dir: fixtures 目录路径

    Returns:
        Path: 空 TXT 文件路径
    """
    return fixtures_dir / "empty.txt"


@pytest.fixture
def empty_docx_file(fixtures_dir: Path) -> Path:
    """获取空 DOCX 文件路径。

    Args:
        fixtures_dir: fixtures 目录路径

    Returns:
        Path: 空 DOCX 文件路径
    """
    return fixtures_dir / "empty.docx"


@pytest.fixture
def empty_pdf_file(fixtures_dir: Path) -> Path:
    """获取空 PDF 文件路径。

    Args:
        fixtures_dir: fixtures 目录路径

    Returns:
        Path: 空 PDF 文件路径
    """
    return fixtures_dir / "empty.pdf"


@pytest.fixture
def empty_excel_file(fixtures_dir: Path) -> Path:
    """获取空 Excel 文件路径。

    Args:
        fixtures_dir: fixtures 目录路径

    Returns:
        Path: 空 Excel 文件路径
    """
    return fixtures_dir / "empty.xlsx"


@pytest.fixture
def temp_document_dir() -> Generator[str, None, None]:
    """创建临时文档目录。

    Yields:
        str: 临时目录路径
    """
    with tempfile.TemporaryDirectory() as tmpdir:
        yield tmpdir
