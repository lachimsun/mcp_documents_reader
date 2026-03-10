"""MCP 工具函数测试。

测试内容：
- read_document MCP 工具函数测试
"""

import os
from pathlib import Path
from unittest import mock

from mcp_documents_reader import read_document

FIXTURES_DIR = Path(__file__).parent / "fixtures"


class TestReadDocument:
    """read_document MCP 工具函数测试类。"""

    def test_read_document_txt_file(self) -> None:
        """测试读取 TXT 文档。"""
        file_path = FIXTURES_DIR / "sample.txt"
        result = read_document(str(file_path))

        assert "测试文本文件" in result
        assert "多行内容" in result

    def test_read_document_docx_file(self) -> None:
        """测试读取 DOCX 文档。"""
        file_path = FIXTURES_DIR / "sample.docx"
        result = read_document(str(file_path))

        assert "测试文档标题" in result

    def test_read_document_pdf_file(self) -> None:
        """测试读取 PDF 文档。"""
        file_path = FIXTURES_DIR / "sample.pdf"
        result = read_document(str(file_path))

        assert "test PDF document" in result

    def test_read_document_excel_file(self) -> None:
        """测试读取 Excel 文档。"""
        file_path = FIXTURES_DIR / "sample.xlsx"
        result = read_document(str(file_path))

        assert "Sheet" in result
        assert "姓名" in result

    def test_read_document_file_not_found(self) -> None:
        """测试读取不存在的文件。"""
        result = read_document("nonexistent.txt")

        assert "Error:" in result
        assert "not found" in result

    def test_read_document_unsupported_type(self) -> None:
        """测试读取不支持的文件类型。"""
        unsupported_file = FIXTURES_DIR / "test.unsupported"
        unsupported_file.write_text("test content")

        try:
            result = read_document(str(unsupported_file))

            assert "Error:" in result
            assert "Unsupported document type" in result
        finally:
            unsupported_file.unlink()

    def test_read_document_empty_file(self) -> None:
        """测试读取空文件。"""
        file_path = FIXTURES_DIR / "empty.txt"
        result = read_document(str(file_path))

        assert "No text found" in result

    def test_read_document_with_corrupted_file(self) -> None:
        """测试读取损坏的文件。"""
        file_path = FIXTURES_DIR / "corrupted.docx"
        result = read_document(str(file_path))

        assert "Error reading DOCX" in result

    def test_read_document_with_gbk_encoding(self) -> None:
        """测试读取 GBK 编码的文件。"""
        file_path = FIXTURES_DIR / "sample_gbk.txt"
        result = read_document(str(file_path))

        assert "GBK 编码" in result
        assert "中文内容" in result

    def test_read_document_with_special_characters_in_filename(
        self, temp_document_dir: str
    ) -> None:
        """测试文件名包含特殊字符。"""
        special_file = Path(temp_document_dir) / "test file (1).txt"
        special_file.write_text("special content", encoding="utf-8")

        result = read_document(str(special_file))

        assert "special content" in result


class TestReadDocumentWithPathTypes:
    """测试不同路径类型的 read_document。"""

    def test_read_document_with_absolute_path(self) -> None:
        """测试使用绝对路径读取文件。"""
        file_path = FIXTURES_DIR / "sample.txt"
        absolute_path = file_path.resolve()

        result = read_document(str(absolute_path))

        assert "测试文本文件" in result

    def test_read_document_with_relative_path(self) -> None:
        """测试使用相对路径读取文件。"""
        original_cwd = os.getcwd()
        try:
            os.chdir(FIXTURES_DIR)
            result = read_document("sample.txt")

            assert "测试文本文件" in result
        finally:
            os.chdir(original_cwd)

    def test_read_document_with_path_object(self) -> None:
        """测试使用 Path 对象读取文件。"""
        file_path = FIXTURES_DIR / "sample.txt"
        result = read_document(str(file_path))

        assert "测试文本文件" in result


class TestReadDocumentWithMockedFilesystem:
    """使用 mock 文件系统的 read_document 测试类。"""

    @mock.patch("mcp_documents_reader.Path.exists")
    @mock.patch("mcp_documents_reader.DocumentReaderFactory.is_supported")
    @mock.patch("mcp_documents_reader.DocumentReaderFactory.get_reader")
    def test_read_document_calls_reader_correctly(
        self,
        mock_get_reader: mock.MagicMock,
        mock_is_supported: mock.MagicMock,
        mock_exists: mock.MagicMock,
    ) -> None:
        """测试 read_document 正确调用 Reader。"""
        mock_exists.return_value = True
        mock_is_supported.return_value = True
        mock_reader = mock.MagicMock()
        mock_reader.read.return_value = "test content"
        mock_get_reader.return_value = mock_reader

        result = read_document("test.txt")

        mock_exists.assert_called_once()
        mock_is_supported.assert_called_once()
        mock_get_reader.assert_called_once()
        mock_reader.read.assert_called_once()

        assert result == "test content"

    @mock.patch("mcp_documents_reader.Path.exists")
    def test_read_document_file_not_exists_mock(
        self, mock_exists: mock.MagicMock
    ) -> None:
        """测试文件不存在的情况（使用 mock）。"""
        mock_exists.return_value = False

        result = read_document("test.txt")

        assert "Error:" in result
        assert "not found" in result

    @mock.patch("mcp_documents_reader.Path.exists")
    @mock.patch("mcp_documents_reader.DocumentReaderFactory.is_supported")
    def test_read_document_unsupported_type_mock(
        self, mock_is_supported: mock.MagicMock, mock_exists: mock.MagicMock
    ) -> None:
        """测试不支持的文件类型（使用 mock）。"""
        mock_exists.return_value = True
        mock_is_supported.return_value = False

        result = read_document("test.xyz")

        assert "Error:" in result
        assert "Unsupported document type" in result

    @mock.patch("mcp_documents_reader.Path.exists")
    @mock.patch("mcp_documents_reader.DocumentReaderFactory.is_supported")
    @mock.patch("mcp_documents_reader.DocumentReaderFactory.get_reader")
    def test_read_document_reader_exception(
        self,
        mock_get_reader: mock.MagicMock,
        mock_is_supported: mock.MagicMock,
        mock_exists: mock.MagicMock,
    ) -> None:
        """测试 Reader 抛出异常的情况。"""
        mock_exists.return_value = True
        mock_is_supported.return_value = True
        mock_get_reader.side_effect = Exception("Reader error")

        result = read_document("test.txt")

        assert "Error reading document" in result
