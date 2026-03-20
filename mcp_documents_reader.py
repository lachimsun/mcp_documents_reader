import os
from abc import ABC, abstractmethod
from pathlib import Path

from docx import Document as DocxDocument
from mcp.server.fastmcp import FastMCP, Image
from openpyxl import load_workbook

import fitz  # PyMuPDF
from typing import Optional, List, Union, Any

from typing_extensions import override

mcp = FastMCP("Document Reader")


class DocumentReader(ABC):
    """Abstract base class for document readers"""

    @abstractmethod
    def read(self, file_path: str, start_page: int = 1, end_page: Optional[int] = None) -> Any:
        """Read and extract text and images from a document with pagination support"""
        pass


class DocxReader(DocumentReader):
    """DOCX document reader implementation with pagination and image support"""

    @override
    def read(self, file_path: str, start_page: int = 1, end_page: Optional[int] = None) -> Any:
        """
        Read and extract text/images from DOCX file.
        Treats structural elements as items for pagination.
        """
        try:
            doc = DocxDocument(file_path)
            content_results: List[Union[str, Image]] = []
            
            # Temporary storage for all text items to apply pagination
            text_items = []
            for paragraph in doc.paragraphs:
                if paragraph.text.strip():
                    text_items.append(paragraph.text.strip())
            
            for table in doc.tables:
                for row in table.rows:
                    row_text = []
                    for cell in row.cells:
                        cell_text = " ".join([p.text for p in cell.paragraphs]).strip()
                        if cell_text:
                            row_text.append(cell_text)
                    if row_text:
                        text_items.append("\t".join(row_text))

            # Pagination logic
            items_per_page = 20
            total_items = len(text_items)
            total_pages = (total_items + items_per_page - 1) // items_per_page
            
            start_idx = (start_page - 1) * items_per_page
            actual_end_page = end_page if end_page else start_page + 4
            end_idx = min(actual_end_page * items_per_page, total_items)

            if start_idx >= total_items and total_items > 0:
                return [f"Requested start page {start_page} is out of range. Total virtual pages: {total_pages}"]

            content_results.append(f"--- INFO: DOCX contains approx {total_pages} virtual pages. Reading pages {start_page}-{min(actual_end_page, total_pages)} ---")
            
            # Add text content for the range
            page_text = "\n".join(text_items[start_idx:end_idx])
            if page_text:
                content_results.append(page_text)

            # Extract images from DOCX
            image_count = 0
            for rel in doc.part.rels.values():
                if "image" in rel.target_ref:
                    image_count += 1
                    # Only send images if we are on the first requested chunk or it's a small set
                    if start_page == 1 and image_count <= 10:
                        content_results.append(Image(data=rel.target_part.blob, format="image/png"))

            if end_idx < total_items:
                content_results.append(f"\n...[NOTE: More text content available. Use start_page={actual_end_page + 1} to continue]...")

            return content_results if (text_items or image_count > 0) else ["No content found in the DOCX."]
        except Exception as e:
            return [f"Error reading DOCX: {str(e)}"]


class PdfReader(DocumentReader):
    """Advanced PDF reader supporting tables and embedded images"""

    @override
    def read(self, file_path: str, start_page: int = 1, end_page: Optional[int] = None) -> Any:
        try:
            doc = fitz.open(file_path)
            total_pages = len(doc)

            start_idx = max(0, start_page - 1)
            end_idx = min(end_page if end_page else total_pages, total_pages)

            results: List[Union[str, Image]] = []
            results.append(f"--- INFO: Document has {total_pages} pages. Reading {start_idx+1}-{end_idx} ---")

            for page_num in range(start_idx, end_idx):
                page = doc[page_num]
                
                # 1. Extract Text
                blocks = page.get_text("blocks")
                blocks.sort(key=lambda b: (b[1], b[0]))
                
                page_header = f"\n[PAGE {page_num + 1}]\n"
                text_content = []
                for block in blocks:
                    if block[4].strip():
                        text_content.append(block[4].strip())
                
                results.append(page_header + "\n".join(text_content))

                # 2. Extract Images
                image_list = page.get_images(full=True)
                for img_index, img in enumerate(image_list):
                    xref = img[0]
                    base_image = doc.extract_image(xref)
                    image_bytes = base_image["image"]
                    image_ext = base_image["ext"]
                    
                    results.append(Image(data=image_bytes, format=f"image/{image_ext}"))
                
                results.append("-" * 20)

            doc.close()
            return results

        except Exception as e:
            return [f"Error reading PDF (PyMuPDF): {str(e)}"]


class TxtReader(DocumentReader):
    """TXT document reader implementation"""

    @override
    def read(self, file_path: str, start_page: int = 1, end_page: Optional[int] = None) -> Any:
        encodings = ["utf-8", "gbk", "gb2312", "latin-1"]
        for encoding in encodings:
            try:
                with open(file_path, "r", encoding=encoding) as f:
                    lines = f.readlines()
                
                lines_per_page = 100
                total_lines = len(lines)
                start_idx = (start_page - 1) * lines_per_page
                actual_end_page = end_page if end_page else start_page + 4
                end_idx = min(actual_end_page * lines_per_page, total_lines)

                if not lines:
                    return ["No text found in the TXT file."]

                chunk = "".join(lines[start_idx:end_idx])
                if end_idx < total_lines:
                    chunk += f"\n\n...[NOTE: Large TXT file. Showing lines {start_idx}-{end_idx} of {total_lines}]..."
                
                return [chunk]
            except UnicodeDecodeError:
                continue
            except Exception as e:
                return [f"Error reading TXT: {str(e)}"]

        return ["Error reading TXT: Could not decode file."]


class ExcelReader(DocumentReader):
    """Excel document reader implementation"""

    @override
    def read(self, file_path: str, start_page: int = 1, end_page: Optional[int] = None) -> Any:
        try:
            wb = load_workbook(file_path, read_only=True)
            text = []
            for sheet_name in wb.sheetnames:
                sheet = wb[sheet_name]
                text.append(f"=== Sheet: {sheet_name} ===")
                for row in sheet.iter_rows(values_only=True):
                    row_text = [str(cell) if cell is not None else "" for cell in row]
                    if any(row_text):
                        text.append("\t".join(row_text))
            
            full_text = "\n".join(text)
            wb.close()
            
            if len(full_text) > 30000:
                full_text = full_text[:30000] + "\n\n...[WARNING: Excel output truncated]..."
            
            return [full_text] if full_text.strip() else ["No text found in the Excel file."]
        except Exception as e:
            return [f"Error reading Excel: {str(e)}"]


class DocumentReaderFactory:
    """Factory for creating document readers based on file extension"""

    _readers: dict[str, type[DocumentReader]] = {
        ".txt": TxtReader,
        ".tex": TxtReader,
        ".docx": DocxReader,
        ".pdf": PdfReader,
        ".xlsx": ExcelReader,
        ".xls": ExcelReader,
    }

    @classmethod
    def get_reader(cls, file_path: str) -> DocumentReader:
        _, ext = os.path.splitext(file_path.lower())
        if ext not in cls._readers:
            raise ValueError(f"Unsupported document type: {ext}")
        return cls._readers[ext]()

    @classmethod
    def is_supported(cls, file_path: str) -> bool:
        _, ext = os.path.splitext(file_path.lower())
        return ext in cls._readers


@mcp.tool()
def read_document(filename: str, start_page: int = 1, end_page: Optional[int] = None) -> Any:
    """
    Reads and extracts text and images from a document (PDF, DOCX, TXT, Excel).
    Images are returned as visual objects that the model can analyze directly.

    :param filename: Path to the document file.
    :param start_page: Page/chunk number to start from (default 1).
    :param end_page: Page/chunk number to stop at.
    """
    file_path = Path(filename)

    if not file_path.exists():
        return [f"Error: File '{filename}' not found."]

    if not DocumentReaderFactory.is_supported(str(file_path)):
        return [f"Error: Unsupported document type for file '{filename}'."]

    try:
        reader = DocumentReaderFactory.get_reader(str(file_path))
        return reader.read(str(file_path), start_page=start_page, end_page=end_page)
    except Exception as e:
        return [f"Error reading document: {str(e)}"]


def main():
    mcp.run()


if __name__ == "__main__":
    main()