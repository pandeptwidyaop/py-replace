"""
Module untuk handling file DOCX - load, scan, dan save
"""
from docx import Document
from docx.text.paragraph import Paragraph
from docx.table import Table, _Cell
from typing import Set, Dict, List
from .placeholder import PlaceholderHandler


class DocxHandler:
    """Handler untuk operasi file DOCX"""

    def __init__(self, file_path: str = None):
        """
        Inisialisasi DocxHandler

        Args:
            file_path: Path ke file DOCX yang akan diload
        """
        self.file_path = file_path
        self.document = None
        if file_path:
            self.load(file_path)

    def load(self, file_path: str):
        """
        Load dokumen DOCX

        Args:
            file_path: Path ke file DOCX
        """
        self.file_path = file_path
        self.document = Document(file_path)

    def find_all_placeholders(self) -> Set[str]:
        """
        Menemukan semua placeholder dalam dokumen

        Returns:
            Set dari nama placeholder yang ditemukan
        """
        if not self.document:
            return set()

        placeholders = set()

        # Scan paragraphs
        for paragraph in self.document.paragraphs:
            placeholders.update(
                PlaceholderHandler.find_placeholders(paragraph.text)
            )

        # Scan tables
        for table in self.document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        placeholders.update(
                            PlaceholderHandler.find_placeholders(paragraph.text)
                        )

        # Scan headers
        for section in self.document.sections:
            header = section.header
            for paragraph in header.paragraphs:
                placeholders.update(
                    PlaceholderHandler.find_placeholders(paragraph.text)
                )

            # Scan footers
            footer = section.footer
            for paragraph in footer.paragraphs:
                placeholders.update(
                    PlaceholderHandler.find_placeholders(paragraph.text)
                )

        return placeholders

    def replace_placeholders(self, replacements: Dict[str, str]):
        """
        Mengganti semua placeholder dalam dokumen

        Args:
            replacements: Dictionary mapping placeholder -> nilai pengganti
        """
        if not self.document:
            return

        # Replace in paragraphs
        for paragraph in self.document.paragraphs:
            self._replace_in_paragraph(paragraph, replacements)

        # Replace in tables
        for table in self.document.tables:
            for row in table.rows:
                for cell in row.cells:
                    for paragraph in cell.paragraphs:
                        self._replace_in_paragraph(paragraph, replacements)

        # Replace in headers and footers
        for section in self.document.sections:
            # Headers
            for paragraph in section.header.paragraphs:
                self._replace_in_paragraph(paragraph, replacements)

            # Footers
            for paragraph in section.footer.paragraphs:
                self._replace_in_paragraph(paragraph, replacements)

    def _replace_in_paragraph(self, paragraph: Paragraph, replacements: Dict[str, str]):
        """
        Mengganti placeholder dalam satu paragraph dengan mempertahankan formatting

        Args:
            paragraph: Paragraph object dari python-docx
            replacements: Dictionary mapping placeholder -> nilai pengganti
        """
        # Kumpulkan semua runs menjadi satu teks
        full_text = ''.join(run.text for run in paragraph.runs)

        # Lakukan replacement
        new_text = PlaceholderHandler.replace_all_placeholders(full_text, replacements)

        # Jika ada perubahan, update paragraph
        if full_text != new_text:
            # Hapus semua runs kecuali yang pertama
            for i in range(len(paragraph.runs) - 1, 0, -1):
                paragraph._element.remove(paragraph.runs[i]._element)

            # Update run pertama dengan teks baru
            if paragraph.runs:
                paragraph.runs[0].text = new_text
            else:
                paragraph.text = new_text

    def save(self, output_path: str):
        """
        Simpan dokumen ke file

        Args:
            output_path: Path output file
        """
        if self.document:
            self.document.save(output_path)

    def get_document_info(self) -> Dict[str, any]:
        """
        Mendapatkan informasi dokumen

        Returns:
            Dictionary berisi informasi dokumen
        """
        if not self.document:
            return {}

        return {
            'paragraphs': len(self.document.paragraphs),
            'tables': len(self.document.tables),
            'sections': len(self.document.sections),
        }
