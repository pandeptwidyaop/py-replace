"""
Module untuk deteksi dan replacement placeholder dalam format ${nama_placeholder}
"""
import re
from typing import List, Dict, Set


class PlaceholderHandler:
    """Handler untuk mengelola placeholder dalam dokumen"""

    PLACEHOLDER_PATTERN = r'\$\{([^}]+)\}'

    @staticmethod
    def find_placeholders(text: str) -> Set[str]:
        """
        Menemukan semua placeholder unik dalam teks

        Args:
            text: Teks yang akan dicari placeholdernya

        Returns:
            Set dari nama placeholder tanpa ${} wrapper
        """
        matches = re.findall(PlaceholderHandler.PLACEHOLDER_PATTERN, text)
        return set(matches)

    @staticmethod
    def replace_placeholder(text: str, placeholder: str, value: str) -> str:
        """
        Mengganti placeholder dengan nilai tertentu

        Args:
            text: Teks yang berisi placeholder
            placeholder: Nama placeholder (tanpa ${})
            value: Nilai pengganti

        Returns:
            Teks dengan placeholder yang sudah diganti
        """
        pattern = r'\$\{' + re.escape(placeholder) + r'\}'
        return re.sub(pattern, value, text)

    @staticmethod
    def replace_all_placeholders(text: str, replacements: Dict[str, str]) -> str:
        """
        Mengganti semua placeholder sesuai dictionary replacements

        Args:
            text: Teks yang berisi placeholder
            replacements: Dictionary mapping placeholder -> nilai pengganti

        Returns:
            Teks dengan semua placeholder yang sudah diganti
        """
        result = text
        for placeholder, value in replacements.items():
            result = PlaceholderHandler.replace_placeholder(result, placeholder, value)
        return result

    @staticmethod
    def validate_placeholder_name(name: str) -> bool:
        """
        Validasi nama placeholder (harus alphanumeric dan underscore)

        Args:
            name: Nama placeholder untuk divalidasi

        Returns:
            True jika valid, False jika tidak
        """
        return bool(re.match(r'^[a-zA-Z0-9_]+$', name))
