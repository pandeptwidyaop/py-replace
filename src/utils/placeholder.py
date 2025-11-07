"""
Module untuk deteksi dan replacement placeholder
Format: ${nama_placeholder} untuk text, @{nama_placeholder} untuk image
"""
import re
from typing import List, Dict, Set, Tuple


class PlaceholderHandler:
    """Handler untuk mengelola placeholder dalam dokumen"""

    TEXT_PLACEHOLDER_PATTERN = r'\$\{([^}]+)\}'
    IMAGE_PLACEHOLDER_PATTERN = r'@\{([^}]+)\}'
    PLACEHOLDER_PATTERN = TEXT_PLACEHOLDER_PATTERN  # Backward compatibility

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

    @staticmethod
    def find_image_placeholders(text: str) -> Set[str]:
        """
        Menemukan semua image placeholder unik dalam teks (@{})

        Args:
            text: Teks yang akan dicari image placeholdernya

        Returns:
            Set dari nama image placeholder tanpa @{} wrapper
        """
        matches = re.findall(PlaceholderHandler.IMAGE_PLACEHOLDER_PATTERN, text)
        return set(matches)

    @staticmethod
    def find_all_placeholders_with_type(text: str) -> Tuple[Set[str], Set[str]]:
        """
        Menemukan semua placeholder (text dan image) dalam teks

        Args:
            text: Teks yang akan dicari placeholdernya

        Returns:
            Tuple (text_placeholders, image_placeholders)
        """
        text_placeholders = PlaceholderHandler.find_placeholders(text)
        image_placeholders = PlaceholderHandler.find_image_placeholders(text)
        return text_placeholders, image_placeholders

    @staticmethod
    def is_image_placeholder(placeholder_name: str, text: str) -> bool:
        """
        Check apakah placeholder adalah image placeholder

        Args:
            placeholder_name: Nama placeholder
            text: Teks untuk dicek

        Returns:
            True jika image placeholder, False jika text placeholder
        """
        pattern = r'@\{' + re.escape(placeholder_name) + r'\}'
        return bool(re.search(pattern, text))
