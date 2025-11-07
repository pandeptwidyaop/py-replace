"""
Module untuk handling image operations - download dan load images
"""
import os
import tempfile
from pathlib import Path
from typing import Tuple, Optional
import urllib.request
import urllib.error


class ImageHandler:
    """Handler untuk operasi image - local files dan URLs"""

    @staticmethod
    def is_url(path: str) -> bool:
        """
        Check apakah path adalah URL

        Args:
            path: String path atau URL

        Returns:
            True jika URL, False jika local path
        """
        return path.startswith(('http://', 'https://'))

    @staticmethod
    def download_image(url: str, timeout: int = 30) -> Tuple[Optional[str], str]:
        """
        Download image dari URL ke temporary file

        Args:
            url: URL image
            timeout: Timeout dalam detik

        Returns:
            Tuple (temp_file_path, error_message)
        """
        try:
            # Create temporary file
            suffix = Path(url).suffix or '.jpg'
            temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
            temp_path = temp_file.name
            temp_file.close()

            # Download image
            headers = {
                'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
            }
            request = urllib.request.Request(url, headers=headers)

            with urllib.request.urlopen(request, timeout=timeout) as response:
                with open(temp_path, 'wb') as f:
                    f.write(response.read())

            return temp_path, ""

        except urllib.error.URLError as e:
            return None, f"Failed to download image: {str(e)}"
        except Exception as e:
            return None, f"Error downloading image: {str(e)}"

    @staticmethod
    def validate_image_path(path: str) -> Tuple[bool, str]:
        """
        Validasi image path (local atau URL)

        Args:
            path: Path atau URL ke image

        Returns:
            Tuple (is_valid, error_message)
        """
        if not path or not path.strip():
            return False, "Image path is empty"

        # Check if URL
        if ImageHandler.is_url(path):
            # Basic URL validation
            if not path.startswith(('http://', 'https://')):
                return False, "Invalid URL format"
            return True, ""

        # Check if local file
        if not os.path.exists(path):
            return False, f"File not found: {path}"

        if not os.path.isfile(path):
            return False, f"Not a file: {path}"

        # Check file extension
        valid_extensions = {'.jpg', '.jpeg', '.png', '.gif', '.bmp', '.tiff', '.webp'}
        ext = Path(path).suffix.lower()
        if ext not in valid_extensions:
            return False, f"Invalid image format: {ext}. Supported: {', '.join(valid_extensions)}"

        return True, ""

    @staticmethod
    def get_image_path(path: str) -> Tuple[Optional[str], bool, str]:
        """
        Get valid image path - download jika URL, validasi jika local

        Args:
            path: Path atau URL ke image

        Returns:
            Tuple (final_path, is_temp_file, error_message)
        """
        # Validate first
        is_valid, error = ImageHandler.validate_image_path(path)
        if not is_valid:
            return None, False, error

        # If URL, download
        if ImageHandler.is_url(path):
            temp_path, error = ImageHandler.download_image(path)
            if error:
                return None, False, error
            return temp_path, True, ""

        # Local file - return as is
        return path, False, ""

    @staticmethod
    def cleanup_temp_file(path: str):
        """
        Cleanup temporary file

        Args:
            path: Path ke temporary file
        """
        try:
            if path and os.path.exists(path):
                os.unlink(path)
        except Exception:
            pass  # Ignore cleanup errors
