"""
Module untuk load config dari CSV atau XLSX
Config format: 2 kolom (placeholder, value)
"""
import pandas as pd
from typing import Dict, Tuple
from pathlib import Path


class ConfigLoader:
    """Handler untuk load config dari CSV/XLSX"""

    SUPPORTED_FORMATS = ['.csv', '.xlsx', '.xls']

    @staticmethod
    def load_config(file_path: str) -> Tuple[Dict[str, str], str]:
        """
        Load config dari file CSV atau XLSX

        Args:
            file_path: Path ke file config

        Returns:
            Tuple (Dictionary mapping placeholder -> value, error message if any)
        """
        try:
            file_ext = Path(file_path).suffix.lower()

            if file_ext not in ConfigLoader.SUPPORTED_FORMATS:
                return {}, f"Unsupported file format: {file_ext}. Use CSV or XLSX."

            # Load file
            if file_ext == '.csv':
                df = pd.read_csv(file_path)
            else:  # .xlsx or .xls
                df = pd.read_excel(file_path)

            # Validate columns
            if df.shape[1] < 2:
                return {}, "Config file must have at least 2 columns (placeholder, value)"

            # Get first 2 columns
            df = df.iloc[:, :2]
            df.columns = ['placeholder', 'value']

            # Remove rows with empty placeholder
            df = df[df['placeholder'].notna()]
            df = df[df['placeholder'].astype(str).str.strip() != '']

            # Convert to dictionary
            config = {}
            for _, row in df.iterrows():
                placeholder = str(row['placeholder']).strip()
                value = str(row['value']) if pd.notna(row['value']) else ''

                # Remove ${} if present in config file
                if placeholder.startswith('${') and placeholder.endswith('}'):
                    placeholder = placeholder[2:-1]

                config[placeholder] = value

            return config, ""

        except Exception as e:
            return {}, f"Failed to load config: {str(e)}"

    @staticmethod
    def validate_config(config: Dict[str, str], placeholders: set) -> Tuple[bool, str, list, list]:
        """
        Validasi config terhadap placeholder yang ditemukan

        Args:
            config: Dictionary config yang diload
            placeholders: Set placeholder yang ditemukan di dokumen

        Returns:
            Tuple (is_valid, message, missing_in_config, extra_in_config)
        """
        config_keys = set(config.keys())

        missing_in_config = placeholders - config_keys
        extra_in_config = config_keys - placeholders

        if not missing_in_config and not extra_in_config:
            return True, "Config matches all placeholders perfectly!", [], []

        warnings = []
        if missing_in_config:
            warnings.append(
                f"Missing in config: {', '.join(sorted(missing_in_config))}"
            )
        if extra_in_config:
            warnings.append(
                f"Extra in config (will be ignored): {', '.join(sorted(extra_in_config))}"
            )

        return False, "\n".join(warnings), list(missing_in_config), list(extra_in_config)

    @staticmethod
    def save_template(file_path: str, placeholders: list) -> Tuple[bool, str]:
        """
        Save template config file dengan placeholder yang ditemukan

        Args:
            file_path: Path output file
            placeholders: List placeholder untuk template

        Returns:
            Tuple (success, error_message)
        """
        try:
            file_ext = Path(file_path).suffix.lower()

            # Create DataFrame
            df = pd.DataFrame({
                'placeholder': [f"${{{p}}}" for p in sorted(placeholders)],
                'value': [''] * len(placeholders)
            })

            # Save based on format
            if file_ext == '.csv':
                df.to_csv(file_path, index=False)
            elif file_ext in ['.xlsx', '.xls']:
                df.to_excel(file_path, index=False)
            else:
                return False, f"Unsupported format: {file_ext}"

            return True, ""

        except Exception as e:
            return False, f"Failed to save template: {str(e)}"
