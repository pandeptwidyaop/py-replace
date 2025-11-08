"""
Entry point untuk DOCX Placeholder Replacer Application
"""
import sys
from pathlib import Path

# Add src directory to path for PyInstaller
if getattr(sys, 'frozen', False):
    # Running as compiled executable
    src_dir = Path(sys._MEIPASS)
else:
    # Running as script
    src_dir = Path(__file__).parent

sys.path.insert(0, str(src_dir))

from gui.app import main

if __name__ == "__main__":
    main()
