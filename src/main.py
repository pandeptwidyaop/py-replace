"""
Entry point untuk DOCX Placeholder Replacer Application
"""
import sys
import os
from pathlib import Path

# Add src directory to path for PyInstaller
if getattr(sys, 'frozen', False):
    # Running as compiled executable
    src_dir = Path(sys._MEIPASS)
else:
    # Running as script
    src_dir = Path(__file__).parent

sys.path.insert(0, str(src_dir))

# Patch python-docx template paths untuk PyInstaller
if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
    import docx.oxml
    import docx.parts.hdrftr
    import docx.parts.styles
    import docx.parts.settings
    import docx.parts.comments

    # Determine template directory based on platform
    if sys.platform == 'darwin':
        # macOS .app bundle - templates are in Frameworks
        base_path = os.path.join(sys._MEIPASS, '..', 'Frameworks')
    else:
        # Windows/Linux - templates are in _MEIPASS
        base_path = sys._MEIPASS

    template_dir = os.path.abspath(os.path.join(base_path, 'docx', 'templates'))

    # Patch all classes that use template files
    def make_template_loader(filename):
        def loader(_cls):
            path = os.path.join(template_dir, filename)
            with open(path, 'rb') as f:
                return docx.oxml.parse_xml(f.read())
        return classmethod(loader)

    docx.parts.hdrftr.Header.default = make_template_loader("default-header.xml")
    docx.parts.hdrftr.Footer.default = make_template_loader("default-footer.xml")
    docx.parts.styles.Styles.default = make_template_loader("default-styles.xml")
    docx.parts.settings.Settings.default = make_template_loader("default-settings.xml")
    docx.parts.comments.Comments.default = make_template_loader("default-comments.xml")

from gui.app import main

if __name__ == "__main__":
    main()
