"""
PyInstaller runtime hook untuk python-docx
Memperbaiki path template di bundled environment
"""
import os
import sys


def patch_docx_template_paths():
    """Patch python-docx template path resolution untuk PyInstaller"""
    try:
        # Import modules yang menggunakan template paths
        import docx.oxml
        import docx.parts.hdrftr
        import docx.parts.styles
        import docx.parts.settings
        import docx.parts.comments

        # Get base path for templates
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            # Running in PyInstaller bundle
            if sys.platform == 'darwin':
                # macOS .app bundle
                base_path = os.path.join(sys._MEIPASS, '..', 'Frameworks')
            else:
                # Windows/Linux
                base_path = sys._MEIPASS
            template_dir = os.path.abspath(os.path.join(base_path, 'docx', 'templates'))
        else:
            # Running in normal Python - tidak perlu patch
            return

        # Patch Header class
        original_header_default = docx.parts.hdrftr.Header.default
        def patched_header_default(cls):
            path = os.path.join(template_dir, "default-header.xml")
            with open(path, 'rb') as f:
                return docx.oxml.parse_xml(f.read())
        docx.parts.hdrftr.Header.default = classmethod(patched_header_default)

        # Patch Footer class
        original_footer_default = docx.parts.hdrftr.Footer.default
        def patched_footer_default(cls):
            path = os.path.join(template_dir, "default-footer.xml")
            with open(path, 'rb') as f:
                return docx.oxml.parse_xml(f.read())
        docx.parts.hdrftr.Footer.default = classmethod(patched_footer_default)

        # Patch Styles class
        original_styles_default = docx.parts.styles.Styles.default
        def patched_styles_default(cls):
            path = os.path.join(template_dir, "default-styles.xml")
            with open(path, 'rb') as f:
                return docx.oxml.parse_xml(f.read())
        docx.parts.styles.Styles.default = classmethod(patched_styles_default)

        # Patch Settings class
        original_settings_default = docx.parts.settings.Settings.default
        def patched_settings_default(cls):
            path = os.path.join(template_dir, "default-settings.xml")
            with open(path, 'rb') as f:
                return docx.oxml.parse_xml(f.read())
        docx.parts.settings.Settings.default = classmethod(patched_settings_default)

        # Patch Comments class
        original_comments_default = docx.parts.comments.Comments.default
        def patched_comments_default(cls):
            path = os.path.join(template_dir, "default-comments.xml")
            with open(path, 'rb') as f:
                return docx.oxml.parse_xml(f.read())
        docx.parts.comments.Comments.default = classmethod(patched_comments_default)

    except (ImportError, AttributeError) as e:
        # Jika terjadi error, print untuk debugging
        print(f"Warning: Failed to patch docx template paths: {e}")


# Apply patch saat module di-load
patch_docx_template_paths()