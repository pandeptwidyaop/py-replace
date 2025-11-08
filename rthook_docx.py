"""
PyInstaller runtime hook untuk python-docx
Memperbaiki path template di bundled environment
"""
import os
import sys

# Debug: Print bahwa runtime hook dijalankan
print("DEBUG: Runtime hook rthook_docx.py is running...")


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
            # Use sys._MEIPASS directly for all platforms
            base_path = sys._MEIPASS
            template_dir = os.path.join(base_path, 'docx', 'templates')

            # Debug logging - print paths for troubleshooting
            print(f"[DEBUG] Running in frozen mode (PyInstaller)")
            print(f"[DEBUG] Platform: {sys.platform}")
            print(f"[DEBUG] sys._MEIPASS: {sys._MEIPASS}")
            print(f"[DEBUG] Template directory: {template_dir}")
            print(f"[DEBUG] Template directory exists: {os.path.exists(template_dir)}")

            # Check if templates exist
            if os.path.exists(template_dir):
                templates = os.listdir(template_dir)
                print(f"[DEBUG] Available templates: {templates}")
            else:
                # If templates not in expected location, try to find them
                print(f"[DEBUG] Template directory not found, searching...")
                for root, dirs, files in os.walk(base_path):
                    if 'templates' in dirs:
                        alt_template_dir = os.path.join(root, 'templates')
                        if any('default-' in f for f in os.listdir(alt_template_dir)):
                            print(f"[DEBUG] Found templates at: {alt_template_dir}")
                            template_dir = alt_template_dir
                            break
        else:
            # Running in normal Python - tidak perlu patch
            return

        # Validate template directory exists before patching
        if not os.path.exists(template_dir):
            raise FileNotFoundError(
                f"Template directory not found: {template_dir}\n"
                f"Please report this issue with the debug information above."
            )

        # Patch Header class
        original_header_default = docx.parts.hdrftr.Header.default
        def patched_header_default(cls):
            path = os.path.join(template_dir, "default-header.xml")
            if not os.path.exists(path):
                raise FileNotFoundError(f"Template file not found: {path}")
            with open(path, 'rb') as f:
                return docx.oxml.parse_xml(f.read())
        docx.parts.hdrftr.Header.default = classmethod(patched_header_default)

        # Patch Footer class
        original_footer_default = docx.parts.hdrftr.Footer.default
        def patched_footer_default(cls):
            path = os.path.join(template_dir, "default-footer.xml")
            if not os.path.exists(path):
                raise FileNotFoundError(f"Template file not found: {path}")
            with open(path, 'rb') as f:
                return docx.oxml.parse_xml(f.read())
        docx.parts.hdrftr.Footer.default = classmethod(patched_footer_default)

        # Patch Styles class
        original_styles_default = docx.parts.styles.Styles.default
        def patched_styles_default(cls):
            path = os.path.join(template_dir, "default-styles.xml")
            if not os.path.exists(path):
                raise FileNotFoundError(f"Template file not found: {path}")
            with open(path, 'rb') as f:
                return docx.oxml.parse_xml(f.read())
        docx.parts.styles.Styles.default = classmethod(patched_styles_default)

        # Patch Settings class
        original_settings_default = docx.parts.settings.Settings.default
        def patched_settings_default(cls):
            path = os.path.join(template_dir, "default-settings.xml")
            if not os.path.exists(path):
                raise FileNotFoundError(f"Template file not found: {path}")
            with open(path, 'rb') as f:
                return docx.oxml.parse_xml(f.read())
        docx.parts.settings.Settings.default = classmethod(patched_settings_default)

        # Patch Comments class
        original_comments_default = docx.parts.comments.Comments.default
        def patched_comments_default(cls):
            path = os.path.join(template_dir, "default-comments.xml")
            if not os.path.exists(path):
                raise FileNotFoundError(f"Template file not found: {path}")
            with open(path, 'rb') as f:
                return docx.oxml.parse_xml(f.read())
        docx.parts.comments.Comments.default = classmethod(patched_comments_default)

        print(f"[DEBUG] Successfully patched all docx template paths")

    except (ImportError, AttributeError) as e:
        # Jika terjadi error, print untuk debugging
        print(f"[ERROR] Failed to patch docx template paths: {e}")
        import traceback
        traceback.print_exc()
    except FileNotFoundError as e:
        print(f"[ERROR] Template files missing: {e}")
        import traceback
        traceback.print_exc()
    except Exception as e:
        print(f"[ERROR] Unexpected error in patch_docx_template_paths: {e}")
        import traceback
        traceback.print_exc()


# Apply patch saat module di-load
patch_docx_template_paths()