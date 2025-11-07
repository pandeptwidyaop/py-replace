"""
Main GUI Application menggunakan CustomTkinter
"""
import customtkinter as ctk
from tkinter import filedialog, messagebox
from typing import Dict, List
import os
from pathlib import Path
import sys

# Add parent directory to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from utils.docx_handler import DocxHandler
from utils.placeholder import PlaceholderHandler
from utils.config_loader import ConfigLoader


class PlaceholderTable(ctk.CTkScrollableFrame):
    """Custom widget untuk menampilkan tabel placeholder dan input values"""

    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        self.entries: Dict[str, ctk.CTkEntry] = {}
        self.placeholders: List[str] = []

        # Header
        header_frame = ctk.CTkFrame(self)
        header_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
        header_frame.grid_columnconfigure(0, weight=1)
        header_frame.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(
            header_frame,
            text="Placeholder",
            font=ctk.CTkFont(size=14, weight="bold")
        ).grid(row=0, column=0, padx=10, pady=5, sticky="w")

        ctk.CTkLabel(
            header_frame,
            text="Value",
            font=ctk.CTkFont(size=14, weight="bold")
        ).grid(row=0, column=1, padx=10, pady=5, sticky="w")

    def set_placeholders(self, placeholders: List[str]):
        """
        Set placeholder list dan buat input fields

        Args:
            placeholders: List nama placeholder
        """
        # Clear existing entries
        for widget in self.winfo_children()[1:]:  # Skip header
            widget.destroy()
        self.entries.clear()

        self.placeholders = sorted(placeholders)

        # Create rows for each placeholder
        for idx, placeholder in enumerate(self.placeholders, start=1):
            row_frame = ctk.CTkFrame(self)
            row_frame.grid(row=idx, column=0, sticky="ew", padx=5, pady=2)
            row_frame.grid_columnconfigure(1, weight=1)

            # Placeholder label
            ctk.CTkLabel(
                row_frame,
                text=f"${{{placeholder}}}",
                font=ctk.CTkFont(size=12)
            ).grid(row=0, column=0, padx=10, pady=5, sticky="w")

            # Value entry
            entry = ctk.CTkEntry(
                row_frame,
                placeholder_text=f"Enter value for {placeholder}",
                width=300
            )
            entry.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

            self.entries[placeholder] = entry

    def get_values(self) -> Dict[str, str]:
        """
        Mendapatkan semua nilai dari input fields

        Returns:
            Dictionary mapping placeholder -> value
        """
        return {
            placeholder: entry.get()
            for placeholder, entry in self.entries.items()
        }

    def clear(self):
        """Clear semua input fields"""
        for entry in self.entries.values():
            entry.delete(0, 'end')

    def set_values(self, values: Dict[str, str]):
        """
        Set values ke input fields dari config

        Args:
            values: Dictionary mapping placeholder -> value
        """
        for placeholder, value in values.items():
            if placeholder in self.entries:
                self.entries[placeholder].delete(0, 'end')
                self.entries[placeholder].insert(0, value)


class DocxReplacerApp(ctk.CTk):
    """Main Application Window"""

    def __init__(self):
        super().__init__()

        # Configure window
        self.title("DOCX Placeholder Replacer")
        self.geometry("800x600")

        # Set theme
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # Variables
        self.docx_handler: DocxHandler = None
        self.current_file: str = None
        self.current_placeholders: set = set()

        # Setup UI
        self._setup_ui()

    def _setup_ui(self):
        """Setup UI components"""

        # Configure grid
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(1, weight=1)

        # Top frame - File operations
        top_frame = ctk.CTkFrame(self)
        top_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        top_frame.grid_columnconfigure(1, weight=1)

        # Load button
        self.load_button = ctk.CTkButton(
            top_frame,
            text="Load DOCX",
            command=self.load_docx,
            width=120
        )
        self.load_button.grid(row=0, column=0, padx=5, pady=10)

        # File label
        self.file_label = ctk.CTkLabel(
            top_frame,
            text="No file loaded",
            font=ctk.CTkFont(size=12)
        )
        self.file_label.grid(row=0, column=1, padx=10, pady=10, sticky="w")

        # Middle frame - Placeholder table
        middle_frame = ctk.CTkFrame(self)
        middle_frame.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)
        middle_frame.grid_columnconfigure(0, weight=1)
        middle_frame.grid_rowconfigure(1, weight=1)

        # Title and Load Config button row
        title_frame = ctk.CTkFrame(middle_frame)
        title_frame.grid(row=0, column=0, sticky="ew", padx=10, pady=10)
        title_frame.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(
            title_frame,
            text="Placeholders Found",
            font=ctk.CTkFont(size=16, weight="bold")
        ).grid(row=0, column=0, sticky="w")

        # Load Config button
        self.load_config_button = ctk.CTkButton(
            title_frame,
            text="Load Config (CSV/XLSX)",
            command=self.load_config,
            width=180,
            state="disabled",
            fg_color="green",
            hover_color="darkgreen"
        )
        self.load_config_button.grid(row=0, column=1, padx=5)

        # Export Template button
        self.export_template_button = ctk.CTkButton(
            title_frame,
            text="Export Template",
            command=self.export_template,
            width=150,
            state="disabled",
            fg_color="orange",
            hover_color="darkorange"
        )
        self.export_template_button.grid(row=0, column=2, padx=5)

        # Placeholder table
        self.placeholder_table = PlaceholderTable(middle_frame)
        self.placeholder_table.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)

        # Bottom frame - Action buttons
        bottom_frame = ctk.CTkFrame(self)
        bottom_frame.grid(row=2, column=0, sticky="ew", padx=10, pady=10)

        # Replace button
        self.replace_button = ctk.CTkButton(
            bottom_frame,
            text="Replace & Save",
            command=self.replace_and_save,
            width=150,
            state="disabled"
        )
        self.replace_button.pack(side="right", padx=5)

        # Clear button
        self.clear_button = ctk.CTkButton(
            bottom_frame,
            text="Clear Values",
            command=self.placeholder_table.clear,
            width=150,
            state="disabled",
            fg_color="gray"
        )
        self.clear_button.pack(side="right", padx=5)

    def load_docx(self):
        """Load DOCX file dan scan placeholders"""
        file_path = filedialog.askopenfilename(
            title="Select DOCX File",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )

        if not file_path:
            return

        try:
            # Load document
            self.docx_handler = DocxHandler(file_path)
            self.current_file = file_path

            # Update file label
            filename = os.path.basename(file_path)
            self.file_label.configure(text=f"Loaded: {filename}")

            # Find placeholders
            placeholders = self.docx_handler.find_all_placeholders()

            if not placeholders:
                messagebox.showinfo(
                    "No Placeholders",
                    "No placeholders found in the document.\n\n"
                    "Placeholders should be in format: ${placeholder_name}"
                )
                return

            # Store placeholders
            self.current_placeholders = placeholders

            # Display placeholders in table
            self.placeholder_table.set_placeholders(list(placeholders))

            # Enable buttons
            self.replace_button.configure(state="normal")
            self.clear_button.configure(state="normal")
            self.load_config_button.configure(state="normal")
            self.export_template_button.configure(state="normal")

            messagebox.showinfo(
                "Success",
                f"Found {len(placeholders)} placeholder(s) in the document."
            )

        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Failed to load document:\n{str(e)}"
            )

    def replace_and_save(self):
        """Replace placeholders dan save ke file baru"""
        if not self.docx_handler:
            return

        # Get values from table
        replacements = self.placeholder_table.get_values()

        # Check for empty values
        empty_placeholders = [
            key for key, value in replacements.items() if not value.strip()
        ]

        if empty_placeholders:
            response = messagebox.askyesno(
                "Empty Values",
                f"The following placeholders have empty values:\n\n"
                f"{', '.join(empty_placeholders)}\n\n"
                f"Do you want to continue?"
            )
            if not response:
                return

        # Ask for output file
        output_path = filedialog.asksaveasfilename(
            title="Save Modified Document",
            defaultextension=".docx",
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")],
            initialfile=f"{Path(self.current_file).stem}_modified.docx"
        )

        if not output_path:
            return

        try:
            # Perform replacement
            self.docx_handler.replace_placeholders(replacements)

            # Save document
            self.docx_handler.save(output_path)

            messagebox.showinfo(
                "Success",
                f"Document saved successfully!\n\n{output_path}"
            )

            # Reload the original document
            self.docx_handler.load(self.current_file)

        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Failed to save document:\n{str(e)}"
            )

    def load_config(self):
        """Load config dari CSV atau XLSX dan auto-fill values"""
        if not self.current_placeholders:
            messagebox.showwarning(
                "No Document",
                "Please load a DOCX file first."
            )
            return

        # Open file dialog
        file_path = filedialog.askopenfilename(
            title="Select Config File",
            filetypes=[
                ("CSV Files", "*.csv"),
                ("Excel Files", "*.xlsx *.xls"),
                ("All Files", "*.*")
            ]
        )

        if not file_path:
            return

        try:
            # Load config
            config, error = ConfigLoader.load_config(file_path)

            if error:
                messagebox.showerror("Error", error)
                return

            if not config:
                messagebox.showwarning(
                    "Empty Config",
                    "No valid placeholder-value pairs found in the config file."
                )
                return

            # Validate config
            is_perfect, message, missing, extra = ConfigLoader.validate_config(
                config, self.current_placeholders
            )

            # Auto-fill values
            self.placeholder_table.set_values(config)

            # Show validation result
            if is_perfect:
                messagebox.showinfo(
                    "Config Loaded",
                    f"Successfully loaded {len(config)} value(s)!\n\n{message}"
                )
            else:
                messagebox.showwarning(
                    "Config Loaded with Warnings",
                    f"Loaded {len(config)} value(s).\n\n{message}"
                )

        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Failed to load config:\n{str(e)}"
            )

    def export_template(self):
        """Export template config file"""
        if not self.current_placeholders:
            messagebox.showwarning(
                "No Document",
                "Please load a DOCX file first."
            )
            return

        # Ask for output file
        file_path = filedialog.asksaveasfilename(
            title="Save Config Template",
            defaultextension=".csv",
            filetypes=[
                ("CSV Files", "*.csv"),
                ("Excel Files", "*.xlsx"),
                ("All Files", "*.*")
            ],
            initialfile="config_template.csv"
        )

        if not file_path:
            return

        try:
            success, error = ConfigLoader.save_template(
                file_path,
                list(self.current_placeholders)
            )

            if success:
                messagebox.showinfo(
                    "Success",
                    f"Template exported successfully!\n\n{file_path}\n\n"
                    f"You can now fill in the values and load it back."
                )
            else:
                messagebox.showerror("Error", error)

        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Failed to export template:\n{str(e)}"
            )


def main():
    """Main entry point"""
    app = DocxReplacerApp()
    app.mainloop()


if __name__ == "__main__":
    main()
