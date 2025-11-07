# DOCX Placeholder Replacement

Aplikasi desktop sederhana untuk mengganti placeholder dalam dokumen DOCX.

## Fitur

- Load dokumen DOCX
- Deteksi otomatis placeholder dengan format `${nama_placeholder}`
- Tampilan tabel interaktif untuk input nilai replacement
- **Load config dari CSV/XLSX** - Import nilai replacement secara batch
- **Export template config** - Generate template CSV/XLSX dari placeholder yang terdeteksi
- Validasi config terhadap placeholder yang ditemukan
- Export dokumen DOCX yang sudah dimodifikasi

## Download & Install

### Untuk End Users (Tanpa Python)

Download executable yang sudah di-build untuk platform Anda:
- **Windows**: `DOCX-Replacer.exe`
- **macOS**: `DOCX-Replacer.app`
- **Linux**: `DOCX-Replacer`

Tidak perlu install Python atau dependencies lainnya!

### Untuk Developers

Requirements:
- Python 3.8+
- CustomTkinter
- python-docx
- pandas
- openpyxl
- pyinstaller (untuk build)

## Instalasi (untuk Developer)

1. Clone repository ini
2. Buat virtual environment:
```bash
python3 -m venv venv
```

3. Aktifkan virtual environment:
```bash
# MacOS/Linux
source venv/bin/activate

# Windows
venv\Scripts\activate
```

4. Install dependencies:
```bash
pip install -r requirements.txt
```

## Cara Menggunakan

### Basic Usage

1. Jalankan aplikasi:
```bash
python src/main.py
```

2. Klik "Load DOCX" untuk memilih file dokumen
3. Placeholder akan otomatis terdeteksi dan ditampilkan dalam tabel
4. Isi nilai untuk setiap placeholder secara manual
5. Klik "Replace & Save" untuk menyimpan dokumen baru

### Using Config File (CSV/XLSX)

1. Load dokumen DOCX seperti biasa
2. **Opsi 1: Export Template**
   - Klik "Export Template" untuk generate template config
   - Buka file CSV/XLSX yang diexport
   - Isi kolom `value` untuk setiap placeholder
   - Save file config

3. **Opsi 2: Buat Manual**
   - Buat file CSV/XLSX dengan 2 kolom:
     - Kolom 1: `placeholder` (nama placeholder atau `${placeholder}`)
     - Kolom 2: `value` (nilai pengganti)

   Contoh CSV:
   ```csv
   placeholder,value
   nama,John Doe
   tanggal,2025-11-07
   alamat,Jakarta
   ```

4. Klik "Load Config (CSV/XLSX)" untuk auto-fill values
5. Review dan edit jika perlu
6. Klik "Replace & Save"

## Build Executable (untuk Developer)

Untuk membuat executable yang bisa didistribusikan tanpa Python:

### macOS / Linux
```bash
./build.sh
```

### Windows
```cmd
build.bat
```

Output akan berada di folder `dist/`:
- **macOS**: `dist/DOCX-Replacer.app`
- **Windows**: `dist/DOCX-Replacer.exe`
- **Linux**: `dist/DOCX-Replacer`

### Build Manual dengan PyInstaller
```bash
pyinstaller docx-replacer.spec
```

### Cross-Platform Build Notes

- **macOS**: Builds untuk macOS hanya bisa dilakukan di macOS
- **Windows**: Builds untuk Windows hanya bisa dilakukan di Windows
- **Linux**: Builds untuk Linux hanya bisa dilakukan di Linux

Untuk mendistribusikan ke multiple platform, Anda perlu build di masing-masing OS atau gunakan CI/CD seperti GitHub Actions.

## Struktur Project

```
py-replace/
├── src/
│   ├── main.py              # Entry point aplikasi
│   ├── gui/
│   │   ├── __init__.py
│   │   └── app.py           # Main GUI window
│   ├── utils/
│   │   ├── __init__.py
│   │   ├── docx_handler.py  # Load & save DOCX
│   │   ├── placeholder.py   # Deteksi & replace placeholder
│   │   └── config_loader.py # Load config dari CSV/XLSX
├── tests/                   # Unit tests
├── build.sh                 # Build script untuk macOS/Linux
├── build.bat                # Build script untuk Windows
├── docx-replacer.spec       # PyInstaller configuration
├── requirements.txt
├── README.md
└── .gitignore
```

## Troubleshooting

### Build Issues

**Error: ModuleNotFoundError saat running executable**
- Pastikan semua dependencies ada di `hiddenimports` di file `.spec`
- Re-run build script

**Executable terlalu besar**
- Normal untuk PyInstaller (~50-100MB)
- Bisa dikompres dengan UPX (already enabled in spec)

**Tkinter errors di Linux**
- Install: `sudo apt-get install python3-tk`

## License

MIT
