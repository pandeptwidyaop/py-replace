# DOCX Placeholder Replacement

Aplikasi desktop sederhana untuk mengganti placeholder dalam dokumen DOCX.

## Fitur

- Load dokumen DOCX
- **Dual Placeholder Support**:
  - Text placeholder: `${nama_placeholder}` - untuk text replacement
  - Image placeholder: `@{nama_placeholder}` - untuk image replacement
- Deteksi otomatis kedua tipe placeholder
- **Format Preservation** - Mempertahankan formatting text asli (bold, italic, underline, font, size, color)
- Tampilan tabel interaktif dengan file browser untuk images
- **Image dari URL atau Local** - Support image dari path lokal atau URL
- **Load config dari CSV/XLSX** - Import nilai replacement secara batch
- **Export template config** - Generate template CSV/XLSX dari placeholder yang terdeteksi
- Validasi config terhadap placeholder yang ditemukan
- Export dokumen DOCX yang sudah dimodifikasi

## Download & Install

### Untuk End Users (Tanpa Python)

Download executable yang sudah di-build untuk platform Anda dari [**Releases**](https://github.com/YOUR_USERNAME/py-replace/releases/latest):

- **Windows**: `DOCX-Replacer-Windows.zip`
  - Extract dan jalankan `DOCX-Replacer.exe`

- **macOS**: `DOCX-Replacer-macOS.zip`
  - Extract dan jalankan `DOCX-Replacer.app`
  - Support Apple Silicon (M1/M2/M3) dan Intel

- **Linux**: `DOCX-Replacer-Linux.tar.gz`
  - Extract: `tar -xzf DOCX-Replacer-Linux.tar.gz`
  - Jalankan: `./DOCX-Replacer`

**Tidak perlu install Python atau dependencies lainnya!**

> **Note**: Builds otomatis di-generate via GitHub Actions setiap kali ada release tag baru.

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
3. Placeholder akan otomatis terdeteksi dan ditampilkan dalam tabel:
   - Text placeholders (`${}`) - putih
   - Image placeholders (`@{}`) - orange, dengan tombol Browse
4. Isi nilai untuk setiap placeholder:
   - Text: Ketik langsung nilai pengganti
   - Image: Ketik path/URL atau klik "Browse" untuk pilih file
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
   logo,/path/to/logo.png
   foto,https://example.com/photo.jpg
   ```

   Note: Untuk image placeholder, value bisa berupa:
   - Path lokal: `/path/to/image.png`
   - URL: `https://example.com/image.jpg`

4. Klik "Load Config (CSV/XLSX)" untuk auto-fill values
5. Review dan edit jika perlu
6. Klik "Replace & Save"

### Format Preservation

Aplikasi ini **mempertahankan semua formatting text asli** saat melakukan replacement:
- **Bold**, *Italic*, <u>Underline</u>
- Font family (Arial, Times New Roman, dll)
- Font size (10pt, 12pt, 14pt, dll)
- Font color

**Contoh:**
Jika dalam dokumen template Anda memiliki placeholder `${nama}` dengan format **Bold 14pt Red**,
maka saat diganti dengan "John Doe", text "John Doe" akan tetap **Bold 14pt Red**.

**Cara kerja:**
- Aplikasi membaca formatting dari setiap run text
- Saat replacement, formatting asli di-copy ke text baru
- Support untuk placeholder yang span multiple runs dengan formatting berbeda

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

## Automated Releases (GitHub Actions)

Project ini menggunakan **Semantic Release** dengan GitHub Actions untuk automated versioning dan multi-platform builds.

### Cara Kerja Workflow

Workflow berjalan otomatis pada setiap push ke branch `main`:

1. **Version Check (Semantic Release Dry Run)**
   - Analyze commit messages menggunakan [Conventional Commits](https://www.conventionalcommits.org/)
   - Tentukan apakah perlu release baru dan versi berapa
   - Skip build jika tidak ada perubahan yang memerlukan release

2. **Build (Conditional)**
   - Hanya run jika semantic release detect ada release baru
   - Build parallel untuk Windows, macOS, dan Linux
   - Package artifacts (zip/tar.gz)

3. **Release**
   - Create GitHub Release dengan semantic versioning
   - Upload semua platform binaries
   - Generate CHANGELOG.md otomatis
   - Create git tag dan push

### Commit Message Format

Gunakan **Conventional Commits** untuk trigger automated releases:

```bash
# Minor version bump (1.0.0 -> 1.1.0)
feat: add new image placeholder feature
feat(ui): improve placeholder table UX

# Patch version bump (1.0.0 -> 1.0.1)
fix: resolve formatting preservation issue
fix(build): update PyInstaller configuration
docs: update README with new instructions

# Major version bump (1.0.0 -> 2.0.0)
feat!: change placeholder format to use square brackets
BREAKING CHANGE: placeholder format changed from ${} to []

# No release
chore: update dependencies
chore(ci): improve workflow performance
```

### Release Rules

| Commit Type | Release Type | Example |
|-------------|--------------|---------|
| `feat:` | Minor | `feat: add CSV export` → 1.0.0 → 1.1.0 |
| `fix:` | Patch | `fix: resolve bug` → 1.0.0 → 1.0.1 |
| `perf:` | Patch | `perf: optimize rendering` → 1.0.0 → 1.0.1 |
| `BREAKING CHANGE:` | Major | `feat!: new API` → 1.0.0 → 2.0.0 |
| `chore:` | No release | `chore: update deps` → (skip) |

### Manual Trigger

Trigger workflow manual tanpa push:
1. Go to **Actions** tab di GitHub
2. Pilih "Build and Release" workflow
3. Click "Run workflow" → Select branch `main`

### Download Releases

User bisa download dari [Releases page](https://github.com/YOUR_USERNAME/py-replace/releases)

### Configuration Files

- **Workflow**: [.github/workflows/build-release.yml](.github/workflows/build-release.yml)
- **Semantic Release Config**: [.releaserc.json](.releaserc.json)

**Workflow Features**:
- ✅ Semantic versioning otomatis
- ✅ Conventional commits analysis
- ✅ Dry run check sebelum build
- ✅ Multi-platform build matrix (Windows, macOS, Linux)
- ✅ Conditional build (hanya jika ada release)
- ✅ Automated CHANGELOG generation
- ✅ GitHub Release dengan semua binaries

## Struktur Project

```
py-replace/
├── .github/
│   └── workflows/
│       └── build-release.yml    # GitHub Actions workflow
├── src/
│   ├── main.py                  # Entry point aplikasi
│   ├── gui/
│   │   ├── __init__.py
│   │   └── app.py               # Main GUI window
│   └── utils/
│       ├── __init__.py
│       ├── docx_handler.py      # Load & save DOCX
│       ├── placeholder.py       # Deteksi & replace placeholder
│       ├── config_loader.py     # Load config dari CSV/XLSX
│       └── image_handler.py     # Handle image operations & downloads
├── tests/                       # Unit tests
├── build.sh                     # Build script untuk macOS/Linux
├── build.bat                    # Build script untuk Windows
├── docx-replacer.spec           # PyInstaller configuration
├── requirements.txt             # Python dependencies
├── .releaserc.json              # Semantic Release configuration
├── CHANGELOG.md                 # Auto-generated changelog
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
