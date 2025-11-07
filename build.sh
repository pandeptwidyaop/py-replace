#!/bin/bash
# Build script untuk membuat executable

echo "=========================================="
echo "Building DOCX Placeholder Replacer"
echo "=========================================="

# Activate virtual environment if exists
if [ -d "venv" ]; then
    source venv/bin/activate
fi

# Clean previous builds
echo "Cleaning previous builds..."
rm -rf build dist

# Build using PyInstaller
echo "Building executable..."
pyinstaller docx-replacer.spec

# Check if build succeeded
if [ $? -eq 0 ]; then
    echo ""
    echo "=========================================="
    echo "Build completed successfully!"
    echo "=========================================="
    echo ""
    if [ "$(uname)" == "Darwin" ]; then
        echo "macOS Application: dist/DOCX-Replacer.app"
        echo ""
        echo "To run:"
        echo "  open dist/DOCX-Replacer.app"
    else
        echo "Executable: dist/DOCX-Replacer"
        echo ""
        echo "To run:"
        echo "  ./dist/DOCX-Replacer"
    fi
    echo ""
    echo "You can distribute the 'dist' folder to users."
    echo "They don't need Python installed!"
else
    echo ""
    echo "=========================================="
    echo "Build failed! Check errors above."
    echo "=========================================="
    exit 1
fi
