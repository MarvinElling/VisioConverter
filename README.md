# Visio VSDX to SVG Converter

*A Python utility for batch converting Microsoft Visio VSDX files to SVG using the Visio COM interface*

---

## Features

- Convert a single `.vsdx` file to individual SVG files (one per Visio page)
- Batch-convert all `.vsdx` files in a directory
- Leverages the official Microsoft Visio COM interface for reliable conversion

---

## Requirements

- **Windows** operating system
- **Microsoft Visio** installed (required for the COM Automation)
- Python 3.7+
- `pywin32` package

---

## Installation

1. **Install pywin32** (if you haven't):

   ```sh
   pip install pywin32
Clone or download this repository.
Usage
1. Import and Use in Python
from converter import Converter

# Convert a single .vsdx file to SVG
```python
svg_files = Converter.vsdx2svg(
    src_path="diagram.vsdx",
    out_dir="output_svgs",    # optional, defaults to same as source
    visible=False             # optional, show Visio UI?
)
```

# Batch-convert all .vsdx files in a folder
```python
Converter.folder_vsdx2svg(
    src_folder="vsdx_files",
    out_folder="output_svgs",  # optional, defaults to source folder
    visible=False
)
```
2. From Command Line
No CLI interface is included by default. See Python sample above or integrate as needed.

# Reference
```
Converter.vsdx2svg(src_path, out_dir=None, visible=False) -> list[str]
    src_path: Path to a .vsdx file.
    out_dir: Output directory (default: same as source file).
    visible: Whether to show the Visio UI (default: False).
    Returns list of SVG file paths created.

Converter.folder_vsdx2svg(src_folder, out_folder=None, visible=False)
    src_folder: Folder containing .vsdx files.
    out_folder: Where to save SVGs (default: same as each source file).
    visible: Whether to show the Visio UI.
```

## Notes
Requires Microsoft Visio to be installed and licensed on your system.
Only Windows is supported (due to COM automation).
Each page in a Visio document is exported as a separate SVG (<base>_pageX.svg).
Error messages are printed to stdout.

