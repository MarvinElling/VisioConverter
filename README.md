[![PyPI Downloads](https://static.pepy.tech/personalized-badge/visioconverter?period=total&units=INTERNATIONAL_SYSTEM&left_color=BLACK&right_color=GREEN&left_text=downloads)](https://pepy.tech/projects/visioconverter)
# Visio VSDX to File Converter

*A Python utility for batch converting Microsoft Visio VSDX files to other using the Visio COM interface*

*Possible file formats: "html", "png", "jpg", "gif", "tif", "bmp", "emf", "wmf", "svg"*

---

## Features

- Convert a single `.vsdx` file to individual files (one per Visio page)
- Batch-convert all `.vsdx` files in a directory
- Leverages the official Microsoft Visio COM interface for reliable conversion

---

## Requirements

- **Windows** operating system
- **Microsoft Visio** installed (required for the COM Automation)
- Python 3.7+


---

## Installation

```sh
pip install visioconverter
```

Clone or download this repository.
Usage
1. Import and Use in Python
from converter import Converter

# Convert a single .vsdx file to other format
```python
other_files = Converter.vsdx2other(
    src_path="diagram.vsdx",
    out_format="png",         # Possible: "html", "png", "jpg", "gif", "tif", "bmp", "emf", "wmf", "svg"
    out_dir="output_svgs",    # optional, defaults to same as source
    visible=False             # optional, show Visio UI?
)
```

# Batch-convert all .vsdx files in a folder
```python
Converter.folder_vsdx2other(
    src_folder="vsdx_files",
    out_format="png",          # Possible: "html", "png", "jpg", "gif", "tif", "bmp", "emf", "wmf", "svg"
    out_folder="output_files",  # optional, defaults to source folder
    visible=False   
)
```
2. From Command Line
No CLI interface is included by default. See Python sample above or integrate as needed.

# Reference
```
Converter.vsdx2other(src_path, out_dir=None, visible=False) -> list[str]
    src_path: Path to a .vsdx file.
    out_dir: Output directory (default: same as source file).
    visible: Whether to show the Visio UI (default: False).
    Returns list of file paths created.

Converter.folder_vsdx2other(src_folder, out_folder=None, visible=False)
    src_folder: Folder containing .vsdx files.
    out_folder: Where to save files (default: same as each source file).
    visible: Whether to show the Visio UI.
```

## Notes
Requires Microsoft Visio to be installed and licensed on your system.
Only Windows is supported (due to COM automation).
Each page in a Visio document is exported as a separate SVG (<base>_pageX.svg).
Error messages are printed to stdout.

