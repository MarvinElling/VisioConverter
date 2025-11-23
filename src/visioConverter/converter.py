import os
from pathlib import Path

import win32com.client as win32


def vsdx_to_svg_with_visio(src_path, out_dir=None, visible=False):
    if not os.path.isfile(src_path):
        raise FileNotFoundError(f"The source file '{src_path}' does not exist.")
    if out_dir is None:
        out_dir = os.path.dirname(src_path)
    os.makedirs(out_dir, exist_ok=True)

    visio = win32.Dispatch('Visio.Application')
    visio.Visible = visible
    visio.AlertResponse = 7  # suppress promts

    doc = visio.Documents.Open(src_path)
    try:
        base_name = os.path.splitext(os.path.basename(src_path))[0]
        out_files = []
        for i, page in enumerate(doc.Pages, start=1):
            out_file = os.path.join(out_dir, f'{base_name}_page{i}.svg')
            page.Export(out_file)
            out_files.append(out_file)
        return out_files
    finally:
        doc.Close()
        visio.Quit()

def process_folder(src_folder, out_folder=None, visible=False):
    src = Path(src_folder)
    tgt = Path(out_folder)
    tgt.mkdir(parents=True, exist_ok=True)


    for item in src.iterdir():
        if not item.is_file():
            continue

        ext = item.suffix.lower()
        if ext == '.vsdx':
            print(f'Converting VSDX: {item} to SVG')
            try:
                vsdx_to_svg_with_visio(str(item), str(tgt), visible)
            except Exception as e:
                print(f'Failed to convert {item}: {e}')
        else:
            print(f'Skipping unsupported file: {item}')

if __name__ == '__main__':
    pass
