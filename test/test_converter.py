from visioconverter import Converter

src_path = "src_dir/example1.vsdx"

Converter.vsdx2svg(
    src_path=src_path,
    out_dir="out_dir",
    visible=False,
)

Converter.folder_vsdx2svg(
    src_folder="src_dir",
    out_folder="out_dir",
    visible=False,
)
