"""Test Converter."""

from visioconverter import Converter  # pylint: disable=E0401

src_path = 'test/src_dir/example1.vsdx'
out_path = 'test/out_dir/'

Converter.vsdx2svg(
    src_path=src_path,
    out_dir=out_path,
)

Converter.folder_vsdx2svg(
    src_folder='test/src_dir',
    out_folder=out_path,
)
