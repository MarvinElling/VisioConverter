"""Test Converter."""

from visioconverter import Converter  # pylint: disable=E0401

src_path = 'test/src_dir/'
src_file = 'test/src_dir/example1.vsdx'
out_format = 'svg'
out_path = 'test/out_dir'

Converter.vsdx2other(
    src_path=src_file,
    out_format=out_format,
    out_dir=out_path,
)

Converter.folder_vsdx2other(
    src_folder='test/src_dir',
    out_format=out_format,
    out_folder=out_path,
)
