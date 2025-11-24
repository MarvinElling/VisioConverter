"""Convert Visio VSDX files using Microsoft Visio COM interface."""

from pathlib import Path

import pythoncom
import win32com.client as win32


class UnsupportedFormatError(ValueError):
    """Raised when an unsupported export format is specified."""

    def __init__(self, bad_format: str, allowed_formats: set[str]) -> None:
        """Initialize the UnsupportedFormatError.

        Args:
            bad_format (str): The unsupported format requested.
            allowed_formats (set[str]): The set of supported formats.
        """
        allowed = ', '.join(allowed_formats)
        msg = f"Unsupported format '{bad_format}'. Supported: {allowed}."
        super().__init__(msg)


class Converter:
    """Converter class to handle VSDX to other file conversion."""

    def __init__(self) -> None:
        """Initialize the Converter class."""

    @staticmethod
    def vsdx2other(
        src_path: str,
        out_format: str,
        out_dir: str | None = None,
        *,
        visible: bool = False,
    ) -> list[str]:
        """Convert a Visio VSDX file to other format.

        Args:
            src_path (str): Path to the source .vsdx file.
            out_format (str): The format to convert to, supported:
                "html", "png", "jpg", "gif", "tif", "bmp", "emf", "wmf", "svg".
            out_dir (Optional[str], optional):
                Directory to store output files.
                If None, files are saved in the same directory as the source
                    file. Defaults to None.
            visible (bool, optional): If True, the Visio GUI will be visible
                during conversion. Defaults to False.

        Returns:
            list[str]: A list of filepaths to the generated files.

        Raises:
            FileNotFoundError: If the source file does not exist.
            Exception: For other errors during conversion.
        """
        allowed_formats = {
            'html',
            'png',
            'jpg',
            'gif',
            'tif',
            'bmp',
            'emf',
            'wmf',
            'svg',
        }
        if out_format not in allowed_formats:
            raise UnsupportedFormatError(out_format, allowed_formats)
        src_path = Path(src_path)
        if not src_path.is_absolute():
            src_path = (Path.cwd() / src_path).resolve()

        if out_dir is not None:
            out_dir = Path(out_dir)
            if not out_dir.is_absolute():
                out_dir = (Path.cwd() / out_dir).resolve()
        else:
            out_dir = src_path.parent

        if not src_path.is_file():
            err_text = f"The source file '{src_path}' does not exist."
            raise FileNotFoundError(err_text)
        out_dir.mkdir(parents=True, exist_ok=True)

        visio = win32.Dispatch('Visio.Application')
        visio.Visible = visible
        visio.AlertResponse = 7  # suppress promts

        doc = visio.Documents.Open(src_path)
        doc = visio.Documents.Open(src_path)
        try:
            base_name = src_path.stem
            out_files = []
            for i, page in enumerate(doc.Pages, start=1):
                out_file = out_dir / f'{base_name}_page{i}.{out_format}'
                page.Export(out_file)

                out_files.append(out_file)
            return out_files
        finally:
            doc.Close()
            visio.Quit()

    @staticmethod
    def folder_vsdx2other(
        src_folder: str,
        out_format: str,
        out_folder: str | None = None,
        *,
        visible: bool = False,
    ) -> None:
        """Convert all Visio VSDX files in a folder to other format.

        Args:
            src_folder (str): Path to the folder containing .vsdx files.
            out_format (str): The format to convert to, supported:
                "html", "png", "jpg", "gif", "tif", "bmp", "emf", "wmf", "svg".
            out_folder (Optional[str], optional): Path to the output directory.
                If None, output files will be saved in the same directory as
                    the source files. Defaults to None.
            visible (bool, optional): If True, the Visio GUI will be visible
                during conversion. Defaults to False.

        Returns:
            None

        Raises:
            FileNotFoundError: If the source folder does not exist.
            Exception: For any errors during conversion (printed to standard
                output).
        """
        src = Path(src_folder)
        tgt = Path(out_folder)
        tgt.mkdir(parents=True, exist_ok=True)

        for item in src.iterdir():
            if not item.is_file():
                continue

            ext = item.suffix.lower()
            if ext == '.vsdx':
                print(f'Converting VSDX: {item} to {out_format} format')
                try:
                    Converter.vsdx2other(
                        str(item),
                        str(out_format),
                        str(tgt),
                        visible=visible,
                    )
                except FileNotFoundError as e:
                    print(f'File not found: {item}: {e}')
                except pythoncom.com_error as e:  # pylint: disable=E1101
                    print(e)
                    print(vars(e))
                    print(e.args)

            else:
                print(f'Skipping unsupported file: {item}')


if __name__ == '__main__':
    pass
