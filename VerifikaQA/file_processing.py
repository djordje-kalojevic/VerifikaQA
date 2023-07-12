from pathlib import Path
from shutil import rmtree, copy


def prepare_files(files: list[Path]) -> Path:
    """Prepares a directory with files to be processed.
    If multiple files are provided, they are copied to a temporary sub-directory
    to avoid Verifika's built-in limitation of accepting only one path at a time.

    Args:
        - files list[Path]: File paths to be processed.

    Returns:
        - str: Path to the directory containing the files to be processed.
    If only one file was provided, the path to that file is returned instead."""

    if len(files) == 1:
        return files[0]

    files_dir = files[0].resolve().parent
    temp_dir = Path(files_dir, "temp_dir")

    delete_dir(temp_dir)
    Path.mkdir(temp_dir)

    for file in files:
        file_dst_location = Path(temp_dir, file.name)
        copy(file, file_dst_location)

    return temp_dir


def delete_dir(dir: Path | str) -> None:
    """Deletes the directory.

    Args:
        - dir (Path | str): Directory to be deleted."""

    try:
        rmtree(dir, ignore_errors=True)
    except OSError:
        pass
