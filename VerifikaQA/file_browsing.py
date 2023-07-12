"""This module provides browsing utility for necessary files.
This includes files to be checked,
Verifika profile to be used, as well as the Verifika executable itself."""

from pathlib import Path
from typing import Optional
from psutil import Process, NoSuchProcess, process_iter
from PyQt6.QtWidgets import QFileDialog, QMessageBox
from verifika_config import _ConfigFile, update_config


def browse_files() -> tuple[list[Path], Path]:
    """Prompts user to select files for checking.

    Returns:
        - files (list(Path): list of selected files
        - files_dir (Path): directory fo selected files."""

    files, _ = QFileDialog.getOpenFileNames(
        caption="Choose a file", filter="Any .xliff file (*.*xliff)")

    if not files:
        raise SystemExit

    file_paths = [Path(file) for file in files]
    files_dir = file_paths[0].resolve().parent

    return file_paths, files_dir


def browse_verifika_profile(config_file: _ConfigFile) -> Path:
    """Prompts the user to browse for a Verifika profile file,
    if its location is not specified in the config file.
    Updates the config file with the location of the Verifika profiles directory.

    Args:
        - config_file (ConfigFile): The configuration file to be updated.

    Returns:
        - verifika_profile (Path): The location of the Verifika profile file."""

    profiles_dir = config_file.get("DEFAULT",
                                   "verifika_profiles_location",
                                   fallback="")

    verifika_profile, _ = QFileDialog.getOpenFileName(
        caption="Choose a Verifika profile to be applied",
        directory=profiles_dir,
        filter="Verifika profile (*.vprofile)")

    if not verifika_profile:
        raise SystemExit

    new_profiles_dir = str(Path(verifika_profile).parent)
    update_config(config_file, "DEFAULT", "verifika_profiles_location",
                  new_profiles_dir)

    return Path(verifika_profile)


def browse_verifika(config_file: _ConfigFile) -> Path:
    """Prompts the user to browse for the Verifika executable
    if its location is not specified in the config file.
    If the executable is running, the user is asked to close it before continuing.
    Updates the config file with the location of the executable.

    Args:
        - config_file (ConfigFile): The configuration file to be updated.

    Returns:
        - verifika_exe_location (Path): The location of the Verifika executable."""

    _close_verifika()

    verifika_exe = Path(
        config_file.get("DEFAULT", "verifika_location", fallback=""))

    if verifika_exe.is_file():
        return verifika_exe

    new_verifika_exe, _ = QFileDialog.getOpenFileName(
        caption="Navigate to Verifika directory",
        filter="Verifika executable (verifika.exe)")

    if not new_verifika_exe:
        raise SystemExit

    update_config(config_file, "DEFAULT", "verifika_location",
                  new_verifika_exe)

    return Path(new_verifika_exe)


def _close_verifika() -> None:
    """Asks the user to close Verifika if it is already running."""

    verifika_process = _is_verifika_running()

    if not verifika_process:
        return

    answer = QMessageBox.question(
        None, "Verifika is already running",
        ("Another instance of Verifika is already running.\n"
         "Would you like to close it before continuing?\n"
         "Warning: no changes will be saved."),
        QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No)

    if answer == QMessageBox.StandardButton.No:
        raise SystemExit

    try:
        verifika_process.kill()

    except NoSuchProcess:
        pass


def _is_verifika_running() -> Optional[Process]:
    """Checks if the Verifika executable is running,
    if so returns the process if it's running."""

    for process in process_iter(attrs=["name"]):
        if process.name().lower() == "verifika.exe":
            return process

    return None
