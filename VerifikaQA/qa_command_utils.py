"""This module contains functions responsible for generating
and running the Verifika QA commands."""

from pathlib import Path
from subprocess import Popen, DEVNULL, STDOUT
from alive_progress import alive_bar


def generate_qa_command(files_dir: Path, verifika_exe_location: Path,
                        files_to_check: Path, verifika_profile: Path,
                        error_types: list[str], manual_report: bool) -> str:
    """Generates a command to run the Verifika QA tool via the command line.

    Args:
        - files_dir (Path): Path to the directory containing the files to be checked.
        - verifika_exe_location (Path): Path to the Verifika executable.
        - files_to_check (Path): Path to the file or directory containing the files to be checked.
        - verifika_profile (Path): Path to the Verifika profile for the QA check.
        - error_types (list[str]): Error types to keep in the QA report.
        - manual_report (bool): Indication on whether manual optimization is requested.

    Returns:
        - cmd_command (str): Full CMD command needed to run Verifika QA check."""

    report_type = _determine_report_preset(error_types)
    cmd_command = (
        f'"{verifika_exe_location}" -files "{files_to_check}" '
        f'-profile "{verifika_profile}" -startcheck -type {report_type}')

    if not manual_report:
        temp_report_name = Path(files_dir, "temp_report.xlsx")
        cmd_command += f" -result {temp_report_name}"

    return cmd_command


def run_qa_command(cmd_command: str) -> None:
    """Runs a quality assurance (QA) check by executing a specified CMD command on the command line.
    Displays an accompanying progress bar while the command runs.

    Note that the "DEVNULL" constant is used to suppress the lengthy output from Verifika,
    however, error output is left enabled.

    Args:
        - cmd_command (str): CMD command to be executed."""

    with alive_bar(spinner=None,
                   title="Performing QA check",
                   stats=False,
                   monitor=None) as progress_bar:
        with Popen(cmd_command, stdout=DEVNULL, stderr=STDOUT) as process:
            process.wait()
            progress_bar()  # pylint: disable=not-callable


def _determine_report_preset(error_types: list[str]) -> str:
    """Determines the type of report to be generated based on the types of errors user had selected.
    There are three presets: "Common", "Consistency", or "Spelling".
    Anything else will be defaulted to a "Full" report preset.

    Args:
        - error_types (list[str]): List of errors user selected.

    Returns:
        - str: Preset to be used."""

    report_presets = ["Common Errors", "Consistency Errors", "Spelling Errors"]

    if len(error_types) == 1 and error_types[0] in report_presets:
        return error_types[0].split(" ")[0]

    return "Full"
