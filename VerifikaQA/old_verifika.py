"""This module allows the user to run Verifika quality assurance (QA) checks
without having to use Verifika's user interface which can speed up the QA process.

The main function of this module guides the user through the QA process by
prompting the user to select the files to be checked
and specifying the types of errors to search for.
Once this is done, the QA process is performed on the selected files and the report is generated.
The user can choose to manually check and/or correct the errors,
or let Verifika perform an automatic QA check instead.

If errors were found and thus a report was generated, the report is processed and saved.
Otherwise, a message is displayed to inform the user, and the program exits."""

import os
from subprocess import Popen, DEVNULL, STDOUT
from alive_progress import alive_bar
from gui import (CustomTk, browse_files, browse_verifika,
                 browse_verifika_profile, customize_report_params)
from verifika_config import create_config
from report_processing import process_report
from file_processing import prepare_files, delete_dir


def generate_cmd_command(files_dir: str, verifika_exe_location: str,
                         files_to_check: str, verifika_profile: str,
                         error_types: list[str],
                         manual_optimization: bool) -> str:
    """Generates a command to run the Verifika QA tool via the command line.

    Args:
        - files_dir (str): Path to the directory containing the files to be checked.
        - verifika_exe_location (str): Path to the Verifika executable.
        - files_to_check (str): Path to the file or directory containing the files to be checked.
        - verifika_profile (str): Path to the Verifika profile for the QA check.
        - error_types (list[str]): Error types to keep in the QA report.
        - manual_optimization (bool): Indication on whether manual optimization is requested.

    Returns:
        - A string of the full CMD command needed to run Verifika QA check via the command line."""

    report_type = get_report_type(error_types)
    cmd_command = (
        f'"{verifika_exe_location}" -files "{files_to_check}" '
        f'-profile "{verifika_profile}" -startcheck -type {report_type}')

    if not manual_optimization:
        temp_report_name = os.path.join(files_dir, "temp_report.xlsx")
        cmd_command += f" -result {temp_report_name}"

    return cmd_command


def run_qa_command(cmd_command: str) -> None:
    """Runs a quality assurance (QA) check by executing a specified CMD command on the command line.
    Displays an accompanying progress bar while the command runs.

    Note that the `DEVNULL` constant is used to suppress the lengthy output from Verifika, however,
    error output is still enabled.

    Args:
        - cmd_command (str): A string containing the CMD command to be executed."""

    with alive_bar(spinner=None,
                   title="Performing QA check",
                   stats=False,
                   monitor=None,
                   monitor_end="Performed in") as progress_bar:
        with Popen(cmd_command, stdout=DEVNULL, stderr=STDOUT) as process:
            process.wait()
            progress_bar()  # pylint: disable=not-callable


def get_report_type(error_types: list[str]) -> str:
    """Determines the type of report to be generated based on the types of errors user had selected.

    Args:
        - error_types (list[str]): List of errors user selected.

    Returns:
        - str: The type of report to be generated. Possible values are "Full" for a full report,
            or the name of a report preset ("Common", "Consistency", or "Spelling")."""

    report_presets = ["Common Errors", "Consistency Errors", "Spelling Errors"]

    if len(error_types) == 1 and error_types[0] in report_presets:
        return error_types[0].split(" ")[0]

    return "Full"


def verifika() -> None:
    """This function guides the user through the QA process by
    prompting the user to select the files to be checked
    and specifying the types of errors to search for.
    Once this is done, the QA process is performed on the selected files
    and the report is generated.
    The user can choose to manually check and/or correct the errors,
    or let Verifika perform an automatic QA check instead.

    If errors were found and thus a report was generated, the report is processed and saved.
    Otherwise, a message is displayed to inform the user, and the program exits."""

    root = CustomTk()
    config_file = create_config()
    verifika_exe_location = browse_verifika(root, config_file)
    files, files_dir = browse_files(root)
    verifika_profile = browse_verifika_profile(root, config_file)
    error_types, manual_optimization = customize_report_params(root)

    files_to_check = prepare_files(files)

    cmd_command = generate_cmd_command(files_dir, verifika_exe_location,
                                       files_to_check, verifika_profile,
                                       error_types, manual_optimization)

    delete_dir(files_dir)

    run_qa_command(cmd_command)

    process_report(files_dir, error_types, manual_optimization)


if __name__ == "__main__":
    verifika()
