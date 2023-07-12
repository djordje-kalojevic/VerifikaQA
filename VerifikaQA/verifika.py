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

from pathlib import Path
from PyQt6.QtWidgets import QApplication
from verifika_config import create_config
from gui import ReportTypeSelector, configure_theme
from file_browsing import browse_verifika, browse_files, browse_verifika_profile
from file_processing import prepare_files, delete_dir
from qa_command_utils import generate_qa_command, run_qa_command
from report_processing import process_report


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

    app = QApplication([])
    configure_theme()
    config_file = create_config()
    verifika_exe_location = browse_verifika(config_file)
    files, files_dir = browse_files()
    verifika_profile = browse_verifika_profile(config_file)
    main_window = ReportTypeSelector()
    main_window.show()
    app.exec()

    error_types = main_window.error_types
    if not error_types:
        raise SystemExit

    manual_report = main_window.manual_report

    files_to_check = prepare_files(files)

    cmd_command = generate_qa_command(files_dir, verifika_exe_location,
                                      files_to_check, verifika_profile,
                                      error_types, manual_report)

    delete_dir(Path(files_dir, "temp_dir"))

    run_qa_command(cmd_command)
    process_report(files_dir, error_types, manual_report)


if __name__ == "__main__":
    verifika()
