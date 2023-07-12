"""This module provides a configuration class for managing settings and options.

The Config class provides a way to store and retrieve relevant information for the main program,
such as path to the Verifika executable,
and the directory containing necessary profiles needed for performing QA checks.

The module includes methods for reading and storing configurations to a file,
as well as for updating individual settings."""

from pathlib import Path
from configparser import ConfigParser


class _ConfigFile(ConfigParser):
    """Custom ConfigParser class that stores its own path.

    This allows it to be read without needing to pass its path to the function.
    Use ConfigFile.read(ConfigFile.path) to read the config file."""

    def __init__(self, path: str) -> None:
        ConfigParser.__init__(self)
        self.path = path
        self.read(self.path)


def create_config(config_location: str = "config.ini") -> _ConfigFile:
    """Creates a simple config file which will store the following:
        - verifika_location: Location of the Verifika executable.
        - verifika_profiles_location: Location of a directory containing Verifika profiles.

    If the config file already exists, it will be read and returned as a ConfigFile instance.
    Otherwise, a new ConfigFile instance will be created and saved to the specified location.

    Args:
        - config_location (str, optional): The path where the configuration file should be saved.
        Defaults to "config.ini".

    Returns:
        - ConfigFile: A ConfigFile instance that represents the configuration file."""

    config_file = _ConfigFile(config_location)
    config_file.read(config_location)

    if not Path(config_location).is_file():
        with open(config_location, "w", encoding="utf-8") as file:
            config_file.write(file)

    return config_file


def update_config(config_file: _ConfigFile, section: str, option: str,
                  value: str) -> None:
    """Updates the configuration file with a new value if it differs from the old one.

    Args:
        - config_file (ConfigFile): The ConfigFile instance to update.
        - section (str): The name of the section in the configuration file.
        - option (str): The name of the option to update.
        - value (str): The new value for the specified option."""

    old_val = config_file.get(section, option, fallback="")

    if value != old_val:
        config_file.set(section, option, value)

        with open(config_file.path, "w", encoding="utf-8") as file:
            config_file.write(file)
