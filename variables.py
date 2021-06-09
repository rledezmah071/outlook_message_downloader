import os
import pathlib


MAIN_PATH = str(pathlib.Path(__file__).parent.absolute())
DEFAULT_SETTINGS_PATH = os.sep.join([MAIN_PATH, "etc", "settings.yml"])
