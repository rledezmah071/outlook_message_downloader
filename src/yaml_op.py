from typing import Dict
from yaml import safe_load


# Typing section
YAML_output = Dict[str, str]


def load_yaml_file(yaml_file_path: str) -> YAML_output:
    with open(yaml_file_path, mode='r') as yaml_file:
        return safe_load(yaml_file)
