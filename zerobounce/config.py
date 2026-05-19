import os
import json

PACKAGE_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(PACKAGE_DIR)
CONFIG_PATH = os.path.join(PROJECT_ROOT, "cleancolumns.json")


def load_config():
    with open(CONFIG_PATH, "r", encoding="utf-8") as f:
        return json.load(f)
