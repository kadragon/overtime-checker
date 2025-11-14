import os
import json
from typing import Any, Dict

# Configuration keys as constants to prevent typos
DOWNLOAD_DIR_KEY = "DOWNLOAD_DIR"
WORK_DIR_KEY = "WORK_DIR"
MEAL_FEE_KEY = "MEAL_FEE"
OFFICIAL_DATA_NAMES_STR_KEY = "OFFICIAL_DATA_NAMES_STR"

CONFIG_FILE = os.path.join(os.path.dirname(os.path.dirname(__file__)), "config.json")

DEFAULT_CONFIG: Dict[str, Any] = {
    DOWNLOAD_DIR_KEY: "",
    WORK_DIR_KEY: "",
    MEAL_FEE_KEY: "5500",
    OFFICIAL_DATA_NAMES_STR_KEY: "",
}


def load_config(path: str = CONFIG_FILE) -> Dict[str, Any]:
    """Load configuration from json file or return defaults."""
    if os.path.isfile(path):
        with open(path, "r", encoding="utf-8") as f:
            try:
                return json.load(f)
            except json.JSONDecodeError:
                # broken file fallback to defaults
                return DEFAULT_CONFIG.copy()
    return DEFAULT_CONFIG.copy()


def save_config(config: Dict[str, Any], path: str = CONFIG_FILE) -> None:
    """Save configuration to json file."""
    with open(path, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)
