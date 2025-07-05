import os
import pickle
from typing import Any, Dict

CONFIG_FILE = os.path.join(os.path.dirname(os.path.dirname(__file__)), "config.pkl")

DEFAULT_CONFIG: Dict[str, Any] = {
    "DOWNLOAD_DIR": "",
    "WORK_DIR": "",
    "MEAL_FEE": "5500",
    "OFFICIAL_DATA_NAMES_STR": "",
}


def load_config(path: str = CONFIG_FILE) -> Dict[str, Any]:
    """Load configuration from pickle file or return defaults."""
    if os.path.isfile(path):
        with open(path, "rb") as f:
            try:
                return pickle.load(f)
            except Exception:
                # broken file fallback to defaults
                return DEFAULT_CONFIG.copy()
    return DEFAULT_CONFIG.copy()


def save_config(config: Dict[str, Any], path: str = CONFIG_FILE) -> None:
    """Save configuration to pickle file."""
    with open(path, "wb") as f:
        pickle.dump(config, f)
