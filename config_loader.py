# config_loader.py
import os
import sys
import configparser

def get_base_path():
    """Return directory of exe (frozen) or script (source)."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

CONFIG_PATH = os.path.join(get_base_path(), 'config.ini')

def load_config():
    """Load config.ini; raise clear error if missing."""
    if not os.path.exists(CONFIG_PATH):
        raise FileNotFoundError(
            f"config.ini not found at {CONFIG_PATH}.\n"
            f"Copy config.example.ini to config.ini and fill in real values."
        )
    cfg = configparser.ConfigParser()
    cfg.read(CONFIG_PATH)
    return cfg

# Module-level cached config
_config = None

def get_config():
    global _config
    if _config is None:
        _config = load_config()
    return _config