import pandas as pd
import yaml
from pathlib import Path

BASE = Path(__file__).parent.parent

def load_segments():
    return pd.read_csv(BASE / "data/raw/cbr_dcp.csv")

def load_damage():
    return pd.read_csv(BASE / "data/raw/damage_obs.csv")

def load_grade_raw():
    return pd.read_csv(BASE / "data/raw/grade.csv")

def load_config():
    with open(BASE / "config.yaml") as f:
        return yaml.safe_load(f)
