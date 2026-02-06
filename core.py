from __future__ import annotations

import json
import sys
import traceback
from datetime import datetime
from pathlib import Path
from typing import Optional

import pandas as pd

BASE_DIR = Path(sys.executable).resolve().parent if getattr(sys, "frozen", False) else Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
HISTORY_DIR = DATA_DIR / "history"
CONFIG_PATH = DATA_DIR / "config.json"
LATEST_CSV = DATA_DIR / "latest.csv"
LOG_PATH = DATA_DIR / "update.log"

DEFAULT_CONFIG = {
    "source_path": "",
    "encoding": "auto",
    "keep_history": True,
    "use_date_template": False,
    "date_format": "%m%d",
    "auto_load_on_start": False,
    "last_updated": "",
    "last_source": "",
    "last_error": "",
    "last_history": "",
}


def ensure_data_dir() -> None:
    DATA_DIR.mkdir(parents=True, exist_ok=True)
    HISTORY_DIR.mkdir(parents=True, exist_ok=True)


def now_str() -> str:
    return datetime.now().astimezone().strftime("%Y-%m-%d %H:%M:%S %Z")


def today_str() -> str:
    return datetime.now().astimezone().strftime("%Y-%m-%d")


def load_config() -> dict:
    ensure_data_dir()
    if CONFIG_PATH.exists():
        try:
            with CONFIG_PATH.open("r", encoding="utf-8") as f:
                data = json.load(f)
            return {**DEFAULT_CONFIG, **data}
        except Exception:
            return DEFAULT_CONFIG.copy()
    return DEFAULT_CONFIG.copy()


def save_config(cfg: dict) -> None:
    ensure_data_dir()
    with CONFIG_PATH.open("w", encoding="utf-8") as f:
        json.dump(cfg, f, indent=2, ensure_ascii=False)


def log_event(message: str) -> None:
    ensure_data_dir()
    line = f"[{now_str()}] {message}\n"
    with LOG_PATH.open("a", encoding="utf-8") as f:
        f.write(line)


def normalize_path(path_str: str) -> Path:
    return Path(path_str).expanduser()


def read_csv_safely(path: Path, encoding: str) -> pd.DataFrame:
    if encoding != "auto":
        return pd.read_csv(path, encoding=encoding)

    try:
        return pd.read_csv(path)
    except UnicodeDecodeError:
        for enc in ["utf-8-sig", "utf-8", "cp950", "big5", "latin1"]:
            try:
                return pd.read_csv(path, encoding=enc)
            except UnicodeDecodeError:
                continue
        raise


def save_latest(df: pd.DataFrame) -> None:
    ensure_data_dir()
    df.to_csv(LATEST_CSV, index=False)


def save_history(df: pd.DataFrame) -> str:
    ensure_data_dir()
    history_path = HISTORY_DIR / f"history_{today_str()}.csv"
    df.to_csv(history_path, index=False)
    return str(history_path)


def update_from_path(path_str: str, encoding: str, source: str) -> pd.DataFrame:
    path = normalize_path(path_str)
    if not path.exists():
        raise FileNotFoundError(f"找不到檔案: {path}")
    df = read_csv_safely(path, encoding)
    save_latest(df)

    cfg = load_config()
    history_path = ""
    if cfg.get("keep_history", True):
        history_path = save_history(df)

    cfg["source_path"] = path_str
    cfg["encoding"] = encoding
    cfg["last_updated"] = now_str()
    cfg["last_source"] = source
    cfg["last_error"] = ""
    if history_path:
        cfg["last_history"] = history_path
    save_config(cfg)

    log_event(f"update ok (source={source}, path={path})")
    return df


def load_latest_df() -> Optional[pd.DataFrame]:
    if not LATEST_CSV.exists():
        return None
    return pd.read_csv(LATEST_CSV)


def safe_update(path_str: str, encoding: str, source: str) -> tuple[Optional[pd.DataFrame], Optional[str]]:
    try:
        df = update_from_path(path_str, encoding, source)
        return df, None
    except Exception as exc:
        cfg = load_config()
        cfg["last_error"] = f"{exc}"
        save_config(cfg)
        log_event(f"update failed: {exc}")
        log_event(traceback.format_exc())
        return None, str(exc)
