import json
import os
import threading
from datetime import datetime

from utils.excel_parser import parse_rxeys
from utils.back_parser import parse_back

_CONFIG_PATH = None
_LOG_PATH = None

_thread = None
_stop_event = threading.Event()
_lock = threading.Lock()
_status = {
    "running": False,
    "last_sync": None,
    "last_file": None,
    "last_imported": 0,
    "last_error": None,
    "total_synced": 0,
}

_DEFAULTS = {
    "enabled": False,
    "watch_folder": "",
    "date_col": "Tarih",
    "banka_col": "Banka",
    "nakit_col": "Nakit",
    "interval_minutes": 5,
}


def init(data_dir: str):
    global _CONFIG_PATH, _LOG_PATH
    _CONFIG_PATH = os.path.join(data_dir, "watcher_config.json")
    _LOG_PATH    = os.path.join(data_dir, "watcher_log.json")


def get_config() -> dict:
    if not _CONFIG_PATH or not os.path.exists(_CONFIG_PATH):
        return dict(_DEFAULTS)
    try:
        with open(_CONFIG_PATH, encoding="utf-8") as f:
            return {**_DEFAULTS, **json.load(f)}
    except Exception:
        return dict(_DEFAULTS)


def save_config(cfg: dict):
    if not _CONFIG_PATH:
        raise RuntimeError("watcher.init() henüz çağrılmadı")
    with open(_CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def get_status() -> dict:
    with _lock:
        return dict(_status)


def _get_log() -> dict:
    if not _LOG_PATH or not os.path.exists(_LOG_PATH):
        return {}
    try:
        with open(_LOG_PATH, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def _save_log(log: dict):
    with open(_LOG_PATH, "w", encoding="utf-8") as f:
        json.dump(log, f, ensure_ascii=False, indent=2)


def _scan_and_import(cfg: dict, db_func) -> int:
    folder = cfg.get("watch_folder", "")
    if not folder or not os.path.isdir(folder):
        return 0

    date_col  = cfg.get("date_col",  "Tarih")
    banka_col = cfg.get("banka_col", "Banka")
    nakit_col = cfg.get("nakit_col", "Nakit")

    log = _get_log()
    total = 0

    for fname in sorted(os.listdir(folder)):
        low = fname.lower()
        if not low.endswith((".xlsx", ".xls", ".back")):
            continue
        fpath = os.path.join(folder, fname)
        try:
            mtime = str(os.path.getmtime(fpath))
        except OSError:
            continue
        key = f"{fname}|{mtime}"
        if key in log:
            continue

        try:
            with open(fpath, "rb") as fp:
                file_bytes = fp.read()
            if low.endswith(".back"):
                parsed = parse_back(file_bytes)
            else:
                parsed = parse_rxeys(file_bytes, date_col, banka_col, nakit_col)
            count = db_func(parsed) if parsed else 0
            log[key] = {"file": fname, "imported": count, "at": datetime.now().isoformat()}
            total += count
            with _lock:
                _status["last_file"]     = fname
                _status["last_imported"] = count
                _status["total_synced"] += count
        except Exception as e:
            with _lock:
                _status["last_error"] = f"{fname}: {e}"

    _save_log(log)
    with _lock:
        _status["last_sync"] = datetime.now().isoformat()

    return total


def _worker(db_func):
    with _lock:
        _status["running"] = True
    try:
        while not _stop_event.is_set():
            cfg = get_config()
            if cfg.get("enabled") and cfg.get("watch_folder"):
                try:
                    _scan_and_import(cfg, db_func)
                except Exception as e:
                    with _lock:
                        _status["last_error"] = str(e)
            interval = max(1, cfg.get("interval_minutes", 5)) * 60
            _stop_event.wait(interval)
    finally:
        with _lock:
            _status["running"] = False


def start(db_func):
    global _thread
    _stop_event.clear()
    _thread = threading.Thread(target=_worker, args=(db_func,), daemon=True)
    _thread.start()


def stop():
    _stop_event.set()


def trigger(db_func) -> dict:
    cfg = get_config()
    if not cfg.get("watch_folder"):
        return {"error": "Klasör ayarlanmamış"}
    try:
        count = _scan_and_import(cfg, db_func)
        with _lock:
            return {"ok": True, "imported": count, "last_sync": _status["last_sync"]}
    except Exception as e:
        return {"error": str(e)}
