import io
import json
import os
import sqlite3
import uuid

from flask import Flask, jsonify, request, send_file, send_from_directory

from utils.excel_generator import TURKISH_MONTHS, generate_excel
from utils.excel_parser import get_columns, parse_rxeys
from utils.holidays import get_turkish_holidays
from utils.back_parser import parse_back
import utils.watcher as watcher

app = Flask(__name__, static_folder="static", static_url_path="")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50 MB upload sınırı

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "data")
UPLOAD_DIR = os.path.join(DATA_DIR, "uploads")
DB_PATH = os.path.join(DATA_DIR, "eczane.db")
VIEW_CONFIG_PATH = os.path.join(DATA_DIR, "view_config.json")


# ── Görünüm ayarı (gizli günler) ────────────────────────────────────────────────

def get_view_config() -> dict:
    default = {"hidden_weekdays": [], "hidden_dates": []}
    if not os.path.exists(VIEW_CONFIG_PATH):
        return default
    try:
        with open(VIEW_CONFIG_PATH, encoding="utf-8") as f:
            return {**default, **json.load(f)}
    except Exception:
        return default


def save_view_config(cfg: dict):
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(VIEW_CONFIG_PATH, "w", encoding="utf-8") as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


# ── Database ──────────────────────────────────────────────────────────────────

def get_db():
    os.makedirs(DATA_DIR, exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    conn.execute("""
        CREATE TABLE IF NOT EXISTS entries (
            id               INTEGER PRIMARY KEY AUTOINCREMENT,
            date             TEXT    UNIQUE NOT NULL,
            cikan            REAL    DEFAULT 0,
            teb              REAL    DEFAULT 0,
            vakifbank        REAL    DEFAULT 0,
            nakit            REAL    DEFAULT 0,
            iban             REAL    DEFAULT 0,
            iade_nakit       REAL    DEFAULT 0,
            iade_banka       REAL    DEFAULT 0,
            tahsilat_banka   REAL    DEFAULT 0,
            tahsilat_nakit   REAL    DEFAULT 0,
            rx_banka         REAL    DEFAULT 0,
            rx_nakit         REAL    DEFAULT 0,
            notes            TEXT    DEFAULT ''
        )
    """)
    # Migrate older DBs that may be missing new columns
    _add_col(conn, "iban",             "REAL DEFAULT 0")
    _add_col(conn, "iade_nakit",       "REAL DEFAULT 0")
    _add_col(conn, "iade_banka",       "REAL DEFAULT 0")
    _add_col(conn, "tahsilat_banka",   "REAL DEFAULT 0")
    _add_col(conn, "tahsilat_nakit",   "REAL DEFAULT 0")
    _add_col(conn, "rx_banka",         "REAL DEFAULT 0")
    _add_col(conn, "rx_nakit",         "REAL DEFAULT 0")
    # migrate old single tahsilat column into tahsilat_banka
    try:
        conn.execute("UPDATE entries SET tahsilat_banka = tahsilat WHERE tahsilat_banka = 0 AND tahsilat != 0")
    except Exception:
        pass
    conn.commit()
    conn.close()


def _add_col(conn, col, definition):
    try:
        conn.execute(f"ALTER TABLE entries ADD COLUMN {col} {definition}")
    except sqlite3.OperationalError:
        pass  # column already exists


# ── API ───────────────────────────────────────────────────────────────────────

@app.route("/api/holidays/<int:year>")
def holidays(year):
    return jsonify(sorted(get_turkish_holidays(year)))


@app.route("/api/entries/<int:year>/<int:month>")
def get_entries(year, month):
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM entries WHERE date LIKE ? ORDER BY date",
        (f"{year}-{month:02d}-%",),
    ).fetchall()
    conn.close()
    return jsonify([dict(r) for r in rows])


@app.route("/api/entries", methods=["POST"])
def save_entry():
    data = request.get_json()
    if not data or "date" not in data:
        return jsonify({"error": "date zorunlu"}), 400

    def _f(key):
        return float(data.get(key) or 0)

    conn = get_db()
    conn.execute(
        """
        INSERT INTO entries
            (date, cikan, teb, vakifbank, nakit, iban, iade_nakit, iade_banka,
             tahsilat_banka, tahsilat_nakit, rx_banka, rx_nakit, notes)
        VALUES
            (:date,:cikan,:teb,:vakifbank,:nakit,:iban,:iade_nakit,:iade_banka,
             :tahsilat_banka,:tahsilat_nakit,:rx_banka,:rx_nakit,:notes)
        ON CONFLICT(date) DO UPDATE SET
            cikan          = excluded.cikan,
            teb            = excluded.teb,
            vakifbank      = excluded.vakifbank,
            nakit          = excluded.nakit,
            iban           = excluded.iban,
            iade_nakit     = excluded.iade_nakit,
            iade_banka     = excluded.iade_banka,
            tahsilat_banka = excluded.tahsilat_banka,
            tahsilat_nakit = excluded.tahsilat_nakit,
            rx_banka       = excluded.rx_banka,
            rx_nakit       = excluded.rx_nakit,
            notes          = excluded.notes
        """,
        {
            "date":           data["date"],
            "cikan":          _f("cikan"),
            "teb":            _f("teb"),
            "vakifbank":      _f("vakifbank"),
            "nakit":          _f("nakit"),
            "iban":           _f("iban"),
            "iade_nakit":     _f("iade_nakit"),
            "iade_banka":     _f("iade_banka"),
            "tahsilat_banka": _f("tahsilat_banka"),
            "tahsilat_nakit": _f("tahsilat_nakit"),
            "rx_banka":       _f("rx_banka"),
            "rx_nakit":       _f("rx_nakit"),
            "notes":          data.get("notes") or "",
        },
    )
    conn.commit()
    conn.close()
    return jsonify({"ok": True})


@app.route("/api/preview-rxeys", methods=["POST"])
def preview_rxeys():
    if "file" not in request.files:
        return jsonify({"error": "Dosya bulunamadı"}), 400
    f = request.files["file"]
    file_bytes = f.read()
    if not file_bytes:
        return jsonify({"error": "Dosya boş"}), 400

    os.makedirs(UPLOAD_DIR, exist_ok=True)
    file_id = str(uuid.uuid4())
    file_path = os.path.join(UPLOAD_DIR, f"{file_id}.xlsx")
    with open(file_path, "wb") as fp:
        fp.write(file_bytes)

    columns = get_columns(file_bytes)
    if not columns:
        return jsonify({"error": "Excel başlık satırı okunamadı"}), 422

    return jsonify({"file_id": file_id, "columns": columns})


@app.route("/api/import-rxeys", methods=["POST"])
def import_rxeys():
    data = request.get_json()
    file_id   = data.get("file_id")
    date_col  = data.get("date_col")
    banka_col = data.get("banka_col")
    nakit_col = data.get("nakit_col")

    if not all([file_id, date_col, banka_col, nakit_col]):
        return jsonify({"error": "Eksik parametre"}), 400

    file_path = os.path.join(UPLOAD_DIR, f"{file_id}.xlsx")
    if not os.path.exists(file_path):
        return jsonify({"error": "Yüklenen dosya bulunamadı, tekrar yükleyin"}), 404

    with open(file_path, "rb") as fp:
        file_bytes = fp.read()

    parsed = parse_rxeys(file_bytes, date_col, banka_col, nakit_col)

    conn = get_db()
    imported = 0
    for date_str, vals in parsed.items():
        conn.execute(
            """
            INSERT INTO entries (date, rx_banka, rx_nakit)
            VALUES (:date, :rx_banka, :rx_nakit)
            ON CONFLICT(date) DO UPDATE SET
                rx_banka = excluded.rx_banka,
                rx_nakit = excluded.rx_nakit
            """,
            {"date": date_str, "rx_banka": vals["rx_banka"], "rx_nakit": vals["rx_nakit"]},
        )
        imported += 1
    conn.commit()
    conn.close()

    try:
        os.remove(file_path)
    except OSError:
        pass

    return jsonify({"ok": True, "imported": imported})


def _rx_db_save(parsed: dict) -> int:
    conn = get_db()
    count = 0
    for date_str, vals in parsed.items():
        conn.execute(
            """
            INSERT INTO entries (date, rx_banka, rx_nakit)
            VALUES (:date, :rx_banka, :rx_nakit)
            ON CONFLICT(date) DO UPDATE SET
                rx_banka = excluded.rx_banka,
                rx_nakit = excluded.rx_nakit
            """,
            {"date": date_str, "rx_banka": vals["rx_banka"], "rx_nakit": vals["rx_nakit"]},
        )
        count += 1
    conn.commit()
    conn.close()
    return count


@app.route("/api/import-back", methods=["POST"])
def import_back():
    if "file" not in request.files:
        return jsonify({"error": "Dosya bulunamadı"}), 400
    f = request.files["file"]
    file_bytes = f.read()
    if not file_bytes:
        return jsonify({"error": "Dosya boş"}), 400

    try:
        parsed = parse_back(file_bytes)
    except Exception as e:
        return jsonify({"error": f"Dosya okunamadı: {e}"}), 422

    if not parsed:
        return jsonify({"error": "Veri bulunamadı. Dosya geçerli bir RxEys yedeği mi?"}), 422

    imported = _rx_db_save(parsed)
    return jsonify({"ok": True, "imported": imported})


@app.route("/api/watcher/config", methods=["GET"])
def get_watcher_config():
    return jsonify(watcher.get_config())


@app.route("/api/watcher/config", methods=["POST"])
def set_watcher_config():
    cfg = request.get_json()
    if not cfg:
        return jsonify({"error": "Geçersiz istek"}), 400
    current = watcher.get_config()
    current.update({
        "enabled":          bool(cfg.get("enabled", current["enabled"])),
        "watch_folder":     str(cfg.get("watch_folder", current["watch_folder"])).strip(),
        "date_col":         str(cfg.get("date_col", current["date_col"])).strip(),
        "banka_col":        str(cfg.get("banka_col", current["banka_col"])).strip(),
        "nakit_col":        str(cfg.get("nakit_col", current["nakit_col"])).strip(),
        "interval_minutes": int(cfg.get("interval_minutes", current["interval_minutes"])),
    })
    watcher.save_config(current)
    return jsonify({"ok": True})


@app.route("/api/watcher/status")
def get_watcher_status():
    return jsonify(watcher.get_status())


@app.route("/api/watcher/sync", methods=["POST"])
def trigger_watcher():
    result = watcher.trigger(_rx_db_save)
    if "error" in result:
        return jsonify(result), 400
    return jsonify(result)


@app.route("/api/view-config", methods=["GET"])
def get_view_config_route():
    return jsonify(get_view_config())


@app.route("/api/view-config", methods=["POST"])
def set_view_config_route():
    cfg = request.get_json() or {}
    current = get_view_config()
    if "hidden_weekdays" in cfg:
        current["hidden_weekdays"] = sorted({int(x) for x in cfg["hidden_weekdays"]})
    if "hidden_dates" in cfg:
        current["hidden_dates"] = sorted({str(x) for x in cfg["hidden_dates"]})
    save_view_config(current)
    return jsonify({"ok": True})


@app.route("/api/export/<int:year>/<int:month>")
def export_excel(year, month):
    conn = get_db()
    rows = conn.execute(
        "SELECT * FROM entries WHERE date LIKE ? ORDER BY date",
        (f"{year}-{month:02d}-%",),
    ).fetchall()
    conn.close()

    entries_dict = {r["date"]: dict(r) for r in rows}
    holidays_set = get_turkish_holidays(year)
    vc = get_view_config()

    excel_bytes = generate_excel(
        year, month, entries_dict, holidays_set,
        hidden_weekdays=vc["hidden_weekdays"],
        hidden_dates=vc["hidden_dates"],
    )
    month_name = TURKISH_MONTHS[month] if 1 <= month <= 12 else str(month)
    filename = f"Eczane_{month_name}_{year}.xlsx"

    return send_file(
        io.BytesIO(excel_bytes),
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ── Frontend ──────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return send_from_directory("static", "index.html")


def _cleanup_old_uploads(max_age_hours: int = 2):
    """Başlangıçta yarım kalan geçici upload dosyalarını temizle."""
    import time
    if not os.path.isdir(UPLOAD_DIR):
        return
    cutoff = time.time() - max_age_hours * 3600
    for fname in os.listdir(UPLOAD_DIR):
        fpath = os.path.join(UPLOAD_DIR, fname)
        try:
            if os.path.getmtime(fpath) < cutoff:
                os.remove(fpath)
        except OSError:
            pass


if __name__ == "__main__":
    init_db()
    _cleanup_old_uploads()
    watcher.init(DATA_DIR)
    watcher.start(_rx_db_save)
    debug = os.environ.get("FLASK_DEBUG", "0") == "1"
    app.run(host="0.0.0.0", port=5000, debug=debug)
