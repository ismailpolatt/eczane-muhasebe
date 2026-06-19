"""
Parse RxEys .back (ZIP → PostgreSQL dump) and extract daily rx_nakit / rx_banka.

Key table: eys_tbl_cari_hareket_v2
  odeme_tipi = 1  → Nakit (cash)
  odeme_tipi = 2  → Kredi Kartı (credit card)
  tutar           → payment amount (negative = return/reversal)
  tarih           → transaction timestamp
"""

import io
import zipfile
from collections import defaultdict
from datetime import datetime


def _parse_copy_block(stream, target_table: bytes):
    """Scan stream for COPY <target_table> ... FROM stdin; and yield parsed rows as dicts."""
    prefix = b"COPY " + target_table + b" "
    buf = b""
    found = False
    cols: list[str] = []

    while True:
        chunk = stream.read(2 * 1024 * 1024)
        if not chunk:
            break
        buf += chunk

        if not found:
            idx = buf.find(prefix)
            if idx == -1:
                buf = buf[-len(prefix):]
                continue

            # Extract column names from COPY header line
            line_end = buf.index(b"\n", idx)
            header = buf[idx:line_end].decode("utf-8", errors="replace")
            cols_part = header.split("(", 1)[1].split(")", 1)[0]
            cols = [c.strip() for c in cols_part.split(",")]
            buf = buf[line_end + 1:]
            found = True

        if found:
            # Process complete lines
            while True:
                nl = buf.find(b"\n")
                if nl == -1:
                    break
                line = buf[:nl]
                buf = buf[nl + 1:]
                decoded = line.decode("utf-8", errors="replace").rstrip("\r")
                if decoded == "\\.":
                    return
                parts = decoded.split("\t")
                row = {}
                for i, col in enumerate(cols):
                    row[col] = parts[i] if i < len(parts) else None
                yield row


def _to_float(val) -> float:
    if val is None or val == r"\N":
        return 0.0
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0


def _to_date(val: str | None) -> str | None:
    if not val or val == r"\N":
        return None
    s = val.strip()
    # PostgreSQL timestamp: "2026-06-01 14:30:00.123+03" — ISO date is always the first 10 chars
    if len(s) >= 10:
        date_part = s[:10]
        try:
            datetime.strptime(date_part, "%Y-%m-%d")
            return date_part
        except ValueError:
            pass
    return None


def parse_back(file_bytes: bytes) -> dict[str, dict]:
    """
    Parse a .back file and return {date_str: {rx_nakit, rx_banka}}.
    """
    daily: dict[str, dict] = defaultdict(lambda: {"rx_nakit": 0.0, "rx_banka": 0.0})

    with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
        inner_name = zf.namelist()[0]
        with zf.open(inner_name) as stream:
            for row in _parse_copy_block(stream, b"eys_tbl_cari_hareket_v2"):
                odeme_tipi = row.get("odeme_tipi", "")
                if odeme_tipi not in ("1", "2"):
                    continue

                tarih = _to_date(row.get("tarih"))
                if not tarih:
                    continue

                tutar = _to_float(row.get("tutar"))

                if odeme_tipi == "1":
                    daily[tarih]["rx_nakit"] += tutar
                else:
                    daily[tarih]["rx_banka"] += tutar

    # Round to 2 decimal places
    return {
        d: {"rx_nakit": round(v["rx_nakit"], 2), "rx_banka": round(v["rx_banka"], 2)}
        for d, v in daily.items()
        if v["rx_nakit"] != 0 or v["rx_banka"] != 0
    }
