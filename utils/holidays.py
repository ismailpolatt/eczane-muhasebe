import requests

_cache: dict[int, set[str]] = {}

_FIXED_TR = [
    "01-01",  # Yılbaşı
    "04-23",  # Ulusal Egemenlik ve Çocuk Bayramı
    "05-01",  # İşçi Bayramı
    "05-19",  # Atatürk'ü Anma, Gençlik ve Spor Bayramı
    "07-15",  # Demokrasi ve Millî Birlik Günü
    "08-30",  # Zafer Bayramı
    "10-29",  # Cumhuriyet Bayramı
]


def get_turkish_holidays(year: int) -> set[str]:
    """Return set of holiday date strings in 'YYYY-MM-DD' format."""
    if year in _cache:
        return _cache[year]

    try:
        resp = requests.get(
            f"https://date.nager.at/api/v3/PublicHolidays/{year}/TR",
            timeout=5,
        )
        if resp.status_code == 200:
            holidays = {h["date"] for h in resp.json()}
            _cache[year] = holidays
            return holidays
    except Exception:
        pass

    # Fallback to fixed holidays only (religious holidays vary by year)
    holidays = {f"{year}-{m}" for m in _FIXED_TR}
    _cache[year] = holidays
    return holidays
