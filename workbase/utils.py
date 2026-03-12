def parse_number(s):
    try:
        if isinstance(s, (int, float)):
            return float(s)
        s = str(s).strip().replace(',', '.')
        return float(s) if s else 0.0
    except Exception:
        return 0.0