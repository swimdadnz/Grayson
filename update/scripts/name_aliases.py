"""Canonical swimmer names for record tables. Extend after source verification."""
ALIASES = {"Samuel Brown": "Sam Brown"}

def canonical_name(name):
    value = str(name or "").strip()
    return ALIASES.get(value, value)

def canonical_holders(value):
    if not value or str(value).strip().lower() == "standard":
        return []
    return list(dict.fromkeys(canonical_name(part) for part in str(value).split(" / ") if str(part).strip()))
