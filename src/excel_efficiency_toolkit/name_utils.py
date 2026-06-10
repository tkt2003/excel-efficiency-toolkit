import re

def clean_sheet_name(name: str | None) -> str:
    if name is None:
        return ""
    # Remove invalid characters: \ / ? * [ ] :
    cleaned = re.sub(r'[\\/?*\[\]:]', '', name)
    # Strip whitespace
    cleaned = cleaned.strip()
    # Truncate to 31 characters
    return cleaned[:31]

def get_safe_sheet_name(name: str | None, fallback: str = "Sheet") -> str:
    cleaned = clean_sheet_name(name)
    if not cleaned:
        cleaned_fallback = clean_sheet_name(fallback)
        if not cleaned_fallback:
            return "Sheet"
        return cleaned_fallback
    return cleaned

def get_unique_sheet_name(base_name: str, existing_names: set[str]) -> str:
    lower_existing = {n.lower() for n in existing_names}
    
    base_cleaned = clean_sheet_name(base_name)
    if not base_cleaned:
        base_cleaned = "Sheet"
        
    if base_cleaned.lower() not in lower_existing:
        existing_names.add(base_cleaned)
        return base_cleaned
        
    counter = 2
    while True:
        suffix = f"_{counter}"
        max_base_len = 31 - len(suffix)
        candidate = f"{base_cleaned[:max_base_len]}{suffix}"
        
        if candidate.lower() not in lower_existing:
            existing_names.add(candidate)
            return candidate
        counter += 1
