import pandas as pd
import datetime

# --- HELPERS ---
def _get_val(row, col_name):
    """Safe value retrieval helper."""
    # Try exact match first
    if col_name in row: return row[col_name]
    # Try case insensitive
    for c in row.index:
        if c.upper() == col_name.upper(): return row[c]
    return None

# --- CUSTOM HOOKS ---

def format_date_yymmdd(row, source_col):
    """
    Converts YYYYMMDD or YYYY-MM-DD to YYMMDD.
    """
    val = str(_get_val(row, source_col))
    val = val.replace('-', '').replace('/', '').replace('.', '')
    if len(val) == 8:
        return val[2:] # 20231027 -> 231027
    return val

def math_calculate_volume(row, source_col):
    """
    Example Math: Calculates Vol = Height * Width * Depth
    Assumes columns MMHEIG, MMWIDT, MMDEPT exist in source.
    """
    try:
        h = float(_get_val(row, 'MMHEIG') or 0)
        w = float(_get_val(row, 'MMWIDT') or 0)
        d = float(_get_val(row, 'MMDEPT') or 0)
        return str(h * w * d)
    except:
        return "0"

def logic_item_status(row, source_col):
    """
    Example Case Statement:
    If Group is 900 -> Status 20
    If Group is 999 -> Status 50
    Else -> Status 10
    """
    group = str(_get_val(row, 'MMITGR'))
    
    if group == '900':
        return '20'
    elif group == '999':
        return '50'
    else:
        return '10'

def concat_description(row, source_col):
    """
    Combines Name 1 + Name 2
    """
    d1 = str(_get_val(row, 'MMITDS') or "").strip()
    d2 = str(_get_val(row, 'MMFUDS') or "").strip()
    return f"{d1} {d2}".strip()