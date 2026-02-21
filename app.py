import io
import re
import json
import os
import openpyxl
from openpyxl import Workbook
import xlrd
import streamlit as st

# ─────────────────────────────────────────────────────────────────────────────
# Config
# ─────────────────────────────────────────────────────────────────────────────
SCRIPT_DIR     = os.path.dirname(os.path.abspath(__file__))
TEMPLATES_FILE = os.path.join(SCRIPT_DIR, "templates.json")

OUT_HEADERS = [
    "Item ID", "Item Name", "Product/Service",
    "Item Type (Buy/Sell/Both)", "Unit of Measurement", "HSN Code",
    "Item Category", "Default Price", "Regular Buying Price",
    "Wholesale Buying Price", "Regular Selling Price", "MRP",
    "Dealer Price", "Distributor Price", "Current Stock",
    "Min Stock Level", "Max Stock Level", "Tax"
]

TARGET_COLS = [
    "Item Name", "Item ID", "HSN Code",
    "Item Category", "Unit of Measurement", "Tax", "Product/Service"
]

# Weighted keywords: (keyword, weight) — higher weight = stronger signal
WEIGHTED_KEYWORDS = {
    "Item Name":            [("name", 3), ("item", 2), ("product", 2), ("description", 1), ("goods", 1)],
    "Item ID":              [("id", 3), ("code", 2), ("sku", 3), ("no.", 1)],
    "HSN Code":             [("hsn", 3), ("sac", 2)],
    "Item Category":        [("category", 3), ("group", 2), ("parent", 2), ("class", 1), ("family", 1)],
    "Unit of Measurement":  [("unit", 3), ("uom", 3), ("measure", 2), ("base", 1)],
    "Tax":                  [("tax", 3), ("gst", 2), ("igst", 2), ("cgst", 2)],
    "Product/Service":      [("service", 2), ("supply", 2)],
}

# System/UI columns to skip when collecting extra columns
SYSTEM_COLS = {
    "edit", "delete", "checkboxvalue", "checkbox", "widget", "inactive",
    "timestamp", "created", "modified", "updated", "sno", "srno",
    "s.no", "sr.no", "s no", "sr no", "serial no", "active", "status",
    "action", "flag", "chk"
}

ILLEGAL_RE = re.compile(r'[\x00-\x08\x0b-\x0c\x0e-\x1f]')

# Values that identify a column as "Item Category" regardless of column name.
# All entries are pre-normalised (lowercase, no spaces/underscores/hyphens).
# Incoming cell values are normalised the same way before comparing.
def _norm(s):
    return re.sub(r'[\s_\-]+', '', s.lower())

CATEGORY_NORM_TOKENS = {
    "rm", "fg", "sfg", "wip",
    "rawmaterial", "rawmaterials",
    "finishedgood", "finishedgoods",
    "finishproduct", "finishedproduct", "finishgoods",
    "semifinished", "semifinishedgood", "semifinishedgoods",
    "workinprogress",
    "consumable", "consumables",
    "trading", "spares", "spareparts",
    "packingmaterial", "packingmaterials",
    "noninventoryitem", "noninventory",
}

# ─────────────────────────────────────────────────────────────────────────────
# Core helpers
# ─────────────────────────────────────────────────────────────────────────────

def clean(v):
    return ILLEGAL_RE.sub('', v).strip() if isinstance(v, str) else v

def load_templates():
    if os.path.exists(TEMPLATES_FILE):
        with open(TEMPLATES_FILE) as f:
            return json.load(f)
    return {}

def save_templates(data):
    with open(TEMPLATES_FILE, 'w') as f:
        json.dump(data, f, indent=2)

def _fmt_cell(val, fmt):
    """
    Apply Excel number format to preserve leading zeros.
    A format like '00000000' means zero-pad the integer to that many digits.
    """
    if not isinstance(val, (int, float)) or not fmt:
        return val
    # Pure zero-padding format: only '0' digits, no decimal or special chars
    if re.match(r'^0+$', fmt):
        return str(int(val)).zfill(len(fmt))
    return val

def read_file(file_bytes):
    """Read all sheets from xls or xlsx. Returns {sheet_name: [[row_values]]}"""
    is_xls = file_bytes[:4] == b'\xd0\xcf\x11\xe0'
    sheets = {}
    if is_xls:
        wb = xlrd.open_workbook(file_contents=file_bytes, formatting_info=True)
        xf_list   = wb.xf_list
        fmt_map   = wb.format_map
        for ws in wb.sheets():
            rows = []
            for r in range(ws.nrows):
                row = []
                for c in range(ws.ncols):
                    cell = ws.cell(r, c)
                    if cell.ctype == xlrd.XL_CELL_EMPTY:
                        row.append(None)
                    elif cell.ctype == xlrd.XL_CELL_NUMBER:
                        v = cell.value
                        num = int(v) if v == int(v) else v
                        # Try to apply leading-zero format from xf record
                        try:
                            xf  = xf_list[cell.xf_index]
                            fmt = fmt_map[xf.format_key].format_str
                            num = _fmt_cell(num, fmt)
                        except Exception:
                            pass
                        row.append(num)
                    else:
                        v = str(cell.value).strip()
                        row.append(v or None)
                rows.append(row)
            sheets[ws.name] = rows
    else:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes))
        for ws in wb.worksheets:
            sheet_rows = []
            for row in ws.iter_rows():
                row_data = []
                for cell in row:
                    val = cell.value
                    if val is not None:
                        val = _fmt_cell(val, cell.number_format)
                    row_data.append(val)
                sheet_rows.append(row_data)
            sheets[ws.title] = sheet_rows
    return sheets

def pick_sheet(sheets):
    """Return the sheet with the most populated rows."""
    return max(sheets, key=lambda s: sum(
        1 for r in sheets[s]
        if any(v is not None and str(v).strip() for v in r)
    ))

def detect_header(rows):
    """
    Scan first 30 rows to find the header row.
    Scores by: number of string cells + keyword hits × 2.
    """
    HWORDS = {
        "name", "item", "product", "hsn", "unit", "uom", "category", "group",
        "code", "type", "tax", "gst", "description", "id", "price", "qty",
        "quantity", "serial", "service", "rate", "no", "sno", "sr", "brand",
        "barcode", "rack", "sac", "measure", "sku"
    }
    best_i, best_s = 0, -1
    for i, row in enumerate(rows[:30]):
        non_null = [v for v in row if v is not None and str(v).strip()]
        if len(non_null) < 2:
            continue
        strings  = sum(1 for v in non_null if isinstance(v, str))
        keywords = sum(1 for v in non_null if any(w in str(v).lower() for w in HWORDS))
        score = strings + keywords * 2
        if score > best_s:
            best_s, best_i = score, i
    return best_i, [str(v).strip() if v is not None else None for v in rows[best_i]]

def is_system_col(h):
    """Return True if column is a UI/system column that should be skipped."""
    if not h:
        return True
    hl = h.strip().lower().rstrip(".")
    if hl in SYSTEM_COLS:
        return True
    if re.match(r'^s\.?r?\.?n\.?o\.?\s*$', hl):   # SNo, SrNo, S.No, Sr.No …
        return True
    return False

def detect_category_by_values(headers, data_rows, skip_cols):
    """
    Scan column values to find a column that contains category-type data
    (RM, FG, SFG, Raw Material, Finished Goods, etc.).

    Returns the column header if found, else None.
    A column qualifies if ≥40% of its unique non-null values match known category tokens.
    """
    col_idx = {h: i for i, h in enumerate(headers) if h}
    best_col, best_score = None, 0

    for h, i in col_idx.items():
        if h in skip_cols or is_system_col(h):
            continue
        unique_vals = set()
        for row in data_rows:
            if i < len(row) and row[i] is not None:
                v = str(row[i]).strip().lower()
                if v:
                    unique_vals.add(v)
        if not unique_vals:
            continue

        # Count how many unique values match category tokens (normalised)
        hits = sum(
            1 for v in unique_vals
            if _norm(v) in CATEGORY_NORM_TOKENS
            or any(tok in _norm(v) for tok in CATEGORY_NORM_TOKENS if len(tok) > 3)
        )
        ratio = hits / len(unique_vals)
        if ratio >= 0.4 and ratio > best_score:
            best_score, best_col = ratio, h

    return best_col


def auto_map(headers, data_rows):
    """
    Map client column headers to target columns automatically.

    Strategy:
      1. Exact name match (case-insensitive)
      2. Category detection by column VALUES (RM / FG / SFG patterns)
      3. Weighted keyword scoring for remaining columns
      4. Item ID fallback: uniqueness check on all unmapped columns

    Returns:
      mapping    – {target_col: client_col_header}
      extra_cols – list of useful unmapped column headers
    """
    result, used = {}, set()

    # ── Pass 1: exact name match ─────────────────────────────────────────
    for h in headers:
        if not h:
            continue
        for t in TARGET_COLS:
            if t not in used and h.lower() == t.lower():
                result[t] = h
                used.add(t)
                break

    # ── Pass 2: category detection by column VALUES ──────────────────────
    if "Item Category" not in result:
        cat_col = detect_category_by_values(headers, data_rows, set(result.values()))
        if cat_col:
            result["Item Category"] = cat_col
            used.add("Item Category")

    # ── Pass 3: weighted keyword scoring ────────────────────────────────
    def _kw_match(kw, h_l):
        # Short keywords (≤2 chars, e.g. "id", "no") need a whole-word match
        # to avoid false hits like "id" inside "widget" or "modified".
        if len(kw) <= 2:
            return bool(re.search(r'\b' + re.escape(kw) + r'\b', h_l))
        return kw in h_l

    mapped_cols  = set(result.values())   # source cols already assigned
    used_sources = set(mapped_cols)        # keep this updated during the loop
    cands = []
    for h in headers:
        if not h or h in mapped_cols or is_system_col(h):
            continue
        h_l = h.lower()
        for t, wkws in WEIGHTED_KEYWORDS.items():
            if t in used:
                continue
            score = sum(w for kw, w in wkws if _kw_match(kw, h_l))
            if score:
                cands.append((score, t, h))

    for score, t, h in sorted(cands, reverse=True):
        if t not in used and h not in used_sources:
            result[t] = h
            used.add(t)
            used_sources.add(h)  # prevent same source col from mapping to two targets

    # ── Pass 4: Item ID by uniqueness (if still not found) ───────────────
    if "Item ID" not in result and data_rows:
        col_idx    = {h: i for i, h in enumerate(headers) if h}
        mapped_set = set(result.values())
        # Keywords that signal "this column holds an item identifier"
        ID_KEYWORDS = {"id", "code", "sku", "no", "number", "num",
                       "ref", "key", "barcode", "upc", "ean"}

        candidates = []
        for h, i in col_idx.items():
            if h in mapped_set or is_system_col(h):
                continue
            vals = [
                str(row[i]).strip() for row in data_rows
                if i < len(row) and row[i] is not None and str(row[i]).strip()
            ]
            if len(vals) < 5:
                continue

            # Skip sequential integers — those are just row/serial numbers
            try:
                nums = sorted(int(v) for v in vals)
                if nums == list(range(nums[0], nums[0] + len(nums))):
                    continue
            except (ValueError, TypeError):
                pass

            ratio = len(set(vals)) / len(vals)
            if ratio < 0.95:
                continue

            # Score by how "ID-like" the column name looks
            h_l      = h.lower()
            id_score = sum(1 for kw in ID_KEYWORDS if kw in h_l)
            candidates.append((id_score, ratio, h))

        if candidates:
            # Best = highest id_name_score first, then highest uniqueness ratio
            candidates.sort(key=lambda x: (x[0], x[1]), reverse=True)
            result["Item ID"] = candidates[0][2]

    # ── Extra cols: useful unmapped columns ──────────────────────────────
    mapped_set = set(result.values())
    extra_cols = [
        h for h in headers
        if h and h not in mapped_set and not is_system_col(h)
    ]

    return result, extra_cols

def find_template(headers, templates):
    """Find best matching saved template (≥85% column overlap)."""
    fp = set(h for h in headers if h)
    best_name, best_r = None, 0
    for name, tmpl in templates.items():
        t_fp = set(tmpl.get("fingerprint", []))
        if not t_fp:
            continue
        r = len(fp & t_fp) / max(len(fp), len(t_fp))
        if r > best_r:
            best_r, best_name = r, name
    if best_r >= 0.85:
        return best_name, templates[best_name]
    return None, None

def do_convert(rows, header_idx, mapping, extra_cols):
    """Convert rows to Product_Add format, appending extra columns at end."""
    hdrs = [str(v).strip() if v is not None else None for v in rows[header_idx]]
    col  = {h: i for i, h in enumerate(hdrs) if h}
    tidx = {t: col[c] for t, c in mapping.items() if c and c in col}
    eidx = [(h, col[h]) for h in extra_cols if h in col]

    out = []
    for row in rows[header_idx + 1:]:
        if not row or all(v is None or str(v).strip() == "" for v in row):
            continue

        def g(t):
            i = tidx.get(t)
            if i is None or i >= len(row):
                return None
            v = row[i]
            return clean(str(v)) if (v is not None and str(v).strip()) else None

        name = g("Item Name")
        if not name:
            continue

        ps    = g("Product/Service")
        ps_out = "Service" if (ps and "service" in ps.lower()) else "Product"

        tax = g("Tax")
        if not tax or tax == "0":
            tax_out = None
        else:
            try:
                tax_out = int(float(tax))
            except (ValueError, TypeError):
                tax_out = None

        row_out = [
            g("Item ID"), name, ps_out, "Both",
            g("Unit of Measurement") or "", g("HSN Code"),
            g("Item Category") or "",
            None, None, None, None, None, None, None,
            None, None, None, tax_out
        ]
        for _, ei in eidx:
            v = row[ei] if ei < len(row) else None
            row_out.append(clean(str(v)) if (v is not None and str(v).strip()) else None)
        out.append(row_out)
    return out

def make_xlsx(rows, extra_col_headers):
    wb = Workbook()
    ws = wb.active
    ws.title = "Data"
    ws.append(OUT_HEADERS + list(extra_col_headers))
    for row in rows:
        ws.append(row)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def out_filename(fname):
    m = re.search(r'\(([^)]+)\)', fname)
    cn = m.group(1) if m else fname.rsplit('.', 1)[0]
    return f"Product_Add_({cn}).xlsx"

# ─────────────────────────────────────────────────────────────────────────────
# Network Master helpers
# ─────────────────────────────────────────────────────────────────────────────

NETWORK_OUT_HEADERS = [
    "Company Name", "Buyer/Supplier/Both", "Company Reference ID",
    "TCS Type", "Company Email", "Company Contact Number",
    "Address Line 1", "Address Line 2", "City", "State", "Country",
    "PIN Code", "GSTIN", "GSTIN Type",
    "Contact Person First Name", "Contact Person Last Name", "Contact Person Email"
]

@st.cache_data(show_spinner=False)
def load_pincode_db():
    path = os.path.join(SCRIPT_DIR, "pincode_db.json")
    if os.path.exists(path):
        with open(path) as f:
            return json.load(f)
    return {}

def clean_gstin(v):
    if not v:
        return None
    s = str(v).strip()
    if s.lower() in ("none", "n/a", "", "0"):
        return None
    return s

def clean_pin(v):
    if not v:
        return None
    digits = re.sub(r"[^\d]", "", str(v).strip())
    return digits[:6] if len(digits) >= 6 else (digits if digits else None)

def split_name(full):
    if not full:
        return None, None
    parts = full.strip().split(None, 1)
    return (parts[0], parts[1]) if len(parts) == 2 else (parts[0], None)

def parse_mshriy_address(addr_str):
    """
    Parse a combined address like
    'D 24, SUNPLAZA, VADSAR ROAD, Vadodara, Gujarat, 390010'
    into (addr1, addr2, city, state, pincode).
    Convention: last = pincode, prev = state, prev-prev = city, rest = address.
    """
    if not addr_str:
        return None, None, None, None, None
    parts = [p.strip() for p in addr_str.split(",") if p.strip()]
    pincode = city = state = None

    # Find pincode — last purely-numeric 6-digit segment
    pin_idx = None
    for i in range(len(parts) - 1, -1, -1):
        d = re.sub(r"[^\d]", "", parts[i])
        if len(d) == 6:
            pincode = d
            pin_idx = i
            break

    if pin_idx is not None and pin_idx >= 2:
        state = parts[pin_idx - 1]
        city  = parts[pin_idx - 2]
        addr_parts = parts[: pin_idx - 2]
    elif pin_idx is not None and pin_idx == 1:
        state = parts[0]
        addr_parts = []
    else:
        addr_parts = parts[: pin_idx] if pin_idx else parts

    # Split address into two lines at midpoint
    if not addr_parts:
        addr1 = addr2 = None
    elif len(addr_parts) == 1:
        addr1, addr2 = addr_parts[0], None
    else:
        mid = max(1, len(addr_parts) // 2)
        addr1 = ", ".join(addr_parts[:mid]) or None
        addr2 = ", ".join(addr_parts[mid:]) or None

    return addr1, addr2, city, state, pincode

def _tally_detect_sheet(sheets):
    """Pick the SVNaturalLanguage sheet from Tally exports, else most populated."""
    if "SVNaturalLanguage" in sheets:
        return "SVNaturalLanguage"
    for name, rows in sheets.items():
        for row in rows[:5]:
            vals = [str(v).lower() for v in row if v]
            if any("$name" in v for v in vals):
                return name
    return pick_sheet(sheets)

def detect_network_format(rows):
    """
    Returns ('tally', header_idx) or ('mshriy', header_idx) or ('unknown', 0).
    """
    for i, row in enumerate(rows[:15]):
        vals = [str(v).strip().lower() for v in row if v is not None and str(v).strip()]
        joined = " ".join(vals)
        if "$name" in joined or "$_primarygroup" in joined:
            return "tally", i
        if "name of ledger" in joined or ("sl" in joined and "name" in joined):
            return "mshriy", i
    return "unknown", 0

def _g(row, col_map, *keys):
    """Get first non-empty value from row by column name keys."""
    for k in keys:
        i = col_map.get(k)
        if i is not None and i < len(row):
            v = row[i]
            if v is not None:
                s = clean(str(v))
                if s and s.lower() not in ("none", "0.00", "0"):
                    return s
    return None

def convert_tally_parties(rows, header_idx):
    hdrs = [str(v).strip() if v is not None else "" for v in rows[header_idx]]
    col  = {h: i for i, h in enumerate(hdrs)}
    parties = []

    for row in rows[header_idx + 1:]:
        if not row or all(v is None or str(v).strip() == "" for v in row):
            continue
        name = _g(row, col, "$Name")
        if not name:
            continue

        group   = _g(row, col, "$_PrimaryGroup") or ""
        group_l = group.lower()
        if "sundry debtor" in group_l:
            party_type, is_red = "Buyer", False
        elif "sundry creditor" in group_l:
            party_type, is_red = "Supplier", False
        else:
            party_type, is_red = group, True

        addr1  = _g(row, col, "$_Address1")
        addr2  = _g(row, col, "$_Address2")
        addr3  = _g(row, col, "$_Address3")
        addr2  = (addr2 + ", " + addr3) if addr2 and addr3 else (addr2 or addr3)
        state  = _g(row, col, "$PriorStateName")
        country= _g(row, col, "$CountryName") or "India"
        pin    = clean_pin(_g(row, col, "$pincode", "$Pincode"))
        gstin  = clean_gstin(_g(row, col, "$_PartyGSTIN", "$PartyGSTIN"))
        mobile = _g(row, col, "$LedgerMobile")
        email  = _g(row, col, "$email", "$Email")
        fname, lname = split_name(_g(row, col, "$LedgerContact"))

        parties.append({
            "Company Name": name,
            "Buyer/Supplier/Both": party_type,
            "Company Reference ID": None, "TCS Type": None,
            "Company Email": email,
            "Company Contact Number": mobile,
            "Address Line 1": addr1, "Address Line 2": addr2,
            "City": None, "State": state, "Country": country,
            "PIN Code": pin,
            "GSTIN": gstin, "GSTIN Type": "Regular" if gstin else None,
            "Contact Person First Name": fname,
            "Contact Person Last Name": lname,
            "Contact Person Email": None,
            "_is_red": is_red,
        })
    return parties

def convert_mshriy_parties(rows, header_idx):
    hdrs = [str(v).strip() if v is not None else "" for v in rows[header_idx]]
    col  = {h: i for i, h in enumerate(hdrs)}
    parties = []

    for row in rows[header_idx + 1:]:
        if not row or all(v is None or str(v).strip() == "" for v in row):
            continue
        name = _g(row, col, "Name of Ledger")
        if not name:
            continue

        group   = _g(row, col, "Under") or ""
        group_l = group.lower()
        if "sundry debtor" in group_l:
            party_type, is_red = "Buyer", False
        elif "sundry creditor" in group_l:
            party_type, is_red = "Supplier", False
        else:
            party_type, is_red = group, True

        addr_raw              = _g(row, col, "Address")
        addr1, addr2, city, state_p, pin_p = parse_mshriy_address(addr_raw)
        state  = _g(row, col, "State Name") or state_p
        pin    = clean_pin(_g(row, col, "Pincode", "PIN", "Pin Code")) or clean_pin(pin_p)
        gstin  = clean_gstin(_g(row, col, "GSTIN/UIN", "GSTIN"))
        email  = _g(row, col, "Mail ID", "Email")
        mobile = _g(row, col, "Contact No.", "Mobile", "Phone")

        parties.append({
            "Company Name": name,
            "Buyer/Supplier/Both": party_type,
            "Company Reference ID": None, "TCS Type": None,
            "Company Email": email,
            "Company Contact Number": mobile,
            "Address Line 1": addr1, "Address Line 2": addr2,
            "City": city, "State": state, "Country": "India",
            "PIN Code": pin,
            "GSTIN": gstin, "GSTIN Type": "Regular" if gstin else None,
            "Contact Person First Name": None,
            "Contact Person Last Name": None, "Contact Person Email": None,
            "_is_red": is_red,
        })
    return parties

def apply_pincode_lookup(parties, pincode_db):
    for p in parties:
        pin = p.get("PIN Code")
        if pin:
            entry = pincode_db.get(str(pin).zfill(6))
            if entry:
                if not p.get("City"):
                    p["City"] = entry.get("c") or None
                if not p.get("State"):
                    p["State"] = entry.get("s") or None
    return parties

def split_network_sheets(parties):
    """
    Sheet 1 (Ready to upload)  : Address Line 1 AND PIN Code both present
    Sheet 2 (Have GSTIN)       : Not in sheet 1, but GSTIN present
    Sheet 3 (Need to update manually): everything else
    """
    ready, have_gstin, manual = [], [], []
    for p in parties:
        if p.get("Address Line 1") and p.get("PIN Code"):
            ready.append(p)
        elif p.get("GSTIN"):
            have_gstin.append(p)
        else:
            manual.append(p)
    return ready, have_gstin, manual

def make_network_xlsx(ready, have_gstin, manual):
    from openpyxl.styles import PatternFill
    RED_FILL = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")

    wb  = Workbook()
    def write_sheet(ws, parties):
        ws.append(NETWORK_OUT_HEADERS)
        for p in parties:
            ws.append([p.get(h) for h in NETWORK_OUT_HEADERS])
            if p.get("_is_red"):
                r = ws.max_row
                for c in range(1, len(NETWORK_OUT_HEADERS) + 1):
                    ws.cell(row=r, column=c).fill = RED_FILL

    ws1 = wb.active;  ws1.title = "Ready to upload";     write_sheet(ws1, ready)
    ws2 = wb.create_sheet("Have GSTIN");                  write_sheet(ws2, have_gstin)
    ws3 = wb.create_sheet("Need to update manually");     write_sheet(ws3, manual)

    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()

def net_filename(fname):
    m  = re.search(r"\(([^)]+)\)", fname)
    cn = m.group(1) if m else fname.rsplit(".", 1)[0]
    return f"Network_Add_({cn}).xlsx"

def read_sheet_rows(file_bytes, sheet_name=None):
    sheets = read_file(file_bytes)
    if sheet_name:
        rows = sheets.get(sheet_name)
        if rows is None:
            raise KeyError(sheet_name)
    else:
        rows = sheets[pick_sheet(sheets)]
    return rows, list(sheets.keys())

# ─────────────────────────────────────────────────────────────────────────────
# Page config & custom CSS
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="TranZact · Item Converter",
    page_icon="⚡",
    layout="centered"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Outfit:wght@400;600;700;900&display=swap');

/* ── Top logo bar ── */
.logo-bar {
    display: flex;
    align-items: center;
    gap: 12px;
    padding: 18px 0 8px 0;
}

/* ── Gradient header bar ── */
.header-bar {
    background: linear-gradient(135deg, #1a1b17 0%, #2a1f2d 100%);
    border: 1px solid #292a27;
    border-left: 4px solid #da5d37;
    border-radius: 10px;
    padding: 16px 24px;
    margin: 4px 0 16px 0;
}
.header-bar p {
    color: #aaaca6 !important;
    font-size: 0.92rem !important;
    margin: 0 !important;
}

/* ── Card-style containers ── */
[data-testid="stExpander"] {
    border: 1px solid #292a27 !important;
    border-radius: 10px !important;
}

/* ── Upload area ── */
[data-testid="stFileUploader"] {
    border: 2px dashed #da5d37 !important;
    border-radius: 10px !important;
    padding: 8px !important;
}

/* ── Subtle divider ── */
hr { border-color: #292a27 !important; }

/* ── Tab styling ── */
[data-testid="stTab"] { font-weight: 600; }
</style>
""", unsafe_allow_html=True)

# ── Logo ──
st.markdown("""
<div class="logo-bar">

  <!-- Colourful X icon (4 arms, each a different brand colour) -->
  <svg width="42" height="42" viewBox="0 0 42 42" style="flex-shrink:0">
    <!-- bottom-left arm : yellow/amber -->
    <line x1="21" y1="21" x2="4"  y2="38" stroke="#F5A623" stroke-width="9" stroke-linecap="round"/>
    <!-- top-right arm   : green -->
    <line x1="21" y1="21" x2="38" y2="4"  stroke="#5CB85C" stroke-width="9" stroke-linecap="round"/>
    <!-- top-left arm    : teal/blue -->
    <line x1="21" y1="21" x2="4"  y2="4"  stroke="#4A90D9" stroke-width="9" stroke-linecap="round"/>
    <!-- bottom-right arm: coral/orange -->
    <line x1="21" y1="21" x2="38" y2="38" stroke="#E05A3C" stroke-width="9" stroke-linecap="round"/>
    <!-- centre cap to clean up overlap -->
    <circle cx="21" cy="21" r="5" fill="#0e0f0c"/>
  </svg>

  <!-- Wordmark -->
  <span style="font-family:'Outfit',sans-serif; font-size:1.75rem; font-weight:900;
               letter-spacing:-0.5px; color:#ffffff; line-height:1;">TRANZACT</span>

  <!-- AI badge -->
  <span style="display:inline-flex; align-items:center; gap:3px;
               border:2px solid #E07A5F; border-radius:7px;
               padding:3px 9px 3px 7px;
               font-family:'Outfit',sans-serif; font-size:0.85rem;
               font-weight:700; color:#ffffff; line-height:1; margin-left:2px;">
    Ai&nbsp;<span style="color:#F5D76E; font-size:0.75rem; vertical-align:middle;">✦</span>
  </span>

</div>
<div class="header-bar">
  <p>Item Master Converter &nbsp;·&nbsp; Upload any client file — auto-detects columns &amp; converts instantly</p>
</div>
""", unsafe_allow_html=True)

st.divider()
tab1, tab2, tab3, tab4 = st.tabs(["📥  Item Master", "📤  My Format → Tally", "🗂  Templates", "🏢  Network Master"])

# ─────────────────────────────────────────────────────────────────────────────
# TAB 1 : Any client format  →  Product_Add_(X).xlsx
# ─────────────────────────────────────────────────────────────────────────────
with tab1:
    st.markdown("Upload any client file — columns are **auto-detected and mapped** to your format.")

    up = st.file_uploader("Upload client file", type=["xlsx", "xls"], key="t1_up",
                          label_visibility="collapsed")

    if up:
        up.seek(0)
        fbytes = up.read()
        fname  = up.name

        # Re-process only when a different file is uploaded
        if st.session_state.get("t1_fname") != fname:
            sheets     = read_file(fbytes)
            sname      = pick_sheet(sheets)
            rows       = sheets[sname]
            hidx, hdrs = detect_header(rows)
            data_rows  = [r for r in rows[hidx + 1:] if any(v for v in r)]
            templates  = load_templates()
            tmpl_name, tmpl = find_template(hdrs, templates)

            if tmpl:
                mapping    = tmpl["mapping"].copy()
                extra_cols = tmpl.get("extra_cols", [])
            else:
                mapping, extra_cols = auto_map(hdrs, data_rows)

            st.session_state.t1_fname   = fname
            st.session_state.t1_rows    = rows
            st.session_state.t1_hidx    = hidx
            st.session_state.t1_headers = hdrs
            st.session_state.t1_tname   = tmpl_name
            st.session_state.t1_mapping = mapping
            st.session_state.t1_extra   = extra_cols
            st.session_state.t1_out     = None

        rows       = st.session_state.t1_rows
        hidx       = st.session_state.t1_hidx
        hdrs       = st.session_state.t1_headers
        tname      = st.session_state.t1_tname
        mapping    = st.session_state.t1_mapping
        extra_cols = st.session_state.t1_extra

        st.success(f"✅  **{fname}** uploaded")
        if tname:
            st.info(f"📋  Matched template: **{tname}**")

        if st.button("▶  Convert Now", type="primary",
                     use_container_width=True, key="t1_go"):
            result = do_convert(rows, hidx, mapping, extra_cols)
            if result:
                st.session_state.t1_out      = make_xlsx(result, extra_cols)
                st.session_state.t1_out_name = out_filename(fname)
                st.session_state.t1_out_cnt  = len(result)
            else:
                st.error("❌  No items found — could not identify the Item Name column.")

        # ── Download ──────────────────────────────────────────────────────
        if st.session_state.get("t1_out") and st.session_state.get("t1_fname") == fname:
            cnt = st.session_state.t1_out_cnt
            st.success(f"✅  **{cnt} items** converted successfully!")
            st.download_button(
                "⬇️  Download Product_Add File",
                data=st.session_state.t1_out,
                file_name=st.session_state.t1_out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
                key="t1_dl"
            )

            # ── Save as template (shown only for new/unknown formats) ─────
            if not tname:
                with st.expander("💾  Save mapping as template for next time"):
                    sc1, sc2 = st.columns([3, 1])
                    tname_in = sc1.text_input(
                        "Template name:",
                        placeholder="e.g.  MSHRIY  /  BOM Client  /  ITEM",
                        key="t1_tname_in",
                        label_visibility="collapsed"
                    )
                    if sc2.button("Save", key="t1_save_btn", use_container_width=True):
                        if tname_in.strip():
                            all_t = load_templates()
                            all_t[tname_in.strip()] = {
                                "fingerprint": sorted(h for h in hdrs if h),
                                "mapping":     {k: v for k, v in mapping.items() if v},
                                "extra_cols":  extra_cols
                            }
                            save_templates(all_t)
                            st.session_state.t1_tname = tname_in.strip()
                            st.success(f"✅  Saved as **{tname_in.strip()}**!")

# ─────────────────────────────────────────────────────────────────────────────
# TAB 2 : Product_Add_(X).xlsx  →  tranzact_item_(X).xlsx  (Tally import)
# ─────────────────────────────────────────────────────────────────────────────
with tab2:
    st.markdown("Upload your **Product Add file** and download it in the **Tally import** format.")

    uploaded_mine = st.file_uploader(
        "Upload My File", type=["xlsx", "xls"],
        key="mine_upload", label_visibility="collapsed"
    )

    if uploaded_mine:
        st.success(f"✅  **{uploaded_mine.name}** uploaded")

        if st.button("▶  Convert to Tally Format", type="primary",
                     use_container_width=True, key="btn_tab2"):
            with st.spinner("Converting..."):
                try:
                    uploaded_mine.seek(0)
                    file_bytes = uploaded_mine.read()
                    all_rows, _ = read_sheet_rows(file_bytes, sheet_name="Data")

                    wb_out = Workbook()
                    ws_out = wb_out.active
                    ws_out.title = "SVNaturalLanguage"
                    ws_out.append([
                        "$Name", "$GSTTypeofSupply", "$BaseUnits", "$_HSNCode",
                        "$Parent", "$OpeningRate", "$_IntegratedTax",
                        "$OpeningBalance", "$_ClosingBalance", "$Description"
                    ])
                    ws_out.append([""] * 10)

                    converted = 0
                    for row in all_rows[1:]:
                        if len(row) < 18:
                            continue
                        item_name = row[1]
                        uom       = row[4]
                        hsn       = row[5]
                        category  = row[6]
                        tax       = row[17]

                        if not item_name or str(item_name).strip() == "":
                            continue

                        gst_type = "" if (tax is not None and str(tax).strip() == "") else "Goods"

                        if tax is None:
                            integrated_tax = "0"
                        elif str(tax).strip() == "":
                            integrated_tax = ""
                        else:
                            try:
                                integrated_tax = str(int(float(tax)))
                            except (ValueError, TypeError):
                                integrated_tax = str(tax)

                        if not hsn or str(hsn).strip() == "":
                            hsn_str = ""
                        else:
                            try:
                                hsn_str = str(int(float(hsn)))
                            except (ValueError, TypeError):
                                hsn_str = str(hsn).strip()

                        ws_out.append([
                            str(item_name).strip(), gst_type,
                            str(uom).strip() if uom else "",
                            hsn_str,
                            str(category).strip() if category else "",
                            "", integrated_tax, "", "", ""
                        ])
                        converted += 1

                    match = re.search(r'\(([^)]+)\)', uploaded_mine.name)
                    cn = match.group(1) if match else "Export"

                    buffer = io.BytesIO()
                    wb_out.save(buffer)
                    buffer.seek(0)

                    st.success(f"✅  **{converted} items** converted!")
                    st.info("💡  Fill in Opening Rate, Opening Balance, and Closing Balance from Tally.")
                    st.download_button(
                        "⬇️  Download Tally File",
                        data=buffer,
                        file_name=f"tranzact_item_({cn}).xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )
                except KeyError:
                    st.error("❌  Could not find the 'Data' sheet. Make sure you're uploading the correct file.")
                except Exception as e:
                    st.error(f"❌  Something went wrong: {e}")

# ─────────────────────────────────────────────────────────────────────────────
# TAB 3 : Manage templates
# ─────────────────────────────────────────────────────────────────────────────
with tab3:
    st.markdown("Templates are saved when you convert a new client format and click **💾 Save**.")

    templates = load_templates()

    if not templates:
        st.info("No templates saved yet. Convert a file in **Tab 1** and save the mapping.")
    else:
        for name, tmpl in list(templates.items()):
            with st.expander(f"📋  {name}"):
                m = tmpl.get("mapping", {})
                for t, c in m.items():
                    if c:
                        st.markdown(f"- **{t}** ← `{c}`")
                ec = tmpl.get("extra_cols", [])
                if ec:
                    st.markdown(f"- *Extra columns:* `{', '.join(ec)}`")
                fp = tmpl.get("fingerprint", [])
                if fp:
                    preview = ", ".join(fp[:5]) + ("…" if len(fp) > 5 else "")
                    st.caption(f"Identified by {len(fp)} columns: {preview}")
                if st.button(f"🗑  Delete  '{name}'", key=f"del_{name}"):
                    del templates[name]
                    save_templates(templates)
                    st.rerun()

# ─────────────────────────────────────────────────────────────────────────────
# TAB 4 : Network Master  →  Network_Add_(X).xlsx
# ─────────────────────────────────────────────────────────────────────────────
with tab4:
    st.markdown("Upload a client ledger / vendor file — auto-converts to your **Network Add** format.")
    st.markdown(
        "**Output has 3 sheets:** &nbsp; ✅ Ready to upload &nbsp;·&nbsp; 🔑 Have GSTIN &nbsp;·&nbsp; ✏️ Need to update manually",
        unsafe_allow_html=True
    )

    net_up = st.file_uploader("Upload client file", type=["xlsx", "xls"], key="net_up",
                               label_visibility="collapsed")

    if net_up:
        net_up.seek(0)
        net_bytes = net_up.read()
        net_fname = net_up.name

        if st.session_state.get("net_fname") != net_fname:
            try:
                sheets   = read_file(net_bytes)
                sname    = _tally_detect_sheet(sheets)
                rows     = sheets[sname]
                fmt, hidx = detect_network_format(rows)
                st.session_state.net_fname  = net_fname
                st.session_state.net_rows   = rows
                st.session_state.net_hidx   = hidx
                st.session_state.net_fmt    = fmt
                st.session_state.net_out    = None
            except Exception as e:
                st.error(f"❌  Could not read file: {e}")
                st.stop()

        rows  = st.session_state.net_rows
        hidx  = st.session_state.net_hidx
        fmt   = st.session_state.net_fmt

        st.success(f"✅  **{net_fname}** uploaded")
        fmt_label = {"tally": "Tally Export", "mshriy": "MSHRIY Format", "unknown": "Unknown"}
        st.caption(f"Detected format: **{fmt_label.get(fmt, fmt)}**")

        if fmt == "unknown":
            st.warning("⚠️  Could not detect file format. Please check the file.")
        else:
            if st.button("▶  Convert Now", type="primary",
                         use_container_width=True, key="net_go"):
                with st.spinner("Converting…"):
                    try:
                        if fmt == "tally":
                            parties = convert_tally_parties(rows, hidx)
                        else:
                            parties = convert_mshriy_parties(rows, hidx)

                        pindb   = load_pincode_db()
                        parties = apply_pincode_lookup(parties, pindb)
                        ready, have_gstin, manual = split_network_sheets(parties)

                        st.session_state.net_out      = make_network_xlsx(ready, have_gstin, manual)
                        st.session_state.net_out_name = net_filename(net_fname)
                        st.session_state.net_counts   = (len(ready), len(have_gstin), len(manual))
                    except Exception as e:
                        st.error(f"❌  Something went wrong: {e}")

        if st.session_state.get("net_out") and st.session_state.get("net_fname") == net_fname:
            r, g, m = st.session_state.net_counts
            c1, c2, c3 = st.columns(3)
            c1.metric("✅ Ready to upload", r)
            c2.metric("🔑 Have GSTIN", g)
            c3.metric("✏️ Need to update manually", m)

            st.info("🔴  Rows highlighted in red = non-standard party type (not Sundry Debtors/Creditors) — review with client.")

            st.download_button(
                "⬇️  Download Network Add File",
                data=st.session_state.net_out,
                file_name=st.session_state.net_out_name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
                key="net_dl"
            )

st.divider()
st.caption("TranZact · Item Master Converter · Auto-detects client formats · Saves templates for repeat clients")
