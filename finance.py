#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Family finance tracker
======================
Reads bank exports (Leumi, Mizrahi) and credit-card exports (Isracard, Max),
drops the monthly card debits on the bank side (the card files have the detail),
drops transfers between your own accounts, categorizes with your rules and the
memory of your past answers, and builds Excel reports.

  python finance.py import              read new export files, write output/review.xlsx
  python finance.py import --all        re-read every export file (known rows are skipped)
  python finance.py apply               read your answers from output/review.xlsx
  python finance.py report              build output/report.xlsx (all years)
  python finance.py report --year 2026  build output/report_2026.xlsx
  python finance.py status              short summary
"""
import argparse
import csv
import datetime as dt
import glob
import hashlib
import os
import re
import shutil
import sys
import warnings
from collections import Counter, defaultdict

warnings.filterwarnings("ignore")
try:
    sys.stdout.reconfigure(errors="replace")
except Exception:
    pass

try:
    import pandas as pd
    import openpyxl
    from openpyxl.chart import BarChart, Reference, Series
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.datavalidation import DataValidation
except ImportError as e:
    sys.exit("Missing package: %s\nRun:  pip install -r requirements.txt" % e.name)

# ------------------------------------------------------------------ settings
THRESHOLD = 0.95            # past answers must agree at least this often to auto-apply
TRANSFER_DAYS = 3           # max days between the two sides of an own-account transfer
CARD_DEBIT_WINDOW = (-2, 5) # bank debit may be this many days before/after the card billing date

BASE = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(BASE, "config")
DATA = os.path.join(BASE, "data")
OUTPUT = os.path.join(BASE, "output")
BACKUPS = os.path.join(BASE, "backups")
EXPORTS = os.path.join(BASE, "exports")
ENC = "utf-8-sig"           # UTF-8 with BOM so Excel shows Hebrew correctly

IGNORE = "התעלם"
ASK = "שאל תמיד"
SPLIT = "פיצול"
UNCAT = "לא מסווג"
SEP = " > "
INCOME, EXPENSE = "הכנסה", "הוצאה"

# How a monthly credit-card debit looks in the bank -> card company
CARD_DEBIT_PATTERNS = [
    ("מקס", ["מקס איט", "ויזה מקס", "מקס-איט", "max it"]),
    ("ישראכרט", ["מאסטרקארד", "ישראכרט", "אמריקן אקספרס"]),
    ("כאל", ["כרטיסי אשראי", "ויזה כאל", "כ.א.ל"]),
]
# A bank line must contain one of these to be paired as a transfer between your accounts
TRANSFER_HINTS = ["העבר", "הע.", "זיכוי", "בנק"]

TX_FIELDS = ["id", "source", "account", "kind", "date", "billing_date", "report_date",
             "amount", "description", "notes", "source_category", "ref", "installment",
             "orig_amount", "orig_currency", "card_company", "file",
             "status", "category", "subcategory", "reason"]
STATUS_HE = {"pending": "ממתין לסיווג", "auto": "אוטומטי", "user": "ידני",
             "card_debit": "חיוב כרטיס בבנק (לא נספר)",
             "transfer": "העברה בין החשבונות שלכם (לא נספר)"}
HEB_MONTHS = {"ינואר": 1, "פברואר": 2, "מרץ": 3, "מרס": 3, "אפריל": 4, "מאי": 5, "יוני": 6,
              "יולי": 7, "אוגוסט": 8, "ספטמבר": 9, "אוקטובר": 10, "נובמבר": 11, "דצמבר": 12}
BIDI = re.compile("[\u200e\u200f\u202a-\u202e]")
INST = re.compile(r"תשלום\s*(\d+)\s*מתוך\s*(\d+)")

# ------------------------------------------------------------- starter config
SEED_CATEGORIES = (
    [(EXPENSE, c, "") for c in ["חיסכון", "שכר דירה", "חשמל", "מים", "ארנונה", "סופר",
                                 "בגדים ונעליים", "ציוד לבית", "טיולים גדולים", "ילדים"]]
    + [(EXPENSE, "פלאפון אינטרנט ושאר אפליקציות", s)
       for s in ["פלא בן זוג 1", "פלא בן זוג 2", "אינטרנט", "ספוטיפיי", "בידור"]]
    + [(EXPENSE, c, "") for c in ["מתנות ותרומות", "שונות", "ביטוחים ורפואה", "הלוואות"]]
    + [(EXPENSE, "תחבורה", s)
       for s in ["אוטובוסים + תחבורה", "דלק", "תיקונים", "תשלומים",
                 "חניות וקנסות", "כביש 6", "שונות", "ביטוח רכב"]]
    + [(INCOME, "משכורת", "בן זוג 1"), (INCOME, "משכורת", "בן זוג 2"), (INCOME, "מילואים", ""),
       (INCOME, "ביטוח לאומי וקצבאות", ""), (INCOME, "ריבית והטבות בנק", ""),
       (INCOME, "הכנסות אחרות", "")]
)

def _r(text, cat, sub="", how="מכיל", acct="", amount=""):
    return (text, how, acct, amount, cat, sub)

SEED_RULES = (
    # investments / savings -> not counted
    [_r(t, IGNORE) for t in ['ני"ע', "איילון כספית", "מיטב טרייד", "קניה-אינטרנט", "מכירה-אינטרנט"]]
    # loans
    + [_r(t, "הלוואות") for t in ["הלוואה- פרעון", "טפחות - לווים", "פרעון הלוואה", "הטבת ריבית בהלוואה"]]
    # income
    + [_r(t, "ביטוח לאומי וקצבאות") for t in ["קצבת ילדים", "ביטוח לאומי", "בטוח לאומי"]]
    + [_r(t, "מילואים") for t in ['מופ"ת מילואי', "מענק מילואים"]]
    + [_r(t, "ריבית והטבות בנק") for t in ["רבית זכות", 'ריבית עו"ש', "מענק בנק לאומי", "הטבת ריבית"]]
    # expenses that are unambiguous
    + [_r("SPOTIFY", "פלאפון אינטרנט ושאר אפליקציות", "ספוטיפיי"),
       _r("חברת חשמל", "חשמל"), _r("ארנונה", "ארנונה")]
    + [_r(t, "סופר") for t in ["שופרסל", "שופר סל", "אושר עד", "רמי לוי", "חצי חינם",
                                "מחסני השוק", "יוחננוף", "ויקטורי"]]
    + [_r(t, "תחבורה", "אוטובוסים + תחבורה") for t in ["רב קו", "רב-קו", "מוביט"]]
    + [_r(t, "תחבורה", "חניות וקנסות") for t in ["פנגו", "חניון"]]
    + [_r(t, "תחבורה", "כביש 6") for t in ["כביש 6", "דרך ארץ"]]
    + [_r(t, "ביטוחים ורפואה") for t in ["מגדל חיים/בריאות", "מכבי שירותי", "שירותי בריאו"]]
    # cash withdrawals: not counted directly - they feed 'הוצאות מזומן נוספות'
    + [_r(t, IGNORE) for t in ["משיכת מזומן", "כספ.", "כספומט"]]
    # always ask: Bit, PayBox, checks, transfers to people
    + [_r("ביט", ASK, how="מילה")]
    + [_r(t, ASK) for t in ["PAYBOX", "פייבוקס", "פירעון שיק",
                             "העברת משכורת", "העברה", "הע. אינטרנט", "זיכוי מיידי", "זיכוי מידי"]]
    + [_r("הוראת קבע", ASK, how="מדויק"), _r("בנק", ASK, how="מילה")]
)

# Max's own category -> suggested category (only a suggestion, you still confirm it)
SEED_HINTS = [("מזון וצריכה", "סופר"), ("מסעדות, קפה וברים", "פנאי > אוכל"),
              ("פנאי, בידור וספורט", "פנאי > בילוי"), ("אופנה", "בגדים ונעליים"),
              ("רפואה ובתי מרקחת", "ביטוחים ורפואה"), ("ביטוח", "ביטוחים ורפואה"),
              ("עיצוב הבית", "ציוד לבית"), ("חשמל ומחשבים", "ציוד לבית"), ("שונות", "שונות"),
              ("תחבורה ורכבים", "תחבורה > אוטובוסים + תחבורה"), ("דלק, חשמל וגז", "תחבורה > דלק"),
              ("שירותי תקשורת", "פלאפון אינטרנט ושאר אפליקציות")]

# ------------------------------------------------------------------- helpers
def p(*parts):
    return os.path.join(*parts)

def stamp():
    return dt.datetime.now().strftime("%Y%m%d-%H%M%S")

def read_csv(path):
    if not os.path.exists(path):
        return []
    with open(path, encoding=ENC, newline="") as f:
        return [{k.strip(): (v or "").strip() for k, v in r.items() if k} for r in csv.DictReader(f)]

def write_csv(path, rows, fields):
    tmp = path + ".tmp"
    with open(tmp, "w", encoding=ENC, newline="") as f:
        w = csv.DictWriter(f, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in fields})
    try:
        os.replace(tmp, path)
    except PermissionError:
        sys.exit("Cannot write %s - is it open in Excel? Close it and run again." % os.path.basename(path))

def clean(v):
    if v is None:
        return ""
    if isinstance(v, float):
        if v != v:
            return ""
        if v.is_integer():
            v = int(v)
    return BIDI.sub("", str(v)).strip()

def num(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return 0.0 if v != v else float(v)
    s = BIDI.sub("", str(v)).replace(",", "").replace("₪", "").replace("$", "").strip()
    try:
        return float(s)
    except ValueError:
        return 0.0

def to_date(v):
    if v is None:
        return None
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    m = re.match(r"^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})$", clean(v))
    if not m:
        return None
    d, mo, y = (int(x) for x in m.groups())
    if y < 100:
        y += 2000
    try:
        return dt.date(y, mo, d)
    except ValueError:
        return None

def iso(d):
    return d.isoformat() if d else ""

def dget(t, field="report_date"):
    return dt.date.fromisoformat(t.get(field) or t["date"])

def amt(t):
    return float(t.get("amount") or 0)

def money(x):
    return "%.2f" % x

def find_col(header, *keys):
    for k in keys:
        for i, h in enumerate(header):
            if h == k:
                return i
    for k in keys:
        for i, h in enumerate(header):
            if k in h:
                return i
    return None

def cell(row, i):
    return row[i] if i is not None and i < len(row) else None

def installment(*texts):
    for s in texts:
        m = INST.search(s or "")
        if m:
            return "%s/%s" % m.groups()
    return ""

def tx(**kw):
    t = {f: "" for f in TX_FIELDS}
    for k, v in kw.items():
        t[k] = v if isinstance(v, str) else ("" if v is None else str(v))
    return t

def mkey(desc):
    """Merchant key: description without branch numbers, codes and punctuation."""
    s = clean(desc)[:20].lower().replace("(י)", " ")   # Isracard cuts names at 20 chars
    s = re.sub(r"-י$", "", s)
    s = " ".join(tok for tok in s.split() if not re.search(r"\d", tok))
    s = re.sub(r"[^\w\s]|_|\d", " ", s)
    return re.sub(r"\s+", " ", s).strip()

def counterparty(notes):
    m = re.search(r"(?:למי|מאת|ממי)\s*:\s*([^,|]+)", notes or "")
    return m.group(1).strip() if m else ""

def open_file(path):
    if sys.platform == "win32":
        try:
            os.startfile(path)
        except OSError:
            pass

# -------------------------------------------------------------- file reading
def df_rows(df):
    rows = []
    if not isinstance(df.columns, pd.RangeIndex):
        rows.append([c[-1] if isinstance(c, tuple) else c for c in df.columns])
    for r in df.itertuples(index=False, name=None):
        rows.append([None if (isinstance(v, float) and v != v) else v for v in r])
    return rows

def load_sheets(path):
    with open(path, "rb") as f:
        head = f.read(1024)
    if head[:2] == b"PK":                                   # .xlsx
        wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
        out = [(ws.title, [list(r) for r in ws.iter_rows(values_only=True)]) for ws in wb.worksheets]
        wb.close()
        return out
    if head.lstrip(b"\xef\xbb\xbf \t\r\n").startswith(b"<"):  # HTML saved as .xls (Leumi, Mizrahi)
        tables = pd.read_html(path, encoding="utf-8")
        return [("table%d" % i, df_rows(t)) for i, t in enumerate(tables)]
    book = pd.read_excel(path, sheet_name=None, header=None)  # old binary .xls
    return [(str(n), df_rows(df)) for n, df in book.items()]

def sniff(sheets):
    for _, rows in sheets:
        for r in rows[:40]:
            cells = {clean(c) for c in r}
            if "שם בית העסק" in cells and "תאריך חיוב" in cells:
                return "max"
            if "שם בית עסק" in cells and "סכום חיוב" in cells:
                return "isracard"
            if "אסמכתא" in cells and "תאריך" in cells:
                return "bank"
    return None

def parse_bank(sheets, fname):
    text = " ".join(clean(c) for _, rows in sheets for r in rows[:12] for c in r)
    m = re.search(r"חשבון\s*:\s*(\d[\d/\-]{3,}\d)", text)
    acct = m.group(1) if m else os.path.splitext(fname)[0]
    out = []
    for _, rows in sheets:
        hdr = None
        for r in rows:
            cells = [clean(c) for c in r]
            if hdr is None:
                if "אסמכתא" in cells and "תאריך" in cells:
                    hdr = cells
                    bank = "לאומי" if "בחובה" in cells else ("מזרחי" if "סוג תנועה" in cells else "בנק")
                    ci = dict(date=find_col(hdr, "תאריך"), desc=find_col(hdr, "תיאור", "סוג תנועה"),
                              ref=find_col(hdr, "אסמכתא"), debit=find_col(hdr, "בחובה", "חובה"),
                              credit=find_col(hdr, "בזכות", "זכות"), note=find_col(hdr, "הערה"))
                continue
            d = to_date(cell(r, ci["date"]))
            if not d:
                continue
            a = num(cell(r, ci["credit"])) - num(cell(r, ci["debit"]))
            if abs(a) < 0.005:
                continue
            out.append(tx(source=bank, account="%s %s" % (bank, acct), kind="bank",
                          date=iso(d), report_date=iso(d), amount=money(a),
                          description=clean(cell(r, ci["desc"])), notes=clean(cell(r, ci["note"])),
                          ref=clean(cell(r, ci["ref"])), file=fname))
    return out

def parse_isracard(sheets, fname):
    out = []
    fm = re.match(r"^(\d{4})_", fname)
    for _, rows in sheets:
        year = month = None
        card = fm.group(1) if fm else ""
        billing, h = None, None
        for r in rows:
            cells = [clean(c) for c in r]
            joined = " ".join(c for c in cells if c)
            if year is None:
                m = re.search(r"(%s)\s+(\d{4})" % "|".join(HEB_MONTHS), joined)
                if m:
                    month, year = HEB_MONTHS[m.group(1)], int(m.group(2))
            m = re.search(r"\s-\s*(\d{4})$", cells[0]) if cells else None
            if m and any("₪" in c for c in cells[1:]):
                card, billing = m.group(1), None
            m = re.search(r"לחיוב\s*ב-?\s*(\d{1,2})\.(\d{1,2})", joined)
            if m and year:
                d, mo = int(m.group(1)), int(m.group(2))
                y = year + (1 if month and mo < month - 6 else 0) - (1 if month and mo > month + 6 else 0)
                try:
                    billing = dt.date(y, mo, d)
                except ValueError:
                    pass
            if "שם בית עסק" in cells or "שם בית העסק" in cells:
                h = cells
                ci = dict(date=find_col(h, "תאריך רכישה", "תאריך עסקה"), bill=find_col(h, "תאריך חיוב"),
                          name=find_col(h, "שם בית עסק", "שם בית העסק"),
                          orig=find_col(h, "סכום עסקה", "סכום מקורי"), cur=find_col(h, "מטבע עסקה", "מטבע מקור"),
                          charge=find_col(h, "סכום חיוב"), ccur=find_col(h, "מטבע חיוב"),
                          voucher=find_col(h, "שובר"), extra=find_col(h, "פירוט נוסף"))
                continue
            if h is None:
                continue
            d = to_date(cell(r, ci["date"]))
            if not d:
                continue
            charge = num(cell(r, ci["charge"]))
            if abs(charge) < 0.005:
                continue
            bd = to_date(cell(r, ci["bill"])) or billing
            notes = re.sub(r"\s+", " ", clean(cell(r, ci["extra"])))
            ccur = clean(cell(r, ci["ccur"]))
            if ccur and ccur != "₪":
                notes = (notes + " | חיוב במטבע " + ccur).strip(" |")
            inst = installment(notes)
            orig = num(cell(r, ci["orig"]))
            out.append(tx(source="ישראכרט", account="ישראכרט " + card, kind="card", card_company="ישראכרט",
                          date=iso(d), billing_date=iso(bd), report_date=iso(bd if (inst and bd) else d),
                          amount=money(-charge), description=clean(cell(r, ci["name"])), notes=notes,
                          ref=clean(cell(r, ci["voucher"])), installment=inst,
                          orig_amount=money(orig) if orig else "", orig_currency=clean(cell(r, ci["cur"])),
                          file=fname))
    return out

def parse_max(sheets, fname):
    out = []
    for title, rows in sheets:
        if "טרם נקלטו" in title:        # approved but not yet charged - comes in a later export
            continue
        if "ארנק" in title:             # paid from the foreign-currency wallet; the shekel cost is the
            continue                     # wallet top-up ("רכישת מט\"ח") on the main sheet
        h, head_card = None, ""
        for r in rows:
            cells = [clean(c) for c in r]
            if h is None:
                m = re.match(r"^(\d{4})-\D", cells[0]) if cells else None
                if m and not head_card:
                    head_card = m.group(1)
                if "שם בית העסק" in cells:
                    h = cells
                    ci = dict(date=find_col(h, "תאריך עסקה"), name=find_col(h, "שם בית העסק"),
                              cat=find_col(h, "קטגוריה"), card=find_col(h, "4 ספרות"),
                              typ=find_col(h, "סוג עסקה"), charge=find_col(h, "סכום חיוב"),
                              orig=find_col(h, "סכום עסקה מקורי"), cur=find_col(h, "מטבע עסקה מקורי"),
                              bill=find_col(h, "תאריך חיוב"), notes=find_col(h, "הערות"))
                continue
            d = to_date(cell(r, ci["date"]))
            if not d:
                continue
            charge = num(cell(r, ci["charge"]))
            if abs(charge) < 0.005:
                continue
            typ = clean(cell(r, ci["typ"]))
            if "זיכוי" in typ and charge > 0:
                charge = -charge
            notes = clean(cell(r, ci["notes"]))
            if typ and typ not in ("רגילה",):
                notes = (notes + " | " + typ).strip(" |")
            inst = installment(notes)
            bd = to_date(cell(r, ci["bill"]))
            card = clean(cell(r, ci["card"])) or head_card
            orig = num(cell(r, ci["orig"]))
            out.append(tx(source="מקס", account="מקס " + card, kind="card", card_company="מקס",
                          date=iso(d), billing_date=iso(bd), report_date=iso(bd if (inst and bd) else d),
                          amount=money(-charge), description=clean(cell(r, ci["name"])), notes=notes,
                          source_category=clean(cell(r, ci["cat"])), installment=inst,
                          orig_amount=money(orig) if orig else "", orig_currency=clean(cell(r, ci["cur"])),
                          file=fname))
    return out

PARSERS = {"bank": parse_bank, "isracard": parse_isracard, "max": parse_max}

def parse_file(path):
    sheets = load_sheets(path)
    kind = sniff(sheets)
    if not kind:
        return None, []
    return kind, PARSERS[kind](sheets, os.path.basename(path))

def make_ids(rows):
    seen = Counter()
    for t in rows:
        base = "|".join([t["source"], t["account"], t["date"], t["billing_date"], t["amount"],
                         t["description"], t["ref"]])
        n = seen[base]
        seen[base] += 1
        t["id"] = hashlib.sha1(("%s|%d" % (base, n)).encode("utf-8")).hexdigest()[:12]

def export_files():
    files = set()
    for f in glob.glob(p(BASE, "*")) + glob.glob(p(EXPORTS, "**", "*"), recursive=True):
        name = os.path.basename(f)
        if (os.path.isfile(f) and name.lower().endswith((".xls", ".xlsx", ".html", ".htm"))
                and not name.startswith("~$")):
            files.add(f)
    return sorted(files)

# ------------------------------------------------------------ config & data
def setup():
    for d in (CONFIG, DATA, OUTPUT, BACKUPS, EXPORTS):
        os.makedirs(d, exist_ok=True)
    if not os.path.exists(p(CONFIG, "categories.csv")):
        write_csv(p(CONFIG, "categories.csv"),
                  [{"סוג": a, "קטגוריה": b, "תת-קטגוריה": c} for a, b, c in SEED_CATEGORIES],
                  ["סוג", "קטגוריה", "תת-קטגוריה"])
    if not os.path.exists(p(CONFIG, "card_hints.csv")):
        write_csv(p(CONFIG, "card_hints.csv"),
                  [{"קטגוריית חברת האשראי": a, "הצעה": b} for a, b in SEED_HINTS],
                  ["קטגוריית חברת האשראי", "הצעה"])
    if not os.path.exists(p(CONFIG, "balances.csv")):
        write_csv(p(CONFIG, "balances.csv"), [], BAL_FIELDS)
    if not os.path.exists(p(CONFIG, "notes.csv")):
        write_csv(p(CONFIG, "notes.csv"), [], NOTE_FIELDS)
    if not os.path.exists(p(CONFIG, "recurring.csv")):
        write_csv(p(CONFIG, "recurring.csv"), [], REC_FIELDS)
    if not os.path.exists(p(CONFIG, "rules.csv")):
        write_csv(p(CONFIG, "rules.csv"),
                  [dict(zip(RULE_FIELDS, r)) for r in SEED_RULES], RULE_FIELDS)

RULE_FIELDS = ["טקסט", "התאמה", "חשבון או כרטיס", "סכום", "קטגוריה", "תת-קטגוריה"]

def load_categories():
    cats, order = {}, []
    for r in read_csv(p(CONFIG, "categories.csv")):
        c, s = r.get("קטגוריה", ""), r.get("תת-קטגוריה", "")
        if not c:
            continue
        if c not in cats:
            cats[c] = {"type": INCOME if "הכנס" in r.get("סוג", "") else EXPENSE, "subs": []}
            order.append(c)
        if s and s not in cats[c]["subs"]:
            cats[c]["subs"].append(s)
    if not cats:
        sys.exit("config/categories.csv is empty.")
    return cats, order

def load_rules(cats):
    rules = []
    letters = "A-Za-z\u05d0-\u05ea"
    for i, r in enumerate(read_csv(p(CONFIG, "rules.csv")), 2):
        text, cat, sub = r.get("טקסט", ""), r.get("קטגוריה", ""), r.get("תת-קטגוריה", "")
        if not text or not cat:
            continue
        if cat not in (IGNORE, ASK) and (cat not in cats or any(x and x not in cats[cat]["subs"] for x in sub.split("&"))):
            print("  ! rules.csv line %d: category '%s%s' is not in categories.csv - rule skipped"
                  % (i, cat, SEP + sub if sub else ""))
            continue
        how = r.get("התאמה") or "מכיל"
        pre = "(?<![%s])" % letters if re.match("[%s]" % letters, text[0]) else ""
        post = "(?![%s])" % letters if re.match("[%s]" % letters, text[-1]) else ""
        try:
            amount = abs(float(r.get("סכום", "").replace(",", ""))) if r.get("סכום", "") else None
        except ValueError:
            amount = None
        rules.append(dict(text=text, how=how, acct=r.get("חשבון או כרטיס", ""), amount=amount, cat=cat, sub=sub,
                          rx=re.compile(pre + re.escape(text) + post, re.I)))
    return rules

def rule_matches(r, t):
    if r["acct"] and r["acct"] not in t["account"]:
        return False
    if r.get("amount") is not None and abs(abs(amt(t)) - r["amount"]) > 0.01:
        return False
    desc = t["description"]
    if r["how"] == "מדויק":
        return desc.strip() == r["text"]
    if r["how"] == "מילה":
        return r["rx"].search(desc) is not None
    return r["text"].lower() in desc.lower()

def load_tx():
    return read_csv(p(DATA, "transactions.csv"))

def save_tx(txns):
    txns.sort(key=lambda t: (t["date"], t["account"], t["id"]))
    write_csv(p(DATA, "transactions.csv"), txns, TX_FIELDS)

def load_splits():
    out = defaultdict(list)
    for r in read_csv(p(DATA, "splits.csv")):
        out[r["id"]].append(r)
    return out

def save_splits(splits):
    write_csv(p(DATA, "splits.csv"), [r for rows in splits.values() for r in rows],
              ["id", "amount", "category", "subcategory"])

def get_state(key):
    return {r["key"]: r["value"] for r in read_csv(p(DATA, "state.csv"))}.get(key, "")

def set_state(key, value):
    st = {r["key"]: r["value"] for r in read_csv(p(DATA, "state.csv"))}
    st[key] = value
    write_csv(p(DATA, "state.csv"), [{"key": k, "value": v} for k, v in st.items()], ["key", "value"])

ACC_FIELDS = ["חשבון / כרטיס", "סוג", "כינוי"]

def update_accounts(txns):
    path = p(CONFIG, "accounts.csv")
    rows = read_csv(path)
    known = {r.get("חשבון / כרטיס") for r in rows}
    for acc, kind in sorted({(t["account"], t["kind"]) for t in txns}):
        if acc not in known:
            rows.append({"חשבון / כרטיס": acc, "סוג": "בנק" if kind == "bank" else "כרטיס אשראי", "כינוי": ""})
    write_csv(path, rows, ACC_FIELDS)
    return {r["חשבון / כרטיס"]: r.get("כינוי", "") for r in rows}

def load_accounts():
    return {r.get("חשבון / כרטיס", ""): r.get("כינוי", "") for r in read_csv(p(CONFIG, "accounts.csv"))}

def acct_label(acc, accounts):
    nick = accounts.get(acc, "")
    return "%s (%s)" % (acc, nick) if nick else acc

def backup():
    files = glob.glob(p(DATA, "*.csv"))
    if not files:
        return
    dest = p(BACKUPS, stamp())
    os.makedirs(dest, exist_ok=True)
    for f in files:
        shutil.copy2(f, dest)
    olds = sorted(d for d in glob.glob(p(BACKUPS, "2*")) if os.path.isdir(d))
    for d in olds[:-30]:
        shutil.rmtree(d, ignore_errors=True)

# ------------------------------------------------------------ categorization
def card_company_of(desc):
    low = desc.lower()
    for comp, pats in CARD_DEBIT_PATTERNS:
        if any(x.lower() in low for x in pats):
            return comp
    return ""

def pair_transfers(txns):
    cand = [t for t in txns if t["status"] == "pending" and t["kind"] == "bank"
            and any(h in t["description"] for h in TRANSFER_HINTS)]
    pos = [t for t in cand if amt(t) > 0]
    used = set()
    for n in sorted((t for t in cand if amt(t) < 0), key=lambda t: t["date"]):
        best = None
        for q in pos:
            if q["id"] in used or q["account"] == n["account"] or abs(amt(q) + amt(n)) > 0.005:
                continue
            gap = abs((dget(q, "date") - dget(n, "date")).days)
            if gap <= TRANSFER_DAYS and (best is None or gap < best[0]):
                best = (gap, q)
        if best:
            q = best[1]
            used.add(q["id"])
            n.update(status="transfer", reason="העברה ל" + q["account"])
            q.update(status="transfer", reason="העברה מ" + n["account"])

def load_hints(cats):
    out = {}
    for r in read_csv(p(CONFIG, "card_hints.csv")):
        res = parse_label(r.get("הצעה", ""), cats)
        if res:
            out[r.get("קטגוריית חברת האשראי", "")] = res
    return out

def classify(txns, rules):
    """Re-runs on everything you didn't answer yourself, so rule edits take effect."""
    try:
        hints = load_hints(load_categories()[0])
    except SystemExit:
        hints = {}
    for t in txns:
        rule = next((r for r in rules if rule_matches(r, t)), None)
        t["_rule"] = rule
        t["_ask"] = bool(rule and rule["cat"] == ASK)
        key = mkey(t["description"])
        if t["_ask"]:
            key += "|" + (counterparty(t["notes"]) or "%.0f" % abs(amt(t)))
        t["_key"] = key
        t["_suggest"] = None
    memory = defaultdict(Counter)
    for t in txns:
        if t["status"] == "user" and t["category"] != SPLIT:
            memory[t["_key"]][(t["category"], t["subcategory"])] += 1
    for t in txns:
        if t["status"] != "user":
            t.update(status="pending", category="", subcategory="", reason="")
            if t["kind"] == "bank":
                t["card_company"] = ""
    for t in txns:
        if t["status"] == "pending" and t["kind"] == "bank":
            comp = card_company_of(t["description"])
            if comp:
                t.update(status="card_debit", card_company=comp,
                         reason="חיוב כרטיס %s - הפירוט נלקח מקובץ הכרטיס" % comp)
    pair_transfers(txns)
    for t in txns:
        if t["status"] != "pending":
            continue
        r = t["_rule"]
        if r and r["cat"] != ASK:
            t.update(status="auto", category=r["cat"], subcategory=r["sub"], reason="כלל: " + r["text"])
            continue
        why = ("תמיד שואלים (%s)" % r["text"]) if t["_ask"] else ""
        m = memory.get(t["_key"])
        if m:
            (c, s), n = m.most_common(1)[0]
            total = sum(m.values())
            share = n / total
            if not t["_ask"] and share >= THRESHOLD:
                t.update(status="auto", category=c, subcategory=s,
                         reason="לפי תשובות קודמות (%d מתוך %d)" % (n, total))
            else:
                t["_suggest"] = (c, s)
                t["reason"] = why or "תשובות קודמות לא חד-משמעיות (%.0f%%)" % (share * 100)
        else:
            t["reason"] = why or "בית עסק חדש"
            hint = hints.get(t["source_category"])
            if hint and not t["_ask"]:
                t["_suggest"] = hint
    return memory

def save_memory(memory):
    rows = [{"key": k, "category": c, "subcategory": s, "count": n}
            for k, cnt in sorted(memory.items()) for (c, s), n in cnt.most_common()]
    write_csv(p(DATA, "memory.csv"), rows, ["key", "category", "subcategory", "count"])

# ------------------------------------------------------- fixed monthly items
REC_FIELDS = ["תיאור", "סכום", "קטגוריה", "תת-קטגוריה", "מחודש", "עד חודש", "שולם במזומן"]

def recurring_tx(cats):
    """Monthly amounts you pay outside the bank files (config/recurring.csv), up to this month."""
    out = []
    today = dt.date.today()
    for i, r in enumerate(read_csv(p(CONFIG, "recurring.csv")), 2):
        c, s = r.get("קטגוריה", ""), r.get("תת-קטגוריה", "")
        if c not in cats or (s and s not in cats[c]["subs"]):
            print("  ! recurring.csv line %d: unknown category '%s' - skipped" % (i, cat_label(c, s)))
            continue
        try:
            y, m = (int(x) for x in r["מחודש"].split("-"))
            end = r.get("עד חודש") or today.strftime("%Y-%m")
            ey, em = (int(x) for x in end.split("-"))
            value = abs(float(r["סכום"].replace(",", "")))
        except (ValueError, KeyError):
            print("  ! recurring.csv line %d: months must look like 2025-01 and amount a number - skipped" % i)
            continue
        sign = 1 if cats[c]["type"] == INCOME else -1
        while (y, m) <= (ey, em) and (y, m) <= (today.year, today.month):
            d = dt.date(y, m, 1).isoformat()
            t = tx(source="קבוע", account="קבוע (ידני)", kind="manual", date=d, report_date=d,
                   amount=money(sign * value), description=r.get("תיאור", ""), status="user",
                   category=c, subcategory=s, reason="תשלום קבוע (recurring.csv)")
            t["id"] = hashlib.sha1(("rec|%s|%s|%s" % (r.get("תיאור", ""), cat_label(c, s), d)).encode()).hexdigest()[:12]
            t["_cash"] = r.get("שולם במזומן", "כן") != "לא"
            out.append(t)
            m += 1
            if m > 12:
                y, m = y + 1, 1
    return out

# ------------------------------------------------------- cash ("הוצאות מזומן נוספות")
CASH_CAT = "הוצאות מזומן נוספות"
CASH_RX = re.compile(r"משיכת מזומן|כספ\.|כספומט")

def is_cash(t):
    return (CASH_RX.search(t["description"]) is not None and amt(t) < 0
            and t["status"] not in ("transfer", "card_debit") and t["category"] in (IGNORE, ""))

def cash_tx(txns, cats):
    """Per month: cash withdrawn minus the fixed items paid in cash (recurring.csv) = other cash spending."""
    if CASH_CAT not in cats:
        return []
    out_m = defaultdict(float)
    for t in txns:
        if t.get("kind") != "manual" and is_cash(t):
            out_m[t["date"][:7]] += -amt(t)
    paid = defaultdict(float)
    for t in txns:
        if t.get("kind") == "manual" and t.get("_cash"):
            paid[t["date"][:7]] += -amt(t)
    out = []
    for m in sorted(set(out_m) | set(paid)):
        v = out_m.get(m, 0.0) - paid.get(m, 0.0)
        if abs(v) < 0.005:
            continue
        d = m + "-01"
        t = tx(source="מזומן", account="מזומן (מחושב)", kind="manual", date=d, report_date=d,
               amount=money(-v), description="משיכות %s פחות תשלומים קבועים במזומן %s"
               % ("{:,.0f}".format(out_m.get(m, 0)), "{:,.0f}".format(paid.get(m, 0))),
               status="user", category=CASH_CAT, reason="חישוב מזומן")
        t["id"] = hashlib.sha1(("cash|" + m).encode()).hexdigest()[:12]
        out.append(t)
    return out

# --------------------------------------------------------------- excel bits
BOLD = Font(bold=True)
HFILL = PatternFill("solid", fgColor="1F4E78")
HFONT = Font(bold=True, color="FFFFFF")
SECFILL = PatternFill("solid", fgColor="D9E1F2")
CATFILL = PatternFill("solid", fgColor="F2F2F2")
TOTFILL = PatternFill("solid", fgColor="FFF2CC")
YELLOW = PatternFill("solid", fgColor="FFFF99")
NUM = '#,##0;[Red]-#,##0;""'
NUM2 = '#,##0.00;[Red]-#,##0.00'
DATEF = "dd/mm/yyyy"

def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        x = ws.cell(row, c)
        x.fill, x.font = HFILL, HFONT
        x.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def new_sheet(wb, title):
    ws = wb.create_sheet(title)
    ws.sheet_view.rightToLeft = True
    return ws

def widths(ws, w):
    for i, x in enumerate(w, 1):
        ws.column_dimensions[get_column_letter(i)].width = x

def save_wb(wb, path):
    try:
        wb.save(path)
    except PermissionError:
        sys.exit("Cannot write %s - close it in Excel and run again." % os.path.basename(path))

def cat_label(c, s):
    return c + SEP + s if s else c

def all_labels(cats, order):
    out = []
    for c in order:
        out.append(c)
        out += [c + SEP + s for s in cats[c]["subs"]]
    return out + [IGNORE]

def parse_label(lbl, cats):
    lbl = clean(lbl)
    if lbl == IGNORE:
        return IGNORE, ""
    c, s = ([x.strip() for x in lbl.split(">", 1)] + [""])[:2] if ">" in lbl else (lbl, "")
    if c not in cats or (s and s not in cats[c]["subs"]):
        return None
    return c, s

# --------------------------------------------------------------- review file
REVIEW = p(OUTPUT, "review.xlsx")

def write_review(txns, cats, order, accounts):
    pend = [t for t in txns if t["status"] == "pending"]
    if os.path.exists(REVIEW):
        try:
            shutil.copy2(REVIEW, p(BACKUPS, "review_%s.xlsx" % stamp()))
        except OSError:
            pass
    if not pend:
        try:
            os.remove(REVIEW)
        except OSError:
            pass
        set_state("review_mtime", "")
        return 0
    freq = Counter(t["_key"] for t in pend)
    pend.sort(key=lambda t: (-freq[t["_key"]], t["_key"], t["date"]))
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "לסיווג"
    ws.sheet_view.rightToLeft = True
    head = ["מזהה", "תאריך", "חשבון / כרטיס", "תיאור", "פרטים נוספים", "סכום", "כמה כאלה ממתינות",
            "למה נשאל", "קטגוריה"]
    ws.append(head)
    style_header(ws, 1, len(head))
    for t in pend:
        sug = t.get("_suggest")
        details = " | ".join(x for x in [t["notes"], t["source_category"],
                                          ("תשלום " + t["installment"]) if t["installment"] else ""] if x)
        ws.append([t["id"], dget(t, "date"), acct_label(t["account"], accounts), t["description"], details,
                   amt(t), freq[t["_key"]], t["reason"], cat_label(*sug) if sug else ""])
        r = ws.max_row
        ws.cell(r, 2).number_format = DATEF
        ws.cell(r, 6).number_format = NUM2
        if sug:
            ws.cell(r, 9).fill = YELLOW
    labels = all_labels(cats, order)
    lists = wb.create_sheet("lists")
    for i, l in enumerate(labels, 1):
        lists.cell(i, 1, l)
    lists.sheet_state = "hidden"
    dv = DataValidation(type="list", formula1="=lists!$A$1:$A$%d" % len(labels), allow_blank=True,
                        showErrorMessage=True, errorTitle="קטגוריה לא מוכרת",
                        error="בחרו קטגוריה מהרשימה (או הוסיפו אותה ל-categories.csv)")
    ws.add_data_validation(dv)
    dv.add("I2:I%d" % (len(pend) + 500))
    widths(ws, [14, 12, 24, 34, 40, 11, 10, 30, 40])
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:I%d" % ws.max_row
    info = new_sheet(wb, "הוראות")
    for line in [
        "איך ממלאים",
        "1. בעמודה 'קטגוריה' בוחרים מהרשימה.",
        "2. תאים צהובים כבר מולאו בהצעה לפי תשובות קודמות - בדקו ושנו אם צריך.",
        "3. מספיק לענות על שורה אחת מכל בית עסק: אחרי apply, שאר השורות של אותו בית עסק יסווגו לבד.",
        "   (ביט, פייבוקס, מזומן, שיקים והעברות נשאלים תמיד, כל אחד בנפרד.)",
        "4. 'התעלם' - התנועה לא תיספר (למשל העברה לחשבון שלכם או השקעה).",
        "5. פיצול לכמה קטגוריות: העתיקו את השורה (אותו מזהה), ושנו את הסכומים כך שיסתכמו לסכום המקורי.",
        "6. שורה שנשארת ריקה תופיע שוב בפעם הבאה. אפשר לענות בכמה סבבים.",
        "7. שמרו, סגרו את הקובץ, והריצו apply (אפשרות 2 ב-run.bat).",
    ]:
        info.append([line])
    info["A1"].font = Font(bold=True, size=14)
    widths(info, [110])
    save_wb(wb, REVIEW)
    set_state("review_mtime", str(os.path.getmtime(REVIEW)))
    return len(pend)

# ---------------------------------------------------------------- shops file
SHOPS = p(OUTPUT, "shops.xlsx")

def write_shops(txns, cats, order, accounts):
    """One row per shop that is still waiting (Bit/PayBox/cash/transfers stay in review.xlsx)."""
    if os.path.exists(SHOPS):
        try:
            shutil.copy2(SHOPS, p(BACKUPS, "shops_%s.xlsx" % stamp()))
        except OSError:
            pass
    groups = defaultdict(list)
    for t in txns:
        if t["status"] == "pending" and not t["_ask"]:
            groups[t["_key"]].append(t)
    if not groups:
        try:
            os.remove(SHOPS)
        except OSError:
            pass
        set_state("shops_mtime", "")
        return 0
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "בתי עסק"
    ws.sheet_view.rightToLeft = True
    head = ["מפתח", "תיאור", "כמה פעמים", 'סה"כ סכום', "מתאריך", "עד תאריך", "חשבונות / כרטיסים",
            "קטגוריית חברת האשראי", "פרטים נוספים", "קטגוריה"]
    ws.append(head)
    style_header(ws, 1, len(head))
    for key, ts in sorted(groups.items(), key=lambda kv: (-len(kv[1]), kv[0])):
        descs = Counter(t["description"] for t in ts)
        dates = sorted(t["date"] for t in ts)
        sug = next((t["_suggest"] for t in ts if t["_suggest"]), None)
        ws.append([key, " / ".join(d for d, _ in descs.most_common(3)), len(ts), sum(amt(t) for t in ts),
                   dt.date.fromisoformat(dates[0]), dt.date.fromisoformat(dates[-1]),
                   ", ".join(sorted({acct_label(t["account"], accounts) for t in ts})),
                   ", ".join(sorted({t["source_category"] for t in ts if t["source_category"]})),
                   " | ".join(sorted({t["notes"] for t in ts if t["notes"]}))[:200],
                   cat_label(*sug) if sug else ""])
        r = ws.max_row
        ws.cell(r, 4).number_format = NUM2
        ws.cell(r, 5).number_format = ws.cell(r, 6).number_format = DATEF
        if sug:
            ws.cell(r, 10).fill = YELLOW
    labels = all_labels(cats, order)
    lists = wb.create_sheet("lists")
    for i, l in enumerate(labels, 1):
        lists.cell(i, 1, l)
    lists.sheet_state = "hidden"
    dv = DataValidation(type="list", formula1="=lists!$A$1:$A$%d" % len(labels), allow_blank=True,
                        showErrorMessage=True, errorTitle="קטגוריה לא מוכרת",
                        error="בחרו קטגוריה מהרשימה (או הוסיפו אותה ל-categories.csv)")
    ws.add_data_validation(dv)
    dv.add("J2:J%d" % (ws.max_row + 50))
    ws.column_dimensions["A"].hidden = True
    widths(ws, [10, 40, 10, 12, 12, 12, 30, 22, 40, 40])
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = "A1:J%d" % ws.max_row
    save_wb(wb, SHOPS)
    set_state("shops_mtime", str(os.path.getmtime(SHOPS)))
    return len(groups)

def shops_csv_files():
    return sorted(f for f in glob.glob(p(BASE, "shops*.csv")) + glob.glob(p(OUTPUT, "shops*.csv")))

def apply_shops(txns, cats, errors):
    """Answers from shops.xlsx (or shops*.csv) apply to every waiting transaction of that shop."""
    sheets = []
    if os.path.exists(SHOPS):
        try:
            wb = openpyxl.load_workbook(SHOPS, data_only=True)
        except PermissionError:
            sys.exit("Close shops.xlsx in Excel and run again.")
        sheets.append(list(wb.worksheets[0].iter_rows(min_row=2, values_only=True)))
    for f in shops_csv_files():
        with open(f, encoding=ENC, newline="") as fh:
            rows = list(csv.reader(fh))
        sheets.append(rows[1:])
        dest = p(BACKUPS, "applied_%s_%s" % (stamp(), os.path.basename(f)))
        try:
            shutil.move(f, dest)
        except OSError:
            pass
    answers = {}
    for r in (r for rows in sheets for r in rows):
        if not r or not r[0] or not clean(cell(r, 9)):
            continue
        res = parse_label(cell(r, 9), cats)
        if not res:
            errors.append("shops.xlsx %s: unknown category '%s'" % (clean(r[1]), clean(cell(r, 9))))
            continue
        answers[clean(r[0])] = res
    n = 0
    for t in txns:
        if t["status"] == "pending" and not t["_ask"] and t["_key"] in answers:
            c, s = answers[t["_key"]]
            t.update(status="user", category=c, subcategory=s, reason="תשובה שלך (קובץ בתי עסק)")
            n += 1
    return n

def summary(txns):
    c = Counter(t["status"] for t in txns)
    print("\nTransactions: %d | auto: %d | answered by you: %d | waiting for you: %d | "
          "card debits skipped: %d | own transfers skipped: %d"
          % (len(txns), c["auto"], c["user"], c["pending"], c["card_debit"], c["transfer"]))

def finish(txns, cats, order, rules, open_review=True):
    memory = classify(txns, rules)
    save_tx(txns)
    save_memory(memory)
    accounts = update_accounts(txns)
    n = write_review(txns, cats, order, accounts)
    ns = write_shops(txns, cats, order, accounts)
    summary(txns)
    if ns:
        print("%d shops waiting -> output\\shops.xlsx (one answer covers all rows of that shop)" % ns)
    if n:
        print("%d transactions need your answer -> output\\review.xlsx" % n)
        if open_review:
            open_file(REVIEW)
    else:
        print("Nothing waiting for you.")

# ------------------------------------------------------------------ commands
def cmd_import(args):
    setup()
    for path, key in ((REVIEW, "review_mtime"), (SHOPS, "shops_mtime")):
        saved = get_state(key)
        if os.path.exists(path) and not args.force and saved \
                and abs(os.path.getmtime(path) - float(saved)) > 2:
            sys.exit("output\\%s has answers that were not applied yet.\n"
                     "Run 'apply' first (option 2), or 'import --force' to discard them." % os.path.basename(path))
    cats, order = load_categories()
    rules = load_rules(cats)
    txns = load_tx()
    ids = {t["id"] for t in txns}
    log = read_csv(p(DATA, "imports.csv"))
    seen = {(r["file"], r["size"], r["mtime"]) for r in log}
    files = export_files()
    if not files:
        print("No export files found. Put bank/card exports in this folder or in: " + EXPORTS)
    for f in files:
        name = os.path.relpath(f, BASE)
        st = os.stat(f)
        sig = (name, str(st.st_size), str(int(st.st_mtime)))
        if sig in seen and not args.all:
            continue
        try:
            kind, rows = parse_file(f)
        except Exception as e:
            print("  ! %s: could not read (%s)" % (name, e))
            continue
        if kind is None:
            print("  ? %s: not a recognized bank/card export - skipped" % name)
            continue
        make_ids(rows)
        new = [t for t in rows if t["id"] not in ids]
        ids.update(t["id"] for t in new)
        txns.extend(new)
        print("  %-50s %-8s %4d rows, %4d new" % (name, kind, len(rows), len(new)))
        if sig not in seen:
            log.append(dict(file=name, size=sig[1], mtime=sig[2], kind=kind, rows=len(rows), new=len(new),
                            imported=dt.datetime.now().strftime("%Y-%m-%d %H:%M")))
    backup()
    write_csv(p(DATA, "imports.csv"), log, ["file", "size", "mtime", "kind", "rows", "new", "imported"])
    finish(txns, cats, order, rules, not args.no_open)

def cmd_apply(args):
    setup()
    if not os.path.exists(REVIEW) and not os.path.exists(SHOPS) and not shops_csv_files():
        print("Nothing to apply - output\\review.xlsx / shops.xlsx not found.")
        return
    cats, order = load_categories()
    rules = load_rules(cats)
    txns = load_tx()
    classify(txns, rules)
    by_id = {t["id"]: t for t in txns}
    splits = load_splits()
    errors = []
    nshop = apply_shops(txns, cats, errors)
    rows_iter = []
    if os.path.exists(REVIEW):
        try:
            wb = openpyxl.load_workbook(REVIEW, data_only=True)
        except PermissionError:
            sys.exit("Close review.xlsx in Excel and run again.")
        ws = wb["לסיווג"] if "לסיווג" in wb.sheetnames else wb.worksheets[0]
        rows_iter = ws.iter_rows(min_row=2, values_only=True)
    groups, seq = defaultdict(list), []
    for r in rows_iter:
        if not r or not r[0]:
            continue
        tid = clean(r[0])
        if tid not in groups:
            seq.append(tid)
        groups[tid].append((num(cell(r, 5)), clean(cell(r, 8))))
    done = nsplit = 0
    for tid in seq:
        t, items = by_id.get(tid), groups[tid]
        if not t or t["status"] != "pending":
            continue
        name = "%s %s %.2f" % (t["date"], t["description"], amt(t))
        if len(items) == 1:
            if not items[0][1]:
                continue
            res = parse_label(items[0][1], cats)
            if not res:
                errors.append("%s: unknown category '%s'" % (name, items[0][1]))
                continue
            t.update(status="user", category=res[0], subcategory=res[1], reason="תשובה שלך")
            done += 1
            continue
        if not all(l for _, l in items):
            if any(l for _, l in items):
                errors.append("%s: in a split every row needs a category" % name)
            continue
        parts = [(abs(a), parse_label(l, cats), l) for a, l in items]
        bad = [l for _, res, l in parts if not res]
        if bad:
            errors.append("%s: unknown category %s" % (name, ", ".join(bad)))
            continue
        total = sum(a for a, _, _ in parts)
        if abs(total - abs(amt(t))) > 0.01:
            errors.append("%s: split amounts add up to %.2f instead of %.2f" % (name, total, abs(amt(t))))
            continue
        sign = 1 if amt(t) > 0 else -1
        splits[tid] = [dict(id=tid, amount=money(sign * a), category=res[0], subcategory=res[1])
                       for a, res, _ in parts]
        t.update(status="user", category=SPLIT, subcategory="",
                 reason="פיצול: " + "; ".join("%.2f %s" % (a, cat_label(*res)) for a, res, _ in parts))
        nsplit += 1
    backup()
    save_splits(splits)
    print("Applied %d shop answers (shops.xlsx), %d answers and %d splits (review.xlsx)." % (nshop, done, nsplit))
    set_state("review_mtime", "")
    set_state("shops_mtime", "")
    finish(txns, cats, order, rules, not args.no_open)
    if errors:
        print("\nNot applied (still waiting - fix and apply again):")
        for e in errors:
            print("  - " + e)

def ctype(cats, c, a):
    return cats[c]["type"] if c in cats else (INCOME if a > 0 else EXPENSE)

SALARY_SHIFT_DAY = 25   # a salary paid on the 25th or later counts in the next month (as in your file)

def shift_salary(t, c):
    if c != "משכורת" or "לאומי" in t["description"]:
        return t
    d = dget(t)
    if d.day < SALARY_SHIFT_DAY:
        return t
    nd = (d.replace(day=1) + dt.timedelta(days=32)).replace(day=1)
    t2 = dict(t)
    t2["report_date"] = nd.isoformat()
    return t2

def report_lines(txns, splits):
    out = []
    for t in txns:
        if t["status"] in ("card_debit", "transfer") or t["category"] == IGNORE:
            continue
        if t["category"] == SPLIT and splits.get(t["id"]):
            out += [(t, float(s["amount"]), s["category"], s["subcategory"])
                    for s in splits[t["id"]] if s["category"] != IGNORE]
        elif t["status"] == "pending" or not t["category"]:
            out.append((t, amt(t), UNCAT, ""))
        elif "&" in t["subcategory"]:              # e.g. "פלא בן זוג 1&פלא בן זוג 2" = half each
            parts = [x for x in t["subcategory"].split("&") if x]
            out += [(t, amt(t) / len(parts), t["category"], x) for x in parts]
        else:
            out.append((t, amt(t), t["category"], t["subcategory"]))
    return [(shift_salary(t, c), a, c, s) for t, a, c, s in out]

def build_plan(cats, order, present):
    plan = {INCOME: [], EXPENSE: []}
    for ty in plan:
        extra = sorted({c for (t2, c, s) in present if t2 == ty and c not in cats and c != UNCAT})
        if any(t2 == ty and c == UNCAT for (t2, c, s) in present):
            extra.append(UNCAT)
        for c in [c for c in order if cats[c]["type"] == ty] + extra:
            subs = list(cats[c]["subs"]) if c in cats else []
            subs += sorted({s for (t2, c2, s) in present if t2 == ty and c2 == c and s and s not in subs})
            if subs and (ty, c, "") in present:
                subs.append("")
            plan[ty].append((c, subs))
    return plan

def write_matrix(ws, title, keys, titles, agg, plan, extra=(), show_total=True):
    ws.cell(1, 1, title).font = Font(bold=True, size=14)
    hr = 3
    head = ["קטגוריה", "תת-קטגוריה"] + list(titles) + (['סה"כ'] if show_total else []) + [e[0] for e in extra]
    for j, h in enumerate(head, 1):
        ws.cell(hr, j, h)
    style_header(ws, hr, len(head))
    ncol = len(head)
    pos = [hr + 1]

    def put(l1, l2, vals, fill=None, bold=False):
        r = pos[0]
        ws.cell(r, 1, l1)
        ws.cell(r, 2, l2)
        for j, k in enumerate(keys):
            ws.cell(r, 3 + j, round(vals.get(k, 0.0), 2)).number_format = NUM
        col = 3 + len(keys)
        if show_total:
            ws.cell(r, col, round(sum(vals.get(k, 0.0) for k in keys), 2)).number_format = NUM
            col += 1
        for _, fn in extra:
            ws.cell(r, col, round(fn(vals), 2)).number_format = NUM
            col += 1
        for j in range(1, ncol + 1):
            if fill:
                ws.cell(r, j).fill = fill
            if bold:
                ws.cell(r, j).font = BOLD
        pos[0] += 1
        return r

    sums = {}
    for ty, name in ((INCOME, "הכנסות"), (EXPENSE, "הוצאות")):
        ws.cell(pos[0], 1, name).font = Font(bold=True, size=12)
        for j in range(1, ncol + 1):
            ws.cell(pos[0], j).fill = SECFILL
        pos[0] += 1
        sec = defaultdict(float)
        for c, subs in plan[ty]:
            cv = defaultdict(float)
            for (t2, c2, s), v in agg.items():
                if t2 == ty and c2 == c:
                    for k, x in v.items():
                        cv[k] += x
            put(c, "", cv, CATFILL, True)
            for s in subs:
                put("", s or "(ללא תת-קטגוריה)", agg.get((ty, c, s), {}))
            for k, x in cv.items():
                sec[k] += x
        sums[ty] = (put('סה"כ ' + name, "", sec, TOTFILL, True), sec)
        pos[0] += 1
    net = {k: sums[INCOME][1].get(k, 0) - sums[EXPENSE][1].get(k, 0)
           for k in set(sums[INCOME][1]) | set(sums[EXPENSE][1])}
    net_row = put("נטו (הכנסות פחות הוצאות)", "", net, TOTFILL, True)
    ws.freeze_panes = ws.cell(hr + 1, 3)
    widths(ws, [32, 24] + [12] * (ncol - 2))
    return hr, sums[INCOME][0], sums[EXPENSE][0], net_row

def add_chart(ws, hr, inc_row, exp_row, nkeys, at_row, title):
    ch = BarChart()
    ch.type, ch.title, ch.height, ch.width = "col", title, 8, 26
    for row, name in ((inc_row, "הכנסות"), (exp_row, "הוצאות")):
        ch.series.append(Series(Reference(ws, min_col=3, max_col=2 + nkeys, min_row=row, max_row=row),
                                title=name))
    ch.set_categories(Reference(ws, min_col=3, max_col=2 + nkeys, min_row=hr, max_row=hr))
    ws.add_chart(ch, "A%d" % (at_row + 2))

def reconcile(txns):
    card = defaultdict(float)
    for t in txns:
        if t["kind"] == "card" and t["billing_date"]:
            card[(t["card_company"], t["billing_date"])] += -amt(t)
    dates = defaultdict(list)
    for comp, b in card:
        dates[comp].append(dt.date.fromisoformat(b))
    bank, unassigned = defaultdict(float), []
    for t in txns:
        if t["status"] != "card_debit":
            continue
        d, best = dget(t, "date"), None
        for b in dates.get(t["card_company"], []):
            off = (d - b).days
            if CARD_DEBIT_WINDOW[0] <= off <= CARD_DEBIT_WINDOW[1] and (best is None or abs(off) < abs((d - best).days)):
                best = b
        if best:
            bank[(t["card_company"], iso(best))] += -amt(t)
        else:
            unassigned.append(t)
    rows = []
    for k in sorted(set(card) | set(bank)):
        c, b = round(card.get(k, 0), 2), round(bank.get(k, 0), 2)
        if k not in bank:
            st = "לא נמצא חיוב בבנק (חיוב עתידי, או מחשבון שלא נטען)"
        elif abs(b - c) < 1:
            st = "תקין"
        else:
            st = "הפרש - כנראה חסר קובץ של אחד הכרטיסים"
        rows.append(dict(company=k[0], date=k[1], card=c, bank=b, diff=round(b - c, 2), status=st))
    return rows, unassigned

def write_checks(wb, txns, accounts):
    ws = new_sheet(wb, "בדיקות")
    ws.cell(1, 1, "בדיקות תקינות").font = Font(bold=True, size=14)
    r = 3
    pend = sum(1 for t in txns if t["status"] == "pending")
    ws.cell(r, 1, "תנועות שממתינות לסיווג (מופיעות בדוח תחת 'לא מסווג')").font = BOLD
    ws.cell(r, 2, pend)
    r += 2
    head = ["חשבון / כרטיס", "מתאריך", "עד תאריך", "תנועות"]
    for j, h in enumerate(head, 1):
        ws.cell(r, j, h)
    style_header(ws, r, 4)
    r += 1
    span = defaultdict(list)
    for t in txns:
        span[t["account"]].append(t["date"])
    for acc in sorted(span):
        ds = sorted(span[acc])
        ws.cell(r, 1, acct_label(acc, accounts))
        ws.cell(r, 2, dt.date.fromisoformat(ds[0])).number_format = DATEF
        ws.cell(r, 3, dt.date.fromisoformat(ds[-1])).number_format = DATEF
        ws.cell(r, 4, len(ds))
        r += 1
    known = {(t["card_company"], t["account"].split()[-1]) for t in txns if t["kind"] == "card"}
    missing_cards = Counter()
    for t in txns:
        if t["status"] == "card_debit" and re.fullmatch(r"\d{4}", t["ref"]) \
                and (t["card_company"], t["ref"]) not in known:
            missing_cards[(t["card_company"], t["ref"])] += -amt(t)
    r += 1
    ws.cell(r, 1, "כרטיסים שמחויבים בבנק ואין להם קובץ פירוט").font = BOLD
    r += 1
    if not missing_cards:
        ws.cell(r, 1, "אין")
        r += 1
    for (comp, last4), total in sorted(missing_cards.items()):
        ws.cell(r, 1, "%s %s" % (comp, last4))
        ws.cell(r, 2, round(total, 2)).number_format = NUM
        ws.cell(r, 3, "סה\"כ חיובים בבנק בתקופה")
        r += 1
    rows, unassigned = reconcile(txns)
    r += 1
    ws.cell(r, 1, "התאמה בין חיובי הכרטיסים בבנק לבין קבצי הכרטיסים").font = BOLD
    r += 1
    head = ["חברת אשראי", "תאריך חיוב", "סה\"כ בקבצי כרטיס", "סה\"כ חיוב בבנק", "הפרש", "מצב"]
    for j, h in enumerate(head, 1):
        ws.cell(r, j, h)
    style_header(ws, r, 6)
    r += 1
    missing = defaultdict(float)
    for x in rows:
        ws.cell(r, 1, x["company"])
        ws.cell(r, 2, dt.date.fromisoformat(x["date"])).number_format = DATEF
        for j, k in ((3, "card"), (4, "bank"), (5, "diff")):
            ws.cell(r, j, x[k]).number_format = NUM2
        ws.cell(r, 6, x["status"])
        if x["status"] == "תקין":
            ws.cell(r, 6).fill = PatternFill("solid", fgColor="C6EFCE")
        r += 1
        if x["diff"] > 1:
            missing[(x["date"][:7], x["company"])] += x["diff"]
    for t in unassigned:
        missing[(t["date"][:7], t["card_company"])] += -amt(t)
    r += 1
    ws.cell(r, 1, "הוצאות בכרטיסי אשראי שחסרות בדוח (חיוב בבנק בלי קובץ פירוט מתאים)").font = BOLD
    r += 1
    for j, h in enumerate(["חודש", "חברת אשראי", "סכום חסר"], 1):
        ws.cell(r, j, h)
    style_header(ws, r, 3)
    r += 1
    for (month, comp), v in sorted(missing.items()):
        ws.cell(r, 1, month)
        ws.cell(r, 2, comp)
        ws.cell(r, 3, round(v, 2)).number_format = NUM
        r += 1
    widths(ws, [44, 16, 22, 18, 12, 50])

def write_tx_sheet(wb, txns, accounts, year):
    ws = new_sheet(wb, "תנועות")
    head = ["תאריך", "חודש בדוח", "חשבון / כרטיס", "תיאור", "פרטים", "סכום", "קטגוריה", "תת-קטגוריה",
            "מצב", "הסבר", "מזהה"]
    ws.append(head)
    style_header(ws, 1, len(head))
    for t in sorted(txns, key=lambda t: t["report_date"] or t["date"], reverse=True):
        rd = dget(t)
        if year and rd.year != year:
            continue
        cat = t["category"] or (UNCAT if t["status"] == "pending" else "")
        ws.append([dget(t, "date"), "%02d/%d" % (rd.month, rd.year), acct_label(t["account"], accounts),
                   t["description"], t["notes"], amt(t), cat, t["subcategory"],
                   STATUS_HE.get(t["status"], t["status"]), t["reason"], t["id"]])
        ws.cell(ws.max_row, 1).number_format = DATEF
        ws.cell(ws.max_row, 6).number_format = NUM2
    ws.auto_filter.ref = "A1:K%d" % ws.max_row
    ws.freeze_panes = "A2"
    widths(ws, [12, 10, 24, 32, 36, 12, 22, 20, 22, 36, 14])

# ------------------------------------------------- yearly workbook ("רישום כלכלי")
MONTHS_HE = ["ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני", "יולי", "אוגוסט", "ספטמבר",
             "אוקטובר", "נובמבר", "דצמבר"]
ITEM_SHEETS = ("שונות", "מתנות ותרומות")        # one line per payment, like your manual file
BAL_FIELDS = ["שם", "חודש", "סכום"]               # config/balances.csv - money on the 15th
NOTE_FIELDS = ["חודש", "אזור", "הערה"]            # config/notes.csv - אזור: הוצאות / הכנסות
SALARY = "משכורת"
NUMF = '#,##0;-#,##0;"-"'
GREEN = PatternFill("solid", fgColor="E2EFDA")
BLUE = PatternFill("solid", fgColor="DDEBF7")

def qs(name):
    return "'%s'" % name.replace("'", "''")

def safe_title(name):
    return re.sub(r"[\[\]:*?/\\]", "", name)[:31]

def month_header(ws, row, first_col, avg_col=None):
    if avg_col:
        ws.cell(row, avg_col, "ממוצע")
    for m in range(12):
        ws.cell(row, first_col + m, MONTHS_HE[m])
    for c in range(1, first_col + 13):
        ws.cell(row, c).font = BOLD

def write_yearly(year, lines, cats, order):
    """Workbook in the layout of 'רישום כלכלי': general sheet + a sheet per category with sub-categories."""
    exp = defaultdict(lambda: defaultdict(float))
    inc = defaultdict(lambda: defaultdict(float))
    items = defaultdict(lambda: defaultdict(list))
    months = set()
    for t, a, c, s in lines:
        d = dget(t)
        if d.year != year:
            continue
        months.add(d.month)
        if ctype(cats, c, a) == EXPENSE:
            exp[(c, s)][d.month] += -a
            if c in ITEM_SHEETS:
                desc = re.sub(r"\s{2,}.*$", "", t["description"])
                if t["notes"]:
                    desc += " - " + t["notes"][:40]
                if s:
                    desc = "%s (%s)" % (desc, s)
                items[c][d.month].append((t["date"], desc, -a))
        else:
            inc[(c, s)][d.month] += a
    last = max(months) if months else 12
    today = dt.date.today()
    if year == today.year:
        last = min(last, today.month)               # don't average over future months
    L = get_column_letter(2 + last)                 # last month column with data
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "כללי"
    ws.sheet_view.rightToLeft = True

    exp_cats = [c for c in order if cats[c]["type"] == EXPENSE]
    exp_cats += sorted({c for (c, s) in exp if c not in cats and c != UNCAT})
    if any(c == UNCAT for (c, s) in exp):
        exp_cats.append(UNCAT)

    # ---- sheets per category (sub-categories / item lists)
    ref = {}
    for c in exp_cats:
        if c in ITEM_SHEETS:
            w = new_sheet(wb, safe_title(c))
            for m in range(12):
                w.cell(1, 2 * m + 1, "תיאור").font = BOLD
                w.cell(1, 2 * m + 2, MONTHS_HE[m]).font = BOLD
                col = get_column_letter(2 * m + 2)
                if m + 1 <= last:
                    w.cell(2, 2 * m + 2, "=SUM(%s3:%s500)" % (col, col)).number_format = NUMF
                for i, (_, desc, v) in enumerate(sorted(items[c][m + 1]), 3):
                    w.cell(i, 2 * m + 1, desc)
                    w.cell(i, 2 * m + 2, round(v, 2)).number_format = NUMF
            w.cell(2, 1, 'סה"כ').font = BOLD
            for m in range(12):
                w.cell(2, 2 * m + 2).font = BOLD
                w.column_dimensions[get_column_letter(2 * m + 1)].width = 24
                w.column_dimensions[get_column_letter(2 * m + 2)].width = 9
            ref[c] = lambda m, c=c: "=%s!%s2" % (qs(safe_title(c)), get_column_letter(2 * m + 2))
        elif c in cats and cats[c]["subs"]:
            subs = list(cats[c]["subs"])
            if any(k == (c, "") for k in exp):
                subs.append("")
            w = new_sheet(wb, safe_title(c))
            month_header(w, 1, 3, 2)
            r = 2
            for sname in subs:
                w.cell(r, 1, sname or "אחר")
                for m in range(1, last + 1):
                    w.cell(r, 2 + m, round(exp[(c, sname)].get(m, 0), 2)).number_format = NUMF
                w.cell(r, 2, "=AVERAGE(C%d:%s%d)" % (r, L, r)).number_format = NUMF
                r += 1
            w.cell(r, 1, "הערות")
            r += 1
            w.cell(r, 1, "סיכום חודשי").font = BOLD
            w.cell(r, 2, "=SUM(B2:B%d)" % (r - 2)).number_format = NUMF
            for m in range(1, last + 1):
                col = get_column_letter(2 + m)
                w.cell(r, 2 + m, "=SUM(%s2:%s%d)" % (col, col, r - 2)).number_format = NUMF
            for j in range(1, 15):
                w.cell(r, j).fill = TOTFILL
                w.cell(r, j).font = BOLD
            widths(w, [26, 10] + [10] * 12)
            ref[c] = lambda m, c=c, r=r: "=%s!%s%d" % (qs(safe_title(c)), get_column_letter(2 + m), r)

    # ---- general sheet: expenses
    ws.cell(1, 1, "הוצאות").font = Font(bold=True, size=14)
    month_header(ws, 2, 3, 2)
    ws.cell(2, 16, 'סה"כ').font = BOLD
    r = 3
    rows = {}
    for c in exp_cats:
        ws.cell(r, 1, c)
        for m in range(1, last + 1):
            v = ref[c](m) if c in ref else round(sum(x.get(m, 0) for (c2, s), x in exp.items() if c2 == c), 2)
            ws.cell(r, 2 + m, v).number_format = NUMF
        ws.cell(r, 2, "=AVERAGE(C%d:%s%d)" % (r, L, r)).number_format = NUMF
        ws.cell(r, 16, "=SUM(C%d:N%d)" % (r, r)).number_format = NUMF
        rows[c] = r
        r += 1
    exp_total = r
    ws.cell(r, 1, "הוצאות חודשי")
    ws.cell(r, 2, "=SUM(B3:B%d)" % (r - 1))
    for m in range(1, last + 1):
        col = get_column_letter(2 + m)
        ws.cell(r, 2 + m, "=SUM(%s3:%s%d)" % (col, col, r - 1))
    ws.cell(r, 16, "=SUM(C%d:N%d)" % (r, r))
    diff_row = r + 1
    notes_exp = r + 2

    # ---- income
    r = notes_exp + 2
    ws.cell(r, 1, "הכנסות חודשי").font = Font(bold=True, size=14)
    r += 1
    inc_first = r
    inc_rows = []
    sal_rows = []
    inc_cats = [c for c in order if cats[c]["type"] == INCOME]
    inc_cats += sorted({c for (c, s) in inc if c not in cats})
    for c in inc_cats:
        subs = (cats[c]["subs"] if c in cats else []) + ([""] if (c, "") in inc and c in cats and cats[c]["subs"] else [])
        parts = [(sname, sname if c == SALARY else "%s - %s" % (c, sname or "אחר")) for sname in subs] if subs else [("", c)]
        for sname, label in parts:
            ws.cell(r, 1, label)
            vals = inc[(c, sname)] if subs else defaultdict(float, {m: sum(x.get(m, 0) for (c2, s2), x in inc.items() if c2 == c) for m in range(1, 13)})
            for m in range(1, last + 1):
                ws.cell(r, 2 + m, round(vals.get(m, 0), 2)).number_format = NUMF
            ws.cell(r, 2, "=AVERAGE(C%d:%s%d)" % (r, L, r)).number_format = NUMF
            ws.cell(r, 16, "=SUM(C%d:N%d)" % (r, r)).number_format = NUMF
            if c == SALARY:
                sal_rows.append(r)
            inc_rows.append(r)
            r += 1
    inc_total = r
    ws.cell(r, 1, 'סה"כ הכנסות')
    ws.cell(r, 2, "=AVERAGE(C%d:%s%d)" % (r, L, r))
    for m in range(1, last + 1):
        col = get_column_letter(2 + m)
        ws.cell(r, 2 + m, "=SUM(%s%d:%s%d)" % (col, inc_first, col, r - 1))
    ws.cell(r, 16, "=SUM(C%d:N%d)" % (r, r))
    notes_inc = r + 2

    # difference + notes rows
    ws.cell(diff_row, 1, 'סה"כ הפרש')
    ws.cell(diff_row, 2, "=SUM(C%d:N%d)" % (diff_row, diff_row))
    for m in range(1, last + 1):
        col = get_column_letter(2 + m)
        ws.cell(diff_row, 2 + m, "=%s%d-%s%d" % (col, inc_total, col, exp_total))
    notes = read_csv(p(CONFIG, "notes.csv"))
    for row, area in ((notes_exp, "הוצאות"), (notes_inc, "הכנסות")):
        for m in range(1, 13):
            ws.cell(row, 2 + m, "הערות:").font = Font(italic=True, color="808080")
            txt = "; ".join(n["הערה"] for n in notes if n.get("חודש") == "%d-%02d" % (year, m)
                            and (n.get("אזור") or "הוצאות") == area and n.get("הערה"))
            if txt:
                ws.cell(row + 1, 2 + m, txt).alignment = Alignment(wrap_text=True, vertical="top")

    # maaser (10% of salaries) - as in your file
    if sal_rows:
        sr = sal_rows[0]
        ws.cell(sr, 18, "=" + "+".join("P%d" % x for x in sal_rows)).number_format = NUMF
        ws.cell(sr, 19, "סהכ משכורת")
        ws.cell(sr + 1, 18, "=R%d/10" % sr).number_format = NUMF
        ws.cell(sr + 1, 19, "סהכ רצון לתרום")
        if "מתנות ותרומות" in rows:
            ws.cell(sr + 3, 18, "כמה תרמנו עד עכשיו")
            ws.cell(sr + 3, 19, "=P%d" % rows["מתנות ותרומות"]).number_format = NUMF
            ws.cell(sr + 4, 18, "נשאר לתרום")
            ws.cell(sr + 4, 19, "=R%d-S%d" % (sr + 1, sr + 3)).number_format = NUMF

    # ---- money on the 15th (config/balances.csv)
    bal = [b for b in read_csv(p(CONFIG, "balances.csv")) if b.get("חודש", "").startswith(str(year))]
    names = []
    for b in bal:
        if b["שם"] not in names:
            names.append(b["שם"])
    if not names:
        prev = sorted((b for b in read_csv(p(CONFIG, "balances.csv")) if b.get("חודש", "") < str(year)),
                      key=lambda b: b["חודש"])
        for b in prev:
            if b["שם"] not in names:
                names.append(b["שם"])
    r = notes_inc + 3
    top = r
    ws.cell(r, 1, "כמות הכסף ב 15 לחודש").font = Font(bold=True, size=12)
    for i, n in enumerate(names, 1):
        ws.cell(r + i, 1, n)
        for b in bal:
            if b["שם"] == n:
                m = int(b["חודש"][5:7])
                ws.cell(r + i, 2 + m, num(b["סכום"])).number_format = NUMF
    end = r + max(len(names), 1)
    for m in range(1, 13):
        col = get_column_letter(2 + m)
        ws.cell(top, 2 + m, "=IF(COUNT(%s%d:%s%d)=0,\"\",SUM(%s%d:%s%d))" % (col, top + 1, col, end, col, top + 1, col, end)).number_format = NUMF
    ws.cell(end + 1, 1, 'סה"כ הפרש חודשי')
    for m in range(2, 13):
        col, pc = get_column_letter(2 + m), get_column_letter(1 + m)
        ws.cell(end + 1, 2 + m, "=IF(OR(%s%d=\"\",%s%d=\"\"),\"\",%s%d-%s%d)" % (col, top, pc, top, col, top, pc, top)).number_format = NUMF
    ws.cell(end + 2, 1, "(למילוי ב-config\\balances.csv)").font = Font(italic=True, color="808080")

    # ---- styling of the general sheet
    for rr in (2,):
        for j in range(1, 17):
            ws.cell(rr, j).fill = BLUE
    for rr in (exp_total, inc_total, diff_row, top):
        for j in range(1, 17):
            ws.cell(rr, j).font = BOLD
            ws.cell(rr, j).fill = TOTFILL if rr != diff_row else GREEN
            if ws.cell(rr, j).number_format == "General":
                ws.cell(rr, j).number_format = NUMF
    ws.cell(end + 1, 1).font = BOLD
    widths(ws, [30, 10] + [10] * 12 + [2, 11, 2, 12, 20])
    ws.freeze_panes = "C3"
    return wb

# ------------------------------------------------------------------ dashboard
def write_dashboard(txns, lines, cats, order, accounts):
    """Local HTML dashboard (output/דשבורד.html) - opens in the browser, no internet needed."""
    import json
    tpl_path = p(BASE, "dashboard_template.html")
    if not os.path.exists(tpl_path):
        print("  ! dashboard_template.html is missing - dashboard skipped")
        return None
    data = {
        "generated": dt.datetime.now().strftime("%d/%m/%Y %H:%M"),
        "today": dt.date.today().isoformat(),
        "pending": sum(1 for t in txns if t["status"] == "pending"),
        "cats": [{"name": c, "type": "i" if cats[c]["type"] == INCOME else "e"} for c in order],
        "balances": [{"name": b["שם"], "m": b["חודש"], "v": num(b["סכום"])}
                     for b in read_csv(p(CONFIG, "balances.csv")) if b.get("חודש")],
        "tx": [],
    }
    for t, a, c, s in lines:
        ty = ctype(cats, c, a)
        data["tx"].append({"d": t.get("report_date") or t["date"], "type": "i" if ty == INCOME else "e",
                           "v": round(a if ty == INCOME else -a, 2), "cat": c, "sub": s,
                           "desc": re.sub(r"\s{2,}.*$", "", t["description"]), "notes": t["notes"],
                           "acc": acct_label(t["account"], accounts)})
    blob = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
    with open(tpl_path, encoding="utf-8") as f:
        html = f.read().replace("/*DATA*/", blob)
    path = p(OUTPUT, "דשבורד.html")
    with open(path, "w", encoding="utf-8") as f:
        f.write(html)
    return path

def cmd_report(args):
    setup()
    cats, order = load_categories()
    rules = load_rules(cats)
    txns = load_tx()
    if not txns:
        sys.exit("No transactions yet - run import first.")
    splits = load_splits()
    accounts = load_accounts()
    classify(txns, rules)
    txns = txns + recurring_tx(cats)
    txns = txns + cash_tx(txns, cats)
    by_month = defaultdict(lambda: defaultdict(float))
    by_year = defaultdict(lambda: defaultdict(float))
    months = defaultdict(set)
    for t, a, c, s in report_lines(txns, splits):
        d = dget(t)
        ty = ctype(cats, c, a)
        v = a if ty == INCOME else -a
        by_month[(ty, c, s)][(d.year, d.month)] += v
        by_year[(ty, c, s)][d.year] += v
        months[d.year].add(d.month)
    years = sorted(months, reverse=True)
    if args.year:
        if args.year not in months:
            sys.exit("No data for %d. Years with data: %s" % (args.year, ", ".join(map(str, sorted(months)))))
        years = [args.year]
    plan = build_plan(cats, order, set(by_month))
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    yrs = sorted(years)
    ws = new_sheet(wb, "סיכום %d" % args.year if args.year else "שנתי")
    extra = [("ממוצע חודשי %d" % y, (lambda vals, y=y: vals.get(y, 0) / max(1, len(months[y])))) for y in yrs]
    write_matrix(ws, "סיכום שנתי" + (" %d" % args.year if args.year else ""), yrs, [str(y) for y in yrs],
                 by_year, plan, extra, show_total=False)
    for y in years:
        ws = new_sheet(wb, "חודשי %d" % y)
        keys = [(y, m) for m in range(1, 13)]
        n = max(1, len(months[y]))
        hr, ir, er, nr = write_matrix(
            ws, "תזרים חודשי %d" % y, keys, ["%02d/%d" % (m, y) for m in range(1, 13)], by_month, plan,
            [("ממוצע חודשי", lambda vals, keys=keys, n=n: sum(vals.get(k, 0) for k in keys) / n)])
        add_chart(ws, hr, ir, er, 12, nr, "הכנסות מול הוצאות %d" % y)
    write_tx_sheet(wb, txns, accounts, args.year)
    write_checks(wb, txns, accounts)
    path = p(OUTPUT, "report_%d.xlsx" % args.year if args.year else "report.xlsx")
    save_wb(wb, path)
    lines = report_lines(txns, splits)
    for y in years:
        if len({dget(t).month for t, _, _, _ in lines if dget(t).year == y}) < 3:
            continue
        yb = write_yearly(y, lines, cats, order)
        ws = new_sheet(yb, "תנועות")
        yb.remove(ws)
        write_tx_sheet(yb, txns, accounts, y)
        ypath = p(OUTPUT, "רישום כלכלי %d.xlsx" % y)
        try:
            yb.save(ypath)
            print("Yearly file written: output\\%s" % os.path.basename(ypath))
        except PermissionError:
            print("  ! %s is open in Excel - not updated. Close it and run the report again." % os.path.basename(ypath))
    dash = write_dashboard(txns, lines, cats, order, accounts)
    if dash:
        print("Dashboard written: output\\%s" % os.path.basename(dash))
    pend = sum(1 for t in txns if t["status"] == "pending")
    print("Report written: output\\%s" % os.path.basename(path))
    if pend:
        print("Note: %d transactions are still waiting for a category and appear as '%s'." % (pend, UNCAT))
    if not args.no_open:
        open_file(dash or path)

def cmd_status(args):
    setup()
    txns = load_tx()
    if not txns:
        print("No transactions yet - run import first.")
        return
    summary(txns)
    span = defaultdict(list)
    for t in txns:
        span[t["account"]].append(t["date"])
    print("\nData per account:")
    for acc in sorted(span):
        ds = sorted(span[acc])
        print("  %-30s %s .. %s  (%d)" % (acc, ds[0], ds[-1], len(ds)))

def main():
    ap = argparse.ArgumentParser(description="Family finance tracker")
    sub = ap.add_subparsers(dest="cmd")
    a = sub.add_parser("import", help="read new export files")
    a.add_argument("--all", action="store_true", help="re-read files that were already imported")
    a.add_argument("--force", action="store_true", help="overwrite review.xlsx even if it has unapplied answers")
    a.add_argument("--no-open", action="store_true")
    a = sub.add_parser("apply", help="apply your answers from output/review.xlsx")
    a.add_argument("--no-open", action="store_true")
    a = sub.add_parser("report", help="build the Excel report")
    a.add_argument("--year", type=int)
    a.add_argument("--no-open", action="store_true")
    sub.add_parser("status", help="short summary")
    args = ap.parse_args()
    cmds = {"import": cmd_import, "apply": cmd_apply, "report": cmd_report, "status": cmd_status}
    if args.cmd not in cmds:
        ap.print_help()
        return
    cmds[args.cmd](args)

if __name__ == "__main__":
    main()
