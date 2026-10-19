import os
import io
import re
import json
import time
import hmac
import email
import struct
import xml.etree.ElementTree as ET
import zipfile
import warnings
from email.header import decode_header
from datetime import datetime, timedelta, timezone

import gspread
import imapclient
import pandas as pd
from flask import Flask, Response, jsonify, request
from google.oauth2.service_account import Credentials

BASE = os.path.dirname(os.path.abspath(__file__))
app = Flask(__name__)

IST = timezone(timedelta(hours=5, minutes=30))
warnings.filterwarnings("ignore", message=".*Parsing dates.*")
warnings.filterwarnings("ignore", message=".*Could not infer format.*")
REQUIRED_COLUMNS = [
    "Admission No.",
    "DOA",
    "DOD",
    "Invoice Date",
    "Submission To SRIT Date",
]
DATE_COLUMNS    = ["DOA", "DOD", "Invoice Date", "Submission To SRIT Date"]
NUMERIC_COLUMNS = ["Sr No", "Net Bill Amount", "Spon Amount"]
UNIT_SHEETS = {
    "Indore":   "1L53r-Adww_UqWwkF-FI7GJTDG1aF5An5seI4wr5bR70",
    "Jabalpur": "1OodEyJHBXb31dbQi5QheLgZVNhgyfNLkB02uhCPqriI",
    "Jaipur":   "1XrMEI8IMa4iLYg42Ttfp_rpqtuBwaXrGzRnSpOZxW9o",
    "Krishna":  "1_xBC3dp47YQCibDtfE4vcIk85bqWOB7oK25HsONXOSE",
    "Mohali":   "1tkhSlXQEvn9e97OZaPWJGHYDjDOXEhOdn58EvoxZWyU",
    "Naroda":   "167rgUg8wISP7qZNbntSzO8T_IyqkUpMXrabsXx09wm4",
    "SG":       "1PWpJOPF07BRfguhVPCwJPkB7DLitoUKlDFFy4-0-kOA",
    "Surat":    "1Betva_Q9ltdL-gCT5Bx2Y9IPsEbmt1GIEDeUe3ww5MQ",
    "Vapi":     "1uwkuW2iHnGJB42vjBwj_1o2YdyQa1Bh32e9qxMCs5cA",
    "Vijay":    "1_8Me5oruzSG4XSI9DLJ29Bv_Xrd-tqsdksYkovrDAiM",
}
UNIT_LOOKUP = {k.lower(): k for k in UNIT_SHEETS}
WRITE_HEADER_IF_EMPTY = True
VALUE_INPUT   = "USER_ENTERED"
CHUNK_SIZE    = 5000
STAGES = ["IMAP", "ZIP", "CONVERT", "TRANSFORM", "SHEET", "DONE"]
COLS = ("run_id", "ts", "stage", "unit", "level", "message")
STALL_SECONDS = int(os.environ.get("STALL_SECONDS", "300"))   # इतनी देर कोई लॉग/हार्टबीट नहीं = STALLED
TAIL_ROWS = int(os.environ.get("TAIL_ROWS", "4000"))          # शीट की आख़िरी इतनी पंक्तियाँ ही पढ़ना
CACHE_TTL = int(os.environ.get("CACHE_TTL", "8"))                                                # सेकंड (Google API कोटा बचाने के लिए)
_cache = {"t": 0.0, "rows": []}


_GC = None


def _gc():
    global _GC
    if _GC is None:
        info = json.loads(os.environ["GOOGLE_CREDENTIALS"])
        creds = Credentials.from_service_account_info(
            info, scopes=["https://www.googleapis.com/auth/spreadsheets"])
        _GC = gspread.authorize(creds)
    return _GC


def _worksheet():
    return _log_ws(create=False)


def load_rows(fresh=False):
    """एक पोल = सिर्फ़ 1 Google Sheets रीड (वर्कशीट ऑब्जेक्ट कैश रहता है)"""
    if not fresh and time.time() - _cache["t"] < CACHE_TTL:
        return _cache["rows"]
    try:
        ws = _worksheet()
        vals = ws.get("A2:F20000")
    except gspread.WorksheetNotFound:      # पाइपलाइन ने अभी तक 'logs' टैब बनाया नहीं
        _cache.update(t=time.time(), rows=[])
        return []
    except Exception:
        if _cache["rows"]:                 # कोटा/नेटवर्क दिक्कत: पुराना डेटा दिखाते रहो
            _cache["t"] = time.time()
            return _cache["rows"]
        raise
    rows = [dict(zip(COLS, list(r) + [""] * (6 - len(r)))) for r in vals if r][-TAIL_ROWS:]
    _cache.update(t=time.time(), rows=rows)
    return rows


def _ts(s):
    try:
        return datetime.fromisoformat(s)
    except Exception:
        return None


def summarize(rid, rs, now):
    times = [t for t in (_ts(r["ts"]) for r in rs) if t]
    start, last = (min(times), max(times)) if times else (None, None)
    errors = sum(r["level"] == "ERROR" for r in rs)
    warns = sum(r["level"] == "WARN" for r in rs)
    done = [r for r in rs if r["stage"] == "DONE" and r["message"].startswith(("COMPLETE", "FAILED"))]
    idle = int((now - last).total_seconds()) if last else 0
    if done and done[-1]["message"].startswith("FAILED"):
        status = "FAILED"
    elif done:
        status = "COMPLETED"
    elif idle > STALL_SECONDS:
        status = "STALLED"
    else:
        status = "RUNNING"
    end = last if status != "RUNNING" else now
    return {
        "id": rid, "status": status, "errors": errors, "warns": warns,
        "count": sum(r["level"] != "BEAT" for r in rs), "idle_s": idle,
        "start": start.isoformat() if start else None,
        "last": last.isoformat() if last else None,
        "duration_s": int((end - start).total_seconds()) if start and end else 0,
    }


def build_state(run_id):
    now = datetime.now(IST)
    rows = load_rows()
    runs = {}
    for r in rows:
        runs.setdefault(r["run_id"], []).append(r)
    summaries = sorted((summarize(k, v, now) for k, v in runs.items()),
                       key=lambda s: s["id"], reverse=True)
    sel = run_id if run_id in runs else (summaries[0]["id"] if summaries else None)
    out = {"now": now.isoformat(), "runs": summaries[:30], "run": None,
           "stages": [], "units": [], "logs": []}
    if not sel:
        return out

    rs = runs[sel]
    run = next(s for s in summaries if s["id"] == sel)
    out["run"] = run
    logs = [r for r in rs if r["level"] != "BEAT"]
    out["logs"] = logs[-1500:]

    seen = [r["stage"] for r in logs if r["stage"] in STAGES]
    current = seen[-1] if seen else None
    for s in STAGES:
        sr = [r for r in logs if r["stage"] == s]
        if not sr:
            st = "pending"
        elif any(r["level"] == "ERROR" for r in sr):
            st = "error"
        elif s == current and run["status"] in ("RUNNING", "STALLED"):
            st = "stalled" if run["status"] == "STALLED" else "active"
        else:
            st = "done"
        out["stages"].append({"name": s, "status": st, "logs": len(sr)})

    units = {}
    for r in logs:
        if not r["unit"]:
            continue
        u = units.setdefault(r["unit"], {"name": r["unit"], "stage": "", "errors": 0,
                                         "warns": 0, "done": False, "last": ""})
        u["stage"], u["last"] = r["stage"], r["message"]
        u["errors"] += r["level"] == "ERROR"
        u["warns"] += r["level"] == "WARN"
        if r["stage"] == "SHEET" and r["level"] == "SUCCESS" and "जोड़ीं" in r["message"]:
            u["done"] = True
    for u in units.values():
        u["status"] = "done" if u["done"] else "error" if u["errors"] else (
            "active" if run["status"] == "RUNNING" else "incomplete")
    out["units"] = sorted(units.values(), key=lambda u: u["name"])
    return out


def match_unit(name):
    """फ़ाइल के नाम से यूनिट पहचानो (छोटे/बड़े अक्षर का फ़र्क़ नहीं)"""
    return UNIT_LOOKUP.get(str(name).strip().lower())


def records(data):
    pos, n = 0, len(data)
    while pos < n:
        b = data[pos]; pos += 1
        rid = b & 0x7F
        if b & 0x80:
            b2 = data[pos]; pos += 1
            rid |= (b2 & 0x7F) << 7
        size, shift = 0, 0
        for _ in range(4):
            b = data[pos]; pos += 1
            size |= (b & 0x7F) << shift
            shift += 7
            if not (b & 0x80):
                break
        yield rid, data[pos:pos + size]
        pos += size


def wide_str(buf, off):
    ln = struct.unpack_from("<I", buf, off)[0]
    s = buf[off + 4: off + 4 + 2 * ln].decode("utf-16-le", "replace")
    return s, off + 4 + 2 * ln


def read_sst(zf):
    out = []
    if "xl/sharedStrings.bin" not in zf.namelist():
        return out
    for rid, d in records(zf.read("xl/sharedStrings.bin")):
        if rid == 19:
            out.append(wide_str(d, 1)[0])
    return out


def read_sheets(zf):
    rels = {}
    root = ET.fromstring(zf.read("xl/_rels/workbook.bin.rels"))
    for el in root:
        rels[el.attrib["Id"]] = el.attrib["Target"]
    sheets = []
    for rid, d in records(zf.read("xl/workbook.bin")):
        if rid == 156:
            r, off = wide_str(d, 8)
            name, _ = wide_str(d, off)
            t = rels[r].lstrip("/")
            if not t.startswith("xl/"):
                t = "xl/" + t
            sheets.append((name, t))
    return sheets


def rk(v):
    if v & 2:
        x = struct.unpack("<i", struct.pack("<I", v & 0xFFFFFFFC))[0] >> 2
    else:
        x = struct.unpack("<d", struct.pack("<Q", (v & 0xFFFFFFFC) << 32))[0]
    if v & 1:
        x /= 100
    return x


def to_text(v):
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        if v == int(v) and abs(v) < 1e15:
            return str(int(v))
        return repr(v)
    if isinstance(v, int):
        return str(v)
    return v


def convert_date(value):
    """तारीख़ को असली तारीख़ में बदलो (टेक्स्ट, Excel सीरियल नंबर, या पहले से तारीख़)"""
    if pd.isna(value) or str(value).strip() == "":
        return pd.NaT

    # पहले से वास्तविक तारीख़ हो
    if isinstance(value, (datetime, pd.Timestamp)):
        return pd.to_datetime(value, errors="coerce")

    # Excel serial date (असली नंबर)
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if 1 <= value <= 100000:
            return pd.to_datetime(value, unit="D", origin="1899-12-30", errors="coerce")
        return pd.NaT

    text = str(value).strip()

    # Excel serial date (टेक्स्ट के रूप में, जैसे "46310") - XLSB→XLSX के ALL_TEXT की वजह से
    try:
        num = float(text.replace(",", ""))
        if 1 <= num <= 100000:
            return pd.to_datetime(num, unit="D", origin="1899-12-30", errors="coerce")
    except ValueError:
        pass

    formats = [
        "%d-%b-%y",    # 15-OCT-26
        "%d-%b-%Y",
        "%d-%m-%Y",
        "%d-%m-%y",
        "%d/%m/%Y",
        "%d/%m/%y",
        "%Y-%m-%d",
        "%d-%B-%Y",
    ]
    for fmt in formats:
        try:
            return pd.Timestamp(datetime.strptime(text, fmt))
        except ValueError:
            continue

    return pd.to_datetime(text, errors="coerce", dayfirst=True)


def convert_to_number(value):
    """टेक्स्ट वाले नंबर को असली नंबर में बदलो"""
    if pd.isna(value):
        return value

    if isinstance(value, (int, float)):
        return value

    text = str(value).strip()
    if text == "":
        return None

    text = (text.replace(",", "")
                .replace("₹", "")
                .replace("Rs.", "")
                .replace("Rs", "")
                .strip())

    # (1200) को -1200 मानो
    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1].strip()

    try:
        number = float(text)
        if number.is_integer():
            return int(number)
        return number
    except (ValueError, OverflowError):
        return value


def transform_dataframe(df):
    """एक यूनिट के डेटा पर पूरी ट्रांसफ़ॉर्मेशन; गड़बड़ हो तो ValueError"""
    df = df.copy()

    # कॉलम नाम साफ़ करो
    df.columns = [str(c).strip() if c is not None else "" for c in df.columns]

    # 1. ज़रूरी कॉलम जाँचो
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing:
        raise ValueError(f"ज़रूरी कॉलम नहीं मिले: {missing} | मौजूद: {df.columns.tolist()}")

    # 2. 'Submission To SRIT Date' के बाद के कॉलम हटाओ
    srit_index = df.columns.get_loc("Submission To SRIT Date")
    df = df.iloc[:, :srit_index + 1].copy()

    # 3. Claim ID का डेटा खाली करो
    if "Claim ID" in df.columns:
        df["Claim ID"] = ""
    else:
        print("      ⚠️  चेतावनी: Claim ID कॉलम नहीं मिला")

    # 4. File Type: OP / IP (पहले से भरी वैल्यू नहीं बदलनी)
    if "File Type" in df.columns:
        existing_type = df["File Type"].copy()
        df = df.drop(columns=["File Type"])
    else:
        existing_type = pd.Series([None] * len(df), index=df.index, dtype=object)

    admission = df["Admission No."].fillna("").astype(str).str.strip()
    calculated_type = admission.apply(lambda x: "OP" if "-" in x else "IP")

    existing_clean = existing_type.fillna("").astype(str).str.strip()
    is_blank = existing_clean.eq("")

    df["File Type"] = existing_type
    df.loc[is_blank, "File Type"] = calculated_type.loc[is_blank]

    # File Type को D कॉलम (index 3) में रखो
    cols = df.columns.tolist()
    cols.remove("File Type")
    cols.insert(3, "File Type")
    df = df[cols]

    # 5. 'Month -Year' कॉलम Invoice Date के ठीक बाद
    if "Month -Year" in df.columns:
        df = df.drop(columns=["Month -Year"])
    invoice_index = df.columns.get_loc("Invoice Date")
    df.insert(invoice_index + 1, "Month -Year", "")

    # 6. चारों तारीख़ कॉलम बदलो
    for col in DATE_COLUMNS:
        df[col] = df[col].apply(convert_date)

    # 7. Month -Year बनाओ (जैसे Oct-2026); Invoice Date खाली तो खाली
    df["Month -Year"] = df["Invoice Date"].apply(
        lambda d: d.strftime("%b-%Y") if pd.notna(d) else ""
    )

    # 8. टेक्स्ट नंबर → असली नंबर
    for col in NUMERIC_COLUMNS:
        if col in df.columns:
            df[col] = df[col].apply(convert_to_number)
        else:
            print(f"      ⚠️  ध्यान दें: '{col}' कॉलम नहीं मिला")

    # 9. Invoice Date से सॉर्ट: पुरानी → नई
    df = df.sort_values(
        by="Invoice Date", ascending=True, na_position="last", kind="stable"
    ).reset_index(drop=True)

    return df


def clean_value(v):
    """शीट में भेजने लायक सादी Python वैल्यू"""
    if hasattr(v, "item"):
        try:
            v = v.item()
        except Exception:
            pass
    if isinstance(v, float) and v != v:   # NaN
        return ""
    return v


def df_to_sheet_rows(df):
    """ट्रांसफ़ॉर्म हुए डेटा को शीट की पंक्तियों में बदलो"""
    out = df.copy()
    epoch = pd.Timestamp("1899-12-30")

    # तारीख़ें सीरियल नंबर बनकर जाएँगी (हर लोकेल में सही), बाद में dd-mm-yyyy फ़ॉर्मेट लगेगा
    for col in DATE_COLUMNS:
        out[col] = out[col].apply(
            lambda d: int((d.normalize() - epoch).days) if pd.notna(d) else ""
        )

    # 'Oct-2026' को शीट तारीख़ न बना दे, इसलिए शुरू में ' लगाया (टेक्स्ट बना रहेगा)
    out["Month -Year"] = out["Month -Year"].apply(lambda v: f"'{v}" if v else "")

    out = out.astype(object).where(out.notna(), "")
    return [[clean_value(v) for v in row] for row in out.values.tolist()]


def append_to_sheet(client, sheet_id, header, rows, date_col_idx):
    """हर बार जोड़ने से पहले एक खाली पंक्ति (सिपरेटर)"""
    if not rows:
        print("      ⚠️  कोई डेटा नहीं")
        return 0

    sh = client.open_by_key(sheet_id)
    ws = sh.sheet1
    existing = ws.get_all_values()

    if not existing:
        start_row = 2                      # पहली पंक्ति खाली (या हेडर)
        if WRITE_HEADER_IF_EMPTY:
            ws.update(values=[header], range_name="A1", value_input_option=VALUE_INPUT)
            print("      🏷️  शीट खाली थी, पहली पंक्ति में हेडर लिखा")
    else:
        start_row = len(existing) + 2      # सिपरेटर + डेटा

    end_row = start_row + len(rows) - 1

    # ग्रिड छोटी पड़े तो बढ़ाओ
    if end_row > ws.row_count:
        ws.add_rows(end_row - ws.row_count + 50)
    if len(header) > ws.col_count:
        ws.add_cols(len(header) - ws.col_count)

    for i in range(0, len(rows), CHUNK_SIZE):
        chunk = rows[i:i + CHUNK_SIZE]
        ws.update(values=chunk, range_name=f"A{start_row + i}",
                  value_input_option=VALUE_INPUT)

    # चारों तारीख़ कॉलम पर dd-mm-yyyy फ़ॉर्मेट (एक ही API कॉल में)
    if date_col_idx:
        requests = [{
            "repeatCell": {
                "range": {
                    "sheetId": ws.id,
                    "startRowIndex": start_row - 1,
                    "endRowIndex": end_row,
                    "startColumnIndex": c,
                    "endColumnIndex": c + 1,
                },
                "cell": {"userEnteredFormat": {
                    "numberFormat": {"type": "DATE", "pattern": "dd-mm-yyyy"}}},
                "fields": "userEnteredFormat.numberFormat",
            }
        } for c in date_col_idx]
        sh.batch_update({"requests": requests})

    print(f"      ✅ {len(rows)} पंक्तियाँ जोड़ीं "
          f"(खाली पंक्ति {start_row - 1} के बाद, डेटा {start_row}–{end_row})")
    return len(rows)


# ============================================================
# पाइपलाइन: मेल → ज़िप → XLSB → ट्रांसफ़ॉर्म → Google Sheet (सब Vercel पर)
# ============================================================
SUBJECT_KEYWORD = os.environ.get("SUBJECT_KEYWORD", "Portal Data")
LOG_HEADER = ["run_id", "ts", "stage", "unit", "level", "message"]


class RunLog:
    """लॉग बफ़र; flush() पर शीट के 'logs' टैब में जुड़ता है"""
    def __init__(self, run_id):
        self.run_id, self.stage, self.unit, self.buf = run_id, "", "", []

    def log(self, msg, level="INFO", stage=None, unit=None):
        self.buf.append([self.run_id, datetime.now(IST).isoformat(timespec="seconds"),
                         self.stage if stage is None else stage,
                         self.unit if unit is None else unit, level, str(msg)[:500]])

    def flush(self):
        if not self.buf:
            return
        global _LOGWS
        for attempt in (1, 2):
            try:
                _log_ws(create=True).append_rows(self.buf, value_input_option="RAW")
                self.buf = []
                return
            except Exception as e:
                print("log flush failed:", e)
                _LOGWS = None
                if attempt == 1:
                    time.sleep(6)


_LOGWS = None


def _log_ws(create=False):
    global _LOGWS
    if _LOGWS is not None:
        return _LOGWS
    sh = _gc().open_by_key(os.environ["LOG_SHEET_ID"])
    tab = os.environ.get("LOG_TAB", "logs")
    try:
        _LOGWS = sh.worksheet(tab)
    except gspread.WorksheetNotFound:
        if not create:
            raise
        _LOGWS = sh.add_worksheet(tab, rows=2000, cols=6)
        _LOGWS.update(values=[LOG_HEADER], range_name="A1")
    return _LOGWS


def _imap():
    c = imapclient.IMAPClient(os.environ.get("IMAP_SERVER", "shalbyhospitals.icewarpcloud.in"),
                              port=993, ssl=True)
    c.login(os.environ.get("IMAP_EMAIL", "fa30.corp@shalby.in"), os.environ["IMAP_PASSWORD"])
    c.select_folder("INBOX")
    return c


def _dec(value):
    out = ""
    for text, enc in decode_header(value or ""):
        out += text.decode(enc or "utf-8", errors="ignore") if isinstance(text, bytes) else text
    return out


def _zip_parts(msg):
    for part in msg.walk():
        fn = part.get_filename()
        if fn and _dec(fn).lower().endswith(".zip"):
            yield _dec(fn), part.get_payload(decode=True)


def scan_mailbox(log):
    today = datetime.now(IST).date()
    c = _imap()
    plan = []
    try:
        since = (today - timedelta(days=1)).strftime("%d-%b-%Y")
        uids = c.search(["SINCE", since, "SUBJECT", SUBJECT_KEYWORD])
        log.log(f"🔎 सर्वर से {len(uids)} मेल मिले")
        dates = c.fetch(uids, ["INTERNALDATE"]) if uids else {}
        for uid in uids:
            d = dates[uid].get(b"INTERNALDATE")
            if d is not None:
                if d.tzinfo is None:
                    d = d.replace(tzinfo=timezone.utc)
                d = d.astimezone(IST).date()
            if d != today:
                continue
            msg = email.message_from_bytes(c.fetch([uid], ["RFC822"])[uid][b"RFC822"])
            log.log(f"✅ मैच ({d}): {_dec(msg.get('Subject', ''))}")
            for zname, payload in _zip_parts(msg):
                with zipfile.ZipFile(io.BytesIO(payload)) as z:
                    for name in z.namelist():
                        if name.endswith("/") or "__MACOSX" in name or not name.lower().endswith(".xlsb"):
                            continue
                        base = os.path.splitext(os.path.basename(name))[0]
                        unit = match_unit(base)
                        plan.append({"uid": int(uid), "zip": zname, "member": name, "unit": unit or ""})
                        if unit:
                            log.log(f"📄 {zname} → {name}", unit=unit)
                        else:
                            log.log(f"⏭️  {base}: कोई शीट नहीं मिली, छोड़ा", "WARN", unit=base)
    finally:
        try:
            c.logout()
        except Exception:
            pass
    return plan


def fetch_member(uid, zipname, member):
    c = _imap()
    try:
        msg = email.message_from_bytes(c.fetch([int(uid)], ["RFC822"])[int(uid)][b"RFC822"])
        for zname, payload in _zip_parts(msg):
            if zname == zipname:
                with zipfile.ZipFile(io.BytesIO(payload)) as z:
                    return z.read(member)
    finally:
        try:
            c.logout()
        except Exception:
            pass
    raise ValueError("मेल/ज़िप दोबारा नहीं मिली")


def xlsb_first_sheet(blob):
    """XLSB की पहली शीट → पंक्तियों की सूची (सब वैल्यू टेक्स्ट, खाली = None)"""
    with zipfile.ZipFile(io.BytesIO(blob)) as zf:
        sst = read_sst(zf)
        sheets = read_sheets(zf)
        if not sheets:
            raise ValueError("XLSB में कोई शीट नहीं")
        data = zf.read(sheets[0][1])
    grid, row, next_col, max_col = {}, 0, 0, 0
    for rid, d in records(data):
        if rid == 0:
            row = struct.unpack_from("<I", d, 0)[0]
            next_col = 0
            if len(d) >= 21 and struct.unpack_from("<I", d, 13)[0] > 0:
                next_col = struct.unpack_from("<I", d, 17)[0]
        elif 1 <= rid <= 18:
            if rid <= 11:
                col, p = struct.unpack_from("<I", d, 0)[0], 8
            else:
                col, p = next_col, 4
            next_col = col + 1
            v = None
            try:
                if rid in (2, 13):   v = rk(struct.unpack_from("<I", d, p)[0])
                elif rid in (4, 15, 10): v = bool(d[p])
                elif rid in (5, 16, 9):  v = struct.unpack_from("<d", d, p)[0]
                elif rid in (6, 17, 8):  v = wide_str(d, p)[0]
                elif rid in (7, 18):     v = sst[struct.unpack_from("<I", d, p)[0]]
                elif rid in (3, 14, 11): v = "#ERR"
            except Exception:
                pass
            if v is not None:
                grid.setdefault(row, {})[col] = to_text(v)
                max_col = max(max_col, col)
    if not grid:
        return []
    return [[grid.get(r, {}).get(c) for c in range(max_col + 1)] for r in range(max(grid) + 1)]


def rows_to_df(rows):
    if len(rows) < 2:
        raise ValueError("फ़ाइल में कोई डेटा नहीं")
    seen, cols = {}, []
    for i, h in enumerate(rows[0]):
        h = ("" if h is None else str(h).strip()) or f"Unnamed: {i}"
        n = seen.get(h, 0)
        seen[h] = n + 1
        cols.append(h if n == 0 else f"{h}.{n}")
    body = [[None if v == "" else v for v in r] for r in rows[1:]]
    df = pd.DataFrame(body, columns=cols, dtype=object).dropna(how="all").reset_index(drop=True)
    if df.empty:
        raise ValueError("फ़ाइल में कोई डेटा नहीं")
    return df


def already_done(run_id, unit):
    return any(r["run_id"] == run_id and r["unit"] == unit and r["stage"] == "SHEET"
               and r["level"] == "SUCCESS" and "जोड़ीं" in r["message"]
               for r in load_rows(fresh=True))


def process_unit(log, uid, zipname, member, unit_raw):
    unit = match_unit(unit_raw)
    if not unit:
        raise ValueError(f"{unit_raw}: कोई शीट नहीं मिली")
    log.stage, log.unit = "CONVERT", unit
    log.log(f"🔄 {member} डाउनलोड/कन्वर्ट हो रही है")
    blob = fetch_member(uid, zipname, member)
    rows = xlsb_first_sheet(blob)
    log.log(f"✅ XLSB पढ़ी ({len(blob)//1024} KB, {max(len(rows)-1, 0)} पंक्तियाँ)", "SUCCESS")
    df = rows_to_df(rows)

    log.stage = "TRANSFORM"
    tdf = transform_dataframe(df)
    log.log(f"✅ {len(tdf)} पंक्तियाँ, {len(tdf.columns)} कॉलम ट्रांसफ़ॉर्म हुए "
            f"(OP: {(tdf['File Type'].astype(str).str.strip() == 'OP').sum()}, "
            f"IP: {(tdf['File Type'].astype(str).str.strip() == 'IP').sum()}, "
            f"खाली Invoice Date: {tdf['Invoice Date'].isna().sum()})", "SUCCESS")

    log.stage = "SHEET"
    log.log("📤 Google Sheet में जोड़ रहे हैं")
    log.flush()
    header = [str(c) for c in tdf.columns]
    date_idx = [tdf.columns.get_loc(c) for c in DATE_COLUMNS]
    n = append_to_sheet(_gc(), UNIT_SHEETS[unit], header, df_to_sheet_rows(tdf), date_idx)
    log.log(f"✅ {n} पंक्तियाँ जोड़ीं (Google Sheet)", "SUCCESS")
    return n


def _guard():
    pin = os.environ.get("RUN_PIN", "")
    if not pin:
        return jsonify(error="Vercel में RUN_PIN environment variable सेट नहीं है"), 403
    if not hmac.compare_digest(request.headers.get("X-Run-Pin", ""), pin):
        return jsonify(error="गलत PIN"), 401
    return None


@app.route("/api/run/start", methods=["POST"])
def run_start():
    bad = _guard()
    if bad:
        return bad
    try:
        run_id = datetime.now(IST).strftime("%Y%m%d-%H%M%S")
        lg = RunLog(run_id)
        lg.log("🚀 Pipeline शुरू (Vercel)", stage="IMAP")
        lg.flush()
        return jsonify(run_id=run_id)
    except Exception as e:
        return jsonify(error=f"{type(e).__name__}: {e}"), 500


@app.route("/api/run/scan", methods=["POST"])
def run_scan():
    bad = _guard()
    if bad:
        return bad
    body = request.get_json(force=True, silent=True) or {}
    lg = RunLog(body.get("run_id", ""))
    lg.stage = "IMAP"
    try:
        plan = scan_mailbox(lg)
        lg.stage, lg.unit = "ZIP", ""
        lg.log(f"📊 कुल {len(plan)} XLSB फ़ाइलें मिलीं", "SUCCESS" if plan else "WARN")
        lg.flush()
        return jsonify(plan=plan)
    except Exception as e:
        lg.log(f"❌ मेल स्कैन फ़ेल: {type(e).__name__}: {e}", "ERROR")
        lg.flush()
        return jsonify(error=f"{type(e).__name__}: {e}"), 500


@app.route("/api/run/unit", methods=["POST"])
def run_unit():
    bad = _guard()
    if bad:
        return bad
    b = request.get_json(force=True, silent=True) or {}
    lg = RunLog(b.get("run_id", ""))
    unit = match_unit(b.get("unit", "")) or b.get("unit", "")
    try:
        if already_done(lg.run_id, unit):
            lg.log("⏭️  इस रन में यह यूनिट पहले ही शीट में जुड़ चुकी है, दोबारा नहीं जोड़ी", "WARN",
                   stage="SHEET", unit=unit)
            lg.flush()
            return jsonify(ok=True, skipped=True)
        n = process_unit(lg, b["uid"], b["zip"], b["member"], unit)
        lg.flush()
        return jsonify(ok=True, rows=n)
    except Exception as e:
        lg.log(f"❌ {type(e).__name__}: {e}", "ERROR", unit=unit)
        lg.flush()
        return jsonify(error=f"{type(e).__name__}: {e}"), 500


@app.route("/api/run/finish", methods=["POST"])
def run_finish():
    bad = _guard()
    if bad:
        return bad
    b = request.get_json(force=True, silent=True) or {}
    lg = RunLog(b.get("run_id", ""))
    for u in b.get("hard", []):
        lg.log(f"❌ {u}: जवाब नहीं आया (टाइमआउट/नेटवर्क) - Retry करें", "ERROR", stage="SHEET", unit=u)
    lg.stage, lg.unit = "DONE", ""
    if b.get("fatal"):
        lg.log(f"FAILED: {b['fatal']}", "ERROR")
    else:
        lg.log(f"COMPLETE: {b.get('ok', 0)} यूनिट पूरी, {b.get('fail', 0)} फ़ेल",
               "SUCCESS" if not b.get("fail") else "WARN")
    lg.flush()
    return jsonify(ok=True)


@app.route("/")
def index():
    with open(os.path.join(BASE, "index.html"), encoding="utf-8") as f:
        return Response(f.read(), mimetype="text/html")


@app.route("/api/state")
def state():
    try:
        return jsonify(build_state(request.args.get("run_id", "")))
    except KeyError as e:
        return jsonify(error=f"Environment variable missing: {e}"), 500
    except Exception as e:
        return jsonify(error=f"{type(e).__name__}: {e}"), 500

