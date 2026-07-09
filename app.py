from flask import Flask, request, jsonify, make_response
from flask_cors import CORS
from openpyxl import load_workbook
import os, json, datetime, difflib
import pytz

app = Flask(__name__)
CORS(app)

DATA_FILE   = "mis_data.json"
AUDIT_FILE  = "audit.json"
UNIQUE_FILE        = "unique_set.json"    # AI unique customers
HANDOFF_MOBILE_FILE= "handoff_mobile_set.json"  # mobiles from Handoff_Leads
BLENDED_CID_FILE   = "blended_cid_set.json"     # CIDs matched in blended calls
MANUAL_CID_FILE    = "manual_cid_set.json"      # CIDs in direct tele
IST         = pytz.timezone("Asia/Kolkata")

PINS = {"1212": "Tarun", "7890": "Dipanshu"}

KNOWN_STATUSES = [
    "Disbursed", "Rejected", "Sanctioned", "Case Dropped",
    "Login Pending", "Login", "Hold", "Credit Review"
]

# AI dispositions that indicate call was NOT connected
NOT_CONNECTED_DISPOS = {"NO ANSWER", "NUMBER BUSY", "SWITCHED OFF", "DUPLICATE LEAD", "BLANK"}
VALID_CALLMODES = {
    "predictive", "predictive-blended", "manual",
    "callback", "redial", "outbound", "progressive"
}
VALID_HANDOFF_DISPOS = {"INTERESTED", "CALL BACK"}


# ── Helpers ───────────────────────────────────────────────────────────────────

def now_ist():
    return datetime.datetime.now(IST).strftime("%d %b %Y, %I:%M %p IST")

def load_json(path):
    if os.path.exists(path):
        with open(path) as f:
            return json.load(f)
    return None

def save_json(path, data):
    with open(path, "w") as f:
        json.dump(data, f, ensure_ascii=False)

def load_unique_set():
    data = load_json(UNIQUE_FILE)
    if data and isinstance(data, list):
        return set(data)
    return set()

def save_unique_set(s):
    save_json(UNIQUE_FILE, list(s))

def load_set(path):
    data = load_json(path)
    return set(data) if isinstance(data, list) else set()

def save_set(path, s):
    save_json(path, list(s))

def excel_to_date(val):
    """Convert Excel serial / datetime / date / string → 'YYYY-MM-DD'."""
    if val is None:
        return None
    if isinstance(val, datetime.datetime):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, datetime.date):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, (int, float)):
        try:
            d = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=float(val))
            return d.strftime("%Y-%m-%d")
        except Exception:
            return None
    if isinstance(val, str):
        s = val.strip()
        for fmt in ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d %b %Y", "%d %B %Y"]:
            try:
                return datetime.datetime.strptime(s, fmt).strftime("%Y-%m-%d")
            except Exception:
                continue
    return None

def date_label(iso):
    """'2026-04-25' → '25 Apr 2026'"""
    if not iso:
        return ""
    try:
        return datetime.datetime.strptime(iso, "%Y-%m-%d").strftime("%-d %b %Y")
    except Exception:
        return iso

def normalize_status(raw):
    """Fuzzy-match raw status string to KNOWN_STATUSES.
    Returns (normalized_value, was_fuzzy_matched)."""
    if not raw:
        return ("Unknown", False)
    s = str(raw).strip()
    for k in KNOWN_STATUSES:
        if k.lower() == s.lower():
            return (k, False)
    matches = difflib.get_close_matches(s, KNOWN_STATUSES, n=1, cutoff=0.6)
    if matches:
        return (matches[0], True)
    return (s, False)

def sf(v, d=0.0):
    try:
        return float(v) if v is not None else d
    except Exception:
        return d

def si(v, d=0):
    try:
        return int(float(v)) if v is not None else d
    except Exception:
        return d

def ss(v):
    return str(v).strip() if v is not None else ""


# ── Sheet reader ──────────────────────────────────────────────────────────────

def find_sheet_name(wb, candidates):
    """Return actual sheet name from wb matching any candidate (flexible)."""
    for c in candidates:
        if c in wb.sheetnames:
            return c
    for c in candidates:
        for s in wb.sheetnames:
            if s.lower() == c.lower():
                return s
    for c in candidates:
        for s in wb.sheetnames:
            if c.lower() in s.lower():
                return s
    return None

def read_sheet(wb, sheet_name):
    """Read sheet → list of dicts (column names stripped of whitespace)."""
    if sheet_name not in wb.sheetnames:
        return []
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return []
    # First non-completely-empty row is header
    header, header_idx = None, 0
    for i, r in enumerate(rows):
        if any(c is not None for c in r):
            header = [ss(c) for c in r]
            header_idx = i
            break
    if not header:
        return []
    result = []
    for r in rows[header_idx + 1:]:
        if all(c is None for c in r):
            continue
        row = {header[i]: r[i] for i in range(min(len(header), len(r)))}
        result.append(row)
    return result


# ── Validators ────────────────────────────────────────────────────────────────

def _e(sheet, row, msg):
    return {"sheet": sheet, "row": str(row), "msg": msg}

def _w(sheet, row, msg):
    return {"sheet": sheet, "row": str(row), "msg": msg}

def validate_ai_dump(rows):
    errors, warnings = [], []
    required = ["#", "Call Date", "Batch", "L1 Disposition", "Call Duration"]
    sheet = "AI Dump"
    if not rows:
        return [_e(sheet, "-", "Sheet is empty or not found")], []
    for col in required:
        if col not in rows[0]:
            errors.append(_e(sheet, "Header", f"Missing required column: '{col}'"))
    if errors:
        return errors, warnings
    for i, r in enumerate(rows, 2):
        rid  = ss(r.get("#", "")) or str(i)
        date = excel_to_date(r.get("Call Date"))
        dur  = sf(r.get("Call Duration", 0))
        bat  = ss(r.get("Batch", ""))
        if not bat:
            errors.append(_e(sheet, rid, "Batch is empty"))
        if date is None:
            errors.append(_e(sheet, rid, "Invalid or missing Call Date"))
        if dur < 0:
            errors.append(_e(sheet, rid, f"Call Duration is negative ({dur})"))
    return errors, warnings

def validate_handoff(rows):
    errors, warnings = [], []
    required = ["#", "Call Date", "Batch", "L1 Disposition"]
    sheet = "Handoff Leads"
    if not rows:
        return [_e(sheet, "-", "Sheet is empty or not found")], []
    for col in required:
        if col not in rows[0]:
            errors.append(_e(sheet, "Header", f"Missing required column: '{col}'"))
    if errors:
        return errors, warnings
    for i, r in enumerate(rows, 2):
        rid   = ss(r.get("#", "")) or str(i)
        date  = excel_to_date(r.get("Call Date"))
        dispo = ss(r.get("L1 Disposition", "")).upper()
        if date is None:
            errors.append(_e(sheet, rid, "Invalid or missing Call Date"))
        if dispo not in VALID_HANDOFF_DISPOS:
            errors.append(_e(sheet, rid,
                f"Invalid disposition '{dispo}' — only INTERESTED or CALL BACK allowed"))
    return errors, warnings

def validate_dialer(rows):
    errors, warnings = [], []
    required = ["session_uuid", "init_time", "callmode", "talk_time", "user"]
    sheet = "Dialer Dump"
    if not rows:
        return [_e(sheet, "-", "Sheet is empty or not found")], []
    for col in required:
        if col not in rows[0]:
            errors.append(_e(sheet, "Header", f"Missing required column: '{col}'"))
    if errors:
        return errors, warnings
    for i, r in enumerate(rows, 2):
        uuid     = ss(r.get("session_uuid", ""))
        callmode = ss(r.get("callmode", "")).lower()
        talk     = sf(r.get("talk_time", 0))
        dispo    = ss(r.get("primary_dispo", ""))
        dialed   = ss(r.get("dialed_status", ""))
        date     = excel_to_date(r.get("init_time"))
        if not uuid:
            errors.append(_e(sheet, str(i), "session_uuid is missing"))
        if date is None:
            errors.append(_e(sheet, str(i), "Invalid or missing init_time"))
        if callmode and callmode not in VALID_CALLMODES:
            errors.append(_e(sheet, str(i), f"Unknown callmode '{callmode}'"))
        if talk < 0:
            errors.append(_e(sheet, str(i), f"talk_time is negative ({talk})"))
        if dialed.lower() == "connected" and not dispo:
            warnings.append(_w(sheet, str(i), "Connected call has no disposition"))
    return errors, warnings

def validate_pipeline(rows):
    errors, warnings = [], []
    required = ["Borrower Name", "Lender Name", "Status", "Date of Lead"]
    sheet = "Login Pipeline"
    if not rows:
        return [_e(sheet, "-", "Sheet is empty or not found")], []
    for col in required:
        if col not in rows[0]:
            errors.append(_e(sheet, "Header", f"Missing required column: '{col}'"))
    if errors:
        return errors, warnings
    for i, r in enumerate(rows, 2):
        borrower = ss(r.get("Borrower Name", ""))
        lender   = ss(r.get("Lender Name", ""))
        amt      = r.get("Expected Sanction Amount(cr)")
        date     = excel_to_date(r.get("Date of Lead"))
        officer  = ss(r.get("Officer Name", ""))
        if not borrower:
            errors.append(_e(sheet, str(i), "Borrower Name is missing"))
        if not lender:
            errors.append(_e(sheet, str(i), "Lender Name is missing"))
        if date is None:
            errors.append(_e(sheet, str(i),
                f"Invalid Date of Lead for '{borrower or 'unknown'}'"))
        if amt is not None and sf(amt) < 0:
            errors.append(_e(sheet, str(i),
                f"Expected Sanction Amount is negative for '{borrower}'"))
        if not officer:
            warnings.append(_w(sheet, str(i),
                f"Officer Name missing for '{borrower}'"))
    return errors, warnings


# ── Aggregators ───────────────────────────────────────────────────────────────

def aggregate_ai(rows, processed_dates, existing_unique=None):
    by_day = {}
    added = skipped = 0
    new_unique = set(existing_unique) if existing_unique else set()

    for r in rows:
        date = excel_to_date(r.get("Call Date"))
        if not date:
            continue
        if date in processed_dates:
            skipped += 1
            continue
        dispo    = ss(r.get("L1 Disposition", "")).upper() or "BLANK"
        dur      = sf(r.get("Call Duration", 0))
        cust     = ss(r.get("Customer Number", ""))
        is_conn  = dispo not in NOT_CONNECTED_DISPOS
        if cust:
            new_unique.add(cust)

        if date not in by_day:
            by_day[date] = {"_leads": set(), "total_dials": 0, "connected": 0, "dispositions": {}}
        d = by_day[date]
        d["_leads"].add(cust)
        d["total_dials"] += 1
        if is_conn:
            d["connected"] += 1
        ds = d["dispositions"]
        if dispo not in ds:
            ds[dispo] = {"dials": 0, "connects": 0, "talk_time_sec": 0.0}
        ds[dispo]["dials"]         += 1
        ds[dispo]["connects"]      += (1 if is_conn else 0)
        ds[dispo]["talk_time_sec"] += dur
        added += 1

    result = {}
    for date, d in by_day.items():
        dispos = {}
        for k, v in d["dispositions"].items():
            dispos[k] = {**v,
                "aht_sec": v["talk_time_sec"] / v["connects"] if v["connects"] > 0 else 0.0}
        result[date] = {
            "unique_leads": len(d["_leads"]),
            "total_dials":  d["total_dials"],
            "connected":    d["connected"],
            "dispositions": dispos,
        }
    return result, set(result.keys()), added, skipped, new_unique


def aggregate_handoff(rows, processed_dates, existing_handoff_mobiles=None):
    by_day = {}
    added = skipped = 0
    new_handoff_mobiles = set(existing_handoff_mobiles) if existing_handoff_mobiles else set()

    for r in rows:
        date  = excel_to_date(r.get("Call Date"))
        if not date:
            continue
        if date in processed_dates:
            skipped += 1
            continue
        dispo  = ss(r.get("L1 Disposition", "")).upper()
        batch  = ss(r.get("Batch", "Unknown")) or "Unknown"
        mobile = ss(r.get("Customer Number") or r.get("Mobile") or r.get("Mobile No.") or r.get("Phone") or "")
        if mobile:
            new_handoff_mobiles.add(mobile)

        if date not in by_day:
            by_day[date] = {"interested": 0, "callback": 0, "batches": {}}
        d = by_day[date]
        if dispo == "INTERESTED":
            d["interested"] += 1
        elif dispo == "CALL BACK":
            d["callback"] += 1
        if batch not in d["batches"]:
            d["batches"][batch] = {"interested": 0, "callback": 0}
        if dispo == "INTERESTED":
            d["batches"][batch]["interested"] += 1
        elif dispo == "CALL BACK":
            d["batches"][batch]["callback"]   += 1
        added += 1

    return by_day, set(by_day.keys()), added, skipped, new_handoff_mobiles


def aggregate_dialer(rows, processed_dates, handoff_set=None,
                     existing_blended_cids=None, existing_manual_cids=None):
    """Process all callmodes. predictive-blended matching handoff mobiles → blended.
    Everything else (manual, others, blended not in handoff) → manual/direct tele."""
    blended_by_day = {}
    manual_by_day  = {}
    added = skipped = 0
    new_blended_cids = set(existing_blended_cids) if existing_blended_cids else set()
    new_manual_cids  = set(existing_manual_cids)  if existing_manual_cids  else set()
    h_set = handoff_set or set()

    for r in rows:
        callmode = ss(r.get("callmode", "")).lower()
        date     = excel_to_date(r.get("init_time"))
        if not date:
            continue
        if date in processed_dates:
            skipped += 1
            continue

        cid      = ss(r.get("Customer Number") or r.get("customer_number") or r.get("cid") or "")
        agent    = ss(r.get("user", "")) or "Unknown"
        talk_sec = sf(r.get("talk_time", 0)) * 86400.0
        dispo    = ss(r.get("primary_dispo", "")) or "No Feedback"
        is_conn  = ss(r.get("dialed_status", "")).lower() == "connected"

        # Route: blended on a handoff lead → Stage 2A; everything else → Stage 2B
        if callmode == "predictive-blended" and cid and cid in h_set:
            target = blended_by_day
            if cid: new_blended_cids.add(cid)
        else:
            target = manual_by_day
            if cid: new_manual_cids.add(cid)

        if date not in target:
            target[date] = {"_unique": set(), "total_dials": 0, "connected": 0,
                            "talk_time_sec": 0.0, "dispositions": {}, "agents": {}}
        d = target[date]
        if cid: d["_unique"].add(cid)
        d["total_dials"]   += 1
        d["connected"]     += (1 if is_conn else 0)
        d["talk_time_sec"] += talk_sec
        d["dispositions"][dispo] = d["dispositions"].get(dispo, 0) + 1

        if agent not in d["agents"]:
            d["agents"][agent] = {"_unique": set(), "dials": 0, "connected": 0,
                                   "talk_time_sec": 0.0, "dispositions": {}}
        a = d["agents"][agent]
        if cid: a["_unique"].add(cid)
        a["dials"]        += 1
        a["connected"]    += (1 if is_conn else 0)
        a["talk_time_sec"]+= talk_sec
        a["dispositions"][dispo] = a["dispositions"].get(dispo, 0) + 1
        added += 1

    def finalize(by_day_dict):
        result = {}
        for date, d in by_day_dict.items():
            agents_out = {}
            for ag, av in d["agents"].items():
                agents_out[ag] = {
                    "unique_leads":  len(av["_unique"]),
                    "dials":         av["dials"],
                    "connected":     av["connected"],
                    "talk_time_sec": av["talk_time_sec"],
                    "dispositions":  av["dispositions"],
                }
            result[date] = {
                "unique_leads":  len(d["_unique"]),
                "total_dials":   d["total_dials"],
                "connected":     d["connected"],
                "talk_time_sec": d["talk_time_sec"],
                "dispositions":  d["dispositions"],
                "agents":        agents_out,
            }
        return result

    all_dates = set(blended_by_day.keys()) | set(manual_by_day.keys())
    return (finalize(blended_by_day), finalize(manual_by_day),
            all_dates, added, skipped, new_blended_cids, new_manual_cids)


def process_pipeline(rows, existing_rows):
    """Update pipeline using borrower|lender|date_of_lead as key.
    Returns (updated_list, added, updated, fuzzy_warnings)."""
    pipeline = {}
    for row in existing_rows:
        key = f"{row.get('borrower','')}|{row.get('lender','')}|{row.get('date_of_lead','')}"
        pipeline[key] = row

    added = updated = 0
    fuzzy_warns = []

    for r in rows:
        borrower = ss(r.get("Borrower Name", ""))
        lender   = ss(r.get("Lender Name", ""))
        if not borrower or not lender:
            continue

        date_lead = excel_to_date(r.get("Date of Lead"))
        label_lead = date_label(date_lead) if date_lead else ""

        key = f"{borrower}|{lender}|{label_lead}"

        status_raw  = ss(r.get("Status", ""))
        status_norm, was_fuzzy = normalize_status(status_raw)
        if was_fuzzy:
            fuzzy_warns.append(
                f"Status '{status_raw}' matched to '{status_norm}' "
                f"for {borrower} / {lender}"
            )

        row_data = {
            "date_of_lead":          label_lead,
            "officer":               ss(r.get("Officer Name", "")),
            "borrower":              borrower,
            "mobile":                ss(r.get("Mobile No.", "")),
            "dsa_code":              ss(r.get("DSA Code", "")),
            "lender":                lender,
            "lead_source":           ss(r.get("Lead Source", "")),
            "lead_type":             ss(r.get("Lead Type", "")),
            "product_type":          ss(r.get("Product Type", "")),
            "login_date":            date_label(excel_to_date(r.get("Login Date"))),
            "target_sanction_date":  date_label(excel_to_date(r.get("Target Sanction Date"))),
            "target_disbursal_date": date_label(excel_to_date(r.get("Target Disbursement Date"))),
            "expected_amount_cr":    sf(r.get("Expected Sanction Amount(cr)", 0)),
            "status":                status_norm,
            "status_raw":            status_raw,
            "remarks":               ss(r.get("Remarks", "")),
        }

        if key in pipeline:
            updated += 1
        else:
            added += 1
        pipeline[key] = row_data

    return list(pipeline.values()), added, updated, fuzzy_warns


# ── Response builder ──────────────────────────────────────────────────────────

def _sum_ai(by_day, dates):
    r = {"unique_leads": 0, "total_dials": 0, "connected": 0, "dispositions": {}}
    for d in dates:
        day = by_day.get(d, {})
        r["unique_leads"] += day.get("unique_leads", 0)
        r["total_dials"]  += day.get("total_dials", 0)
        r["connected"]    += day.get("connected", 0)
        for dispo, s in day.get("dispositions", {}).items():
            if dispo not in r["dispositions"]:
                r["dispositions"][dispo] = {"dials": 0, "connects": 0, "talk_time_sec": 0.0}
            r["dispositions"][dispo]["dials"]         += s.get("dials", 0)
            r["dispositions"][dispo]["connects"]      += s.get("connects", 0)
            r["dispositions"][dispo]["talk_time_sec"] += s.get("talk_time_sec", 0.0)
    for dispo in r["dispositions"]:
        s = r["dispositions"][dispo]
        s["aht_sec"] = s["talk_time_sec"] / s["connects"] if s["connects"] > 0 else 0.0
    return r

def _sum_handoff(by_day, dates):
    r = {"interested": 0, "callback": 0, "batches": {}}
    for d in dates:
        day = by_day.get(d, {})
        r["interested"] += day.get("interested", 0)
        r["callback"]   += day.get("callback", 0)
        for b, s in day.get("batches", {}).items():
            if b not in r["batches"]:
                r["batches"][b] = {"interested": 0, "callback": 0}
            r["batches"][b]["interested"] += s.get("interested", 0)
            r["batches"][b]["callback"]   += s.get("callback", 0)
    return r

def _sum_dialer(by_day, dates):
    r = {"unique_leads": 0, "total_dials": 0, "connected": 0, "talk_time_sec": 0.0,
         "dispositions": {}, "agents": {}}
    for d in dates:
        day = by_day.get(d, {})
        r["unique_leads"]  += day.get("unique_leads", 0)
        r["total_dials"]   += day.get("total_dials", 0)
        r["connected"]     += day.get("connected", 0)
        r["talk_time_sec"] += day.get("talk_time_sec", 0.0)
        for dispo, cnt in day.get("dispositions", {}).items():
            r["dispositions"][dispo] = r["dispositions"].get(dispo, 0) + cnt
        for agent, s in day.get("agents", {}).items():
            if agent not in r["agents"]:
                r["agents"][agent] = {
                    "unique_leads": 0, "dials": 0, "connected": 0,
                    "talk_time_sec": 0.0, "dispositions": {}
                }
            a = r["agents"][agent]
            a["unique_leads"] += s.get("unique_leads", 0)
            a["dials"]        += s.get("dials", 0)
            a["connected"]    += s.get("connected", 0)
            a["talk_time_sec"]+= s.get("talk_time_sec", 0.0)
            for dispo, cnt in s.get("dispositions", {}).items():
                a["dispositions"][dispo] = a["dispositions"].get(dispo, 0) + cnt
    return r

def build_response(stored):
    ai_dates  = sorted(stored.get("ai_by_day", {}).keys())
    h_dates   = sorted(stored.get("handoff_by_day", {}).keys())
    b_dates   = sorted(stored.get("dialer_blended_by_day", {}).keys())
    dm_dates  = sorted(stored.get("dialer_manual_by_day", {}).keys())

    def package(sum_fn, by_day, dates):
        mtd = sum_fn(by_day, dates)
        mtd["label"] = "MTD"
        last = None
        if dates:
            last = sum_fn(by_day, [dates[-1]])
            last["label"] = date_label(dates[-1])
        all_days = []
        for d in dates:
            day = sum_fn(by_day, [d])
            day["label"] = date_label(d)
            all_days.append(day)
        return {"mtd": mtd, "lastDay": last, "allDays": all_days}

    # Compute handoff not dialled (mobiles handed off but never dialled in blended)
    handoff_set  = load_set(HANDOFF_MOBILE_FILE)
    blended_set  = load_set(BLENDED_CID_FILE)
    not_dialled  = len(handoff_set - blended_set)

    return {
        "generated":            stored.get("generated", ""),
        "uploadedBy":           stored.get("uploadedBy", ""),
        "totalUniqueLeads":     stored.get("total_unique_leads", 0),
        "totalBlenUnique":      stored.get("total_blended_unique", 0),
        "totalManualUnique":    stored.get("total_manual_unique", 0),
        "handoffNotDialled":    not_dialled,
        "ai":           package(_sum_ai,      stored.get("ai_by_day", {}),              ai_dates),
        "handoff":      package(_sum_handoff, stored.get("handoff_by_day", {}),         h_dates),
        "dialerBlended":package(_sum_dialer,  stored.get("dialer_blended_by_day", {}),  b_dates),
        "dialerManual": package(_sum_dialer,  stored.get("dialer_manual_by_day", {}),   dm_dates),
        "pipeline":     {"rows": stored.get("pipeline_rows", [])},
    }


# ── Audit ─────────────────────────────────────────────────────────────────────

def append_audit(entry):
    log = load_json(AUDIT_FILE) or []
    log.insert(0, entry)
    save_json(AUDIT_FILE, log[:100])


# ── Routes ────────────────────────────────────────────────────────────────────

@app.route("/")
def home():
    return "CarmaOne MIS Backend is running"


@app.route("/upload", methods=["POST"])
def upload_file():
    # PIN
    pin = request.form.get("pin", "").strip()
    if not pin:
        return jsonify({"error": "PIN is required"}), 403
    uploader = PINS.get(pin)
    if not uploader:
        return jsonify({"error": "Invalid PIN. Please check and try again."}), 403

    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    file = request.files["file"]
    if not file.filename.endswith(".xlsx"):
        return jsonify({"error": "Only .xlsx files allowed"}), 400

    tmp = "upload_tmp.xlsx"
    file.save(tmp)

    try:
        wb = load_workbook(tmp, read_only=True, data_only=True)

        # Locate sheets flexibly
        sn_ai     = find_sheet_name(wb, ["AI_Dump", "AI Dump", "AIDump"])
        sn_hand   = find_sheet_name(wb, ["Handoff_Leads", "Handoff Leads", "Handoff"])
        sn_dial   = find_sheet_name(wb, ["Dialer_Dump", "Dialer Dump", "DialerDump"])
        sn_pipe   = find_sheet_name(wb, ["Login Pipeline", "Login_Pipeline", "Pipeline"])

        missing_hard = [n for n, s in [
            ("AI_Dump", sn_ai), ("Handoff_Leads", sn_hand), ("Dialer_Dump", sn_dial)
        ] if not s]

        if missing_hard:
            wb.close(); os.remove(tmp)
            return jsonify({"error": f"Missing required sheets: {', '.join(missing_hard)}. "
                            f"Found: {', '.join(wb.sheetnames)}"}), 422

        pipeline_missing = not sn_pipe

        ai_rows   = read_sheet(wb, sn_ai)
        h_rows    = read_sheet(wb, sn_hand)
        d_rows    = read_sheet(wb, sn_dial)
        p_rows    = read_sheet(wb, sn_pipe) if sn_pipe else []
        wb.close()
        os.remove(tmp)

        # Validate all sheets
        all_errors, all_warnings = [], []
        for errs, warns in [
            validate_ai_dump(ai_rows),
            validate_handoff(h_rows),
            validate_dialer(d_rows),
        ]:
            all_errors.extend(errs)
            all_warnings.extend(warns)

        if p_rows:
            errs, warns = validate_pipeline(p_rows)
            all_errors.extend(errs)
            all_warnings.extend(warns)
        elif pipeline_missing:
            all_warnings.append(_w("Login Pipeline", "-",
                "Sheet not found in this upload — pipeline data unchanged"))

        if all_errors:
            append_audit({
                "ts": now_ist(), "by": uploader, "status": "REJECTED",
                "errors":   [f"[{e['sheet']}] Row {e['row']}: {e['msg']}" for e in all_errors],
                "warnings": [], "changes": {}
            })
            return jsonify({
                "success": False,
                "validationErrors": all_errors,
                "message": "Upload rejected due to data errors. See details below."
            }), 422

        # Load existing stored data
        stored = load_json(DATA_FILE) or {
            "processed_dates":         {"ai_dump": [], "handoff": [], "dialer": []},
            "pipeline_rows":           [],
            "ai_by_day":               {},
            "handoff_by_day":          {},
            "dialer_blended_by_day":   {},
            "dialer_manual_by_day":    {},
        }
        # Migrate old processed_dates key if needed
        if "dialer_manual" in stored.get("processed_dates", {}):
            stored["processed_dates"]["dialer"] = stored["processed_dates"].pop("dialer_manual", [])

        pd_ai  = set(stored["processed_dates"].get("ai_dump", []))
        pd_h   = set(stored["processed_dates"].get("handoff", []))
        pd_d   = set(stored["processed_dates"].get("dialer", []))

        # Aggregate AI
        existing_unique = load_unique_set()
        ai_new, ai_dates, ai_add, ai_skip, new_unique = aggregate_ai(ai_rows, pd_ai, existing_unique)

        # Aggregate Handoff (also tracks mobile numbers for cross-referencing)
        existing_handoff_mobiles = load_set(HANDOFF_MOBILE_FILE)
        h_new, h_dates, h_add, h_skip, new_handoff_mobiles = aggregate_handoff(
            h_rows, pd_h, existing_handoff_mobiles)

        # Aggregate Dialer (split blended vs manual by handoff mobile match)
        handoff_set = new_handoff_mobiles  # Use updated set
        existing_blended_cids = load_set(BLENDED_CID_FILE)
        existing_manual_cids  = load_set(MANUAL_CID_FILE)
        b_new, dm_new, d_dates, d_add, d_skip, new_blended_cids, new_manual_cids = aggregate_dialer(
            d_rows, pd_d, handoff_set, existing_blended_cids, existing_manual_cids)

        # Pipeline
        p_list, p_add, p_upd, p_warns = process_pipeline(
            p_rows, stored.get("pipeline_rows", []))

        # Merge
        stored["ai_by_day"].update(ai_new)
        save_unique_set(new_unique)
        stored["total_unique_leads"] = len(new_unique)

        stored["handoff_by_day"].update(h_new)
        save_set(HANDOFF_MOBILE_FILE, new_handoff_mobiles)

        stored["dialer_blended_by_day"].update(b_new)
        stored["dialer_manual_by_day"].update(dm_new)
        save_set(BLENDED_CID_FILE, new_blended_cids)
        save_set(MANUAL_CID_FILE,  new_manual_cids)
        stored["total_blended_unique"] = len(new_blended_cids)
        stored["total_manual_unique"]  = len(new_manual_cids)

        if p_rows:
            stored["pipeline_rows"] = p_list
        stored["processed_dates"]["ai_dump"] = sorted(pd_ai | ai_dates)
        stored["processed_dates"]["handoff"]  = sorted(pd_h  | h_dates)
        stored["processed_dates"]["dialer"]   = sorted(pd_d  | d_dates)
        stored["generated"]  = now_ist()
        stored["uploadedBy"] = uploader
        save_json(DATA_FILE, stored)

        changes = {
            "AI Dump":        f"{ai_add} rows added, {ai_skip} skipped (already processed)",
            "Handoff Leads":  f"{h_add} rows added, {h_skip} skipped",
            "Dialer Dump":    f"{dm_add} manual calls added, {dm_skip} skipped",
            "Login Pipeline": (f"{p_add} new entries, {p_upd} updated"
                               if p_rows else "Not in this upload — unchanged"),
        }

        warn_strs = ([f"[{w['sheet']}] Row {w['row']}: {w['msg']}" for w in all_warnings]
                     + [f"[Login Pipeline] {w}" for w in p_warns])

        append_audit({
            "ts": now_ist(), "by": uploader, "status": "SUCCESS",
            "errors": [], "warnings": warn_strs, "changes": changes
        })

        return jsonify({
            "success":  True,
            "message":  f"File uploaded successfully by {uploader}",
            "changes":  changes,
            "warnings": warn_strs,
        })

    except Exception as e:
        if os.path.exists(tmp):
            os.remove(tmp)
        return jsonify({"error": str(e)}), 500


@app.route("/data", methods=["GET"])
def get_data():
    stored = load_json(DATA_FILE)
    if not stored:
        return jsonify({"error": "No data available"}), 404
    resp = make_response(jsonify(build_response(stored)))
    resp.headers["Cache-Control"] = "no-store"
    return resp


@app.route("/audit", methods=["GET"])
def get_audit():
    log = load_json(AUDIT_FILE) or []
    resp = make_response(jsonify(log))
    resp.headers["Cache-Control"] = "no-store"
    return resp


@app.route("/status", methods=["GET"])
def get_status():
    stored = load_json(DATA_FILE)
    if not stored:
        return jsonify({"uploaded": False, "message": "No data uploaded yet"})
    return jsonify({
        "uploaded":   True,
        "generated":  stored.get("generated", ""),
        "uploadedBy": stored.get("uploadedBy", ""),
    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
