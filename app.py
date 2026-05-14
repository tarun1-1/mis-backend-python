from flask import Flask, request, jsonify, make_response
from flask_cors import CORS
from openpyxl import load_workbook
import os
import json
import datetime
import pytz

app = Flask(__name__)
CORS(app)

DATA_FILE  = "mis_data.json"
AUDIT_FILE = "audit.json"
IST        = pytz.timezone("Asia/Kolkata")

# ── PIN registry ──
PINS = {
    "1212": "Tarun",
    "7890": "Dipanshu",
}

# ── helpers ──
def now_ist():
    return datetime.datetime.now(IST).strftime("%d %b %Y, %I:%M %p IST")

def load_json(path):
    if os.path.exists(path):
        with open(path) as f:
            return json.load(f)
    return None

def save_json(path, data):
    with open(path, "w") as f:
        json.dump(data, f)

# ── Excel parser ──
def parse_excel(file_path):
    wb = load_workbook(file_path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    header = rows[0]
    dates = []
    for i in range(1, len(header)):
        val = header[i]
        if val is None:
            continue
        if isinstance(val, datetime.datetime):
            dates.append({"col": i, "label": val.strftime("%-d %b %Y")})
        elif isinstance(val, str) and val.strip():
            dates.append({"col": i, "label": val.strip()})

    def get_row(label):
        for r in rows:
            if r[0] and label.lower() in str(r[0]).lower():
                return r
        return None

    def val(r, col):
        if r is None:
            return 0
        v = r[col]
        if v is None:
            return 0
        try:
            return float(v)
        except:
            return 0

    rTotal = get_row('total ai calls')
    rConn  = get_row('ai connected')
    rNMC   = get_row('ai nmc')
    rMFC   = get_row('ai meaning')
    rAiInt = get_row('ai interested')
    rAiCB  = get_row('ai call back')

    tai_idx = next((i for i, r in enumerate(rows) if r[0] and 'tele team (ai int' in str(r[0]).lower()), -1)
    tcb_idx = next((i for i, r in enumerate(rows) if r[0] and 'tele team (call back' in str(r[0]).lower()), -1)
    doc_idx = next((i for i, r in enumerate(rows) if r[0] and 'total documents' in str(r[0]).lower()), -1)

    dkeys = ['Interested', 'Call Back', 'Not Interested', 'RNR', 'Not Elligible']

    def build_day(col, label):
        connected = val(rConn, col)
        nmc       = val(rNMC, col)
        mfc_raw   = val(rMFC, col)
        mfc       = mfc_raw if mfc_raw > 0 else max(0, connected - nmc)

        ti = {}
        tc = {}
        for i, k in enumerate(dkeys):
            ti[k] = val(rows[tai_idx + i + 1], col) if tai_idx >= 0 else 0
            tc[k] = val(rows[tcb_idx + i + 1], col) if tcb_idx >= 0 else 0

        doc_total  = (ti.get('Interested', 0) + tc.get('Interested', 0))
        collected  = val(rows[doc_idx + 1], col) if doc_idx >= 0 else 0
        partial    = val(rows[doc_idx + 2], col) if doc_idx >= 0 else 0
        wip        = val(rows[doc_idx + 3], col) if doc_idx >= 0 else 0
        dropout    = val(rows[doc_idx + 4], col) if doc_idx >= 0 else 0
        rejected   = val(rows[doc_idx + 5], col) if doc_idx >= 0 else 0
        disbursed  = val(rows[doc_idx + 6], col) if doc_idx >= 0 else 0

        return {
            "label":     label,
            "totalAI":   int(val(rTotal, col)),
            "connected": int(connected),
            "nmc":       int(nmc),
            "mfc":       int(mfc),
            "aiInt":     int(val(rAiInt, col)),
            "aiCB":      int(val(rAiCB, col)),
            "ti":        {k: int(ti[k]) for k in dkeys},
            "tc":        {k: int(tc[k]) for k in dkeys},
            "doc": {
                "docTotal":  int(doc_total),
                "collected": int(collected),
                "partial":   int(partial),
                "wip":       int(wip),
                "dropout":   int(dropout),
                "rejected":  int(rejected),
                "disbursed": int(disbursed),
            }
        }

    days = [build_day(d["col"], d["label"]) for d in dates]

    def sum_days(days_list):
        if not days_list:
            return build_day(dates[0]["col"], "MTD")
        def s(key):   return sum(d[key] for d in days_list)
        def sti(k):   return sum(d["ti"][k] for d in days_list)
        def stc(k):   return sum(d["tc"][k] for d in days_list)
        def sd(k):    return sum(d["doc"][k] for d in days_list)
        ti = {k: int(sti(k)) for k in dkeys}
        tc = {k: int(stc(k)) for k in dkeys}
        return {
            "label":     "MTD",
            "totalAI":   int(s("totalAI")),
            "connected": int(s("connected")),
            "nmc":       int(s("nmc")),
            "mfc":       int(s("mfc")),
            "aiInt":     int(s("aiInt")),
            "aiCB":      int(s("aiCB")),
            "ti": ti, "tc": tc,
            "doc": {
                "docTotal":  int(sd("docTotal")),
                "collected": int(sd("collected")),
                "partial":   int(sd("partial")),
                "wip":       int(sd("wip")),
                "dropout":   int(sd("dropout")),
                "rejected":  int(sd("rejected")),
                "disbursed": int(sd("disbursed")),
            }
        }

    mtd      = sum_days(days)
    last_day = days[-1] if days else None

    return {
        "generated": now_ist(),
        "uploadedBy": "",          # filled at upload time
        "mtd":      mtd,
        "lastDay":  last_day,
        "allDays":  days,
    }

# ── Validation ──
def validate(days):
    """Returns list of error strings. Empty list = pass."""
    errors = []
    for d in days:
        lbl    = d["label"]
        total  = d["totalAI"]
        conn   = d["connected"]
        nmc    = d["nmc"]
        mfc    = d["mfc"]
        ai_int = d["aiInt"]
        ai_cb  = d["aiCB"]

        # Negative numbers
        for field, fval in [("Total AI Calls", total), ("AI Connected", conn),
                             ("AI NMC", nmc), ("MFC", mfc),
                             ("AI Interested", ai_int), ("AI Callback", ai_cb)]:
            if fval < 0:
                errors.append(f"{lbl}: {field} is negative ({fval})")

        doc = d["doc"]
        for field, fval in [("Collected", doc["collected"]), ("Partial", doc["partial"]),
                             ("WIP", doc["wip"]), ("Dropout", doc["dropout"]),
                             ("Rejected", doc["rejected"]), ("Disbursed", doc["disbursed"])]:
            if fval < 0:
                errors.append(f"{lbl}: Doc field {field} is negative ({fval})")

        # Math checks
        if conn > total:
            errors.append(f"{lbl}: AI Connected ({conn}) > Total AI Calls ({total})")
        if nmc > conn:
            errors.append(f"{lbl}: AI NMC ({nmc}) > AI Connected ({conn})")
        if mfc > conn:
            errors.append(f"{lbl}: MFC ({mfc}) > AI Connected ({conn})")
        if ai_int > mfc:
            errors.append(f"{lbl}: AI Interested ({ai_int}) > MFC ({mfc})")
        if ai_cb > mfc:
            errors.append(f"{lbl}: AI Callback ({ai_cb}) > MFC ({mfc})")
        if (ai_int + ai_cb) > mfc:
            errors.append(f"{lbl}: AI Interested + Callback ({ai_int+ai_cb}) > MFC ({mfc})")

        # NMC + MFC within 15% tolerance of AI Connected
        if conn > 0:
            diff = abs((nmc + mfc) - conn)
            tolerance = conn * 0.15
            if diff > tolerance:
                errors.append(
                    f"{lbl}: NMC ({nmc}) + MFC ({mfc}) = {nmc+mfc}, "
                    f"but AI Connected = {conn}. "
                    f"Difference {diff:.0f} exceeds 15% tolerance ({tolerance:.0f}). "
                    f"Possible data entry error."
                )

    return errors

# ── Diff detection ──
def compute_diff(old_data, new_data):
    """Returns human-readable list of changes."""
    changes = []
    if not old_data:
        return ["First upload — no previous data to compare"]

    old_days = {d["label"]: d for d in old_data.get("allDays", [])}
    new_days = {d["label"]: d for d in new_data.get("allDays", [])}

    old_labels = set(old_days.keys())
    new_labels = set(new_days.keys())

    added   = new_labels - old_labels
    removed = old_labels - new_labels

    for lbl in sorted(added):
        changes.append(f"New date added: {lbl}")
    for lbl in sorted(removed):
        changes.append(f"Date removed: {lbl}")

    # Value changes in common dates
    FIELDS = [
        ("totalAI","Total AI Calls"), ("connected","AI Connected"),
        ("nmc","AI NMC"), ("mfc","MFC"),
        ("aiInt","AI Interested"), ("aiCB","AI Callback"),
    ]
    DOC_FIELDS = [
        ("collected","Collected"), ("partial","Partially Collected"),
        ("wip","WIP"), ("dropout","Dropout"),
        ("rejected","Rejected"), ("disbursed","Disbursed"),
    ]

    for lbl in sorted(old_labels & new_labels):
        od = old_days[lbl]
        nd = new_days[lbl]
        for fkey, fname in FIELDS:
            ov, nv = od.get(fkey, 0), nd.get(fkey, 0)
            if ov != nv:
                changes.append(f"{lbl} → {fname}: {ov} → {nv}")
        for fkey, fname in DOC_FIELDS:
            ov = od.get("doc", {}).get(fkey, 0)
            nv = nd.get("doc", {}).get(fkey, 0)
            if ov != nv:
                changes.append(f"{lbl} → Doc {fname}: {ov} → {nv}")

    if not changes:
        changes.append("No data changes detected")

    return changes

# ── Audit log ──
def append_audit(entry):
    log = load_json(AUDIT_FILE) or []
    log.insert(0, entry)          # newest first
    log = log[:100]               # keep last 100
    save_json(AUDIT_FILE, log)

# ─────────────────────────────────────────
# ROUTES
# ─────────────────────────────────────────

@app.route("/")
def home():
    return "CarmaOne MIS Backend is running"

@app.route("/upload", methods=["POST"])
def upload_file():
    # ── PIN check ──
    pin = request.form.get("pin", "").strip()
    if not pin:
        return jsonify({"error": "PIN is required"}), 403
    uploader = PINS.get(pin)
    if not uploader:
        return jsonify({"error": "Invalid PIN. Please check and try again."}), 403

    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    file = request.files['file']
    if not file.filename.endswith('.xlsx'):
        return jsonify({"error": "Only .xlsx files allowed"}), 400

    tmp = "upload_tmp.xlsx"
    file.save(tmp)

    try:
        new_data = parse_excel(tmp)
        os.remove(tmp)

        # ── Validate ──
        errors = validate(new_data["allDays"])
        if errors:
            audit_entry = {
                "ts":       now_ist(),
                "by":       uploader,
                "status":   "REJECTED",
                "errors":   errors,
                "changes":  [],
            }
            append_audit(audit_entry)
            return jsonify({
                "success": False,
                "validationErrors": errors,
                "message": "Upload rejected due to data errors. See details below."
            }), 422

        # ── Diff ──
        old_data = load_json(DATA_FILE)
        changes  = compute_diff(old_data, new_data)

        # ── Save ──
        new_data["generated"]  = now_ist()
        new_data["uploadedBy"] = uploader
        save_json(DATA_FILE, new_data)

        audit_entry = {
            "ts":      now_ist(),
            "by":      uploader,
            "status":  "SUCCESS",
            "errors":  [],
            "changes": changes,
        }
        append_audit(audit_entry)

        return jsonify({
            "success": True,
            "message": f"File uploaded successfully by {uploader}",
            "changes": changes,
        })

    except Exception as e:
        if os.path.exists(tmp):
            os.remove(tmp)
        return jsonify({"error": str(e)}), 500


@app.route("/data", methods=["GET"])
def get_data():
    if not os.path.exists(DATA_FILE):
        return jsonify({"error": "No data available"}), 404
    with open(DATA_FILE) as f:
        data = json.load(f)
    resp = make_response(jsonify(data))
    resp.headers["Cache-Control"] = "no-store"
    resp.headers["Last-Updated"]  = data.get("generated", "")
    return resp


@app.route("/audit", methods=["GET"])
def get_audit():
    log = load_json(AUDIT_FILE) or []
    resp = make_response(jsonify(log))
    resp.headers["Cache-Control"] = "no-store"
    return resp


@app.route("/status", methods=["GET"])
def get_status():
    data = load_json(DATA_FILE)
    if not data:
        return jsonify({"uploaded": False, "message": "No data uploaded yet"}), 200
    return jsonify({
        "uploaded":   True,
        "generated":  data.get("generated", ""),
        "uploadedBy": data.get("uploadedBy", "Unknown"),
    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
