from flask import Flask, request, jsonify, make_response
from flask_cors import CORS
from openpyxl import load_workbook
import os
import json
import datetime

app = Flask(__name__)
CORS(app)

DATA_FILE = "mis_data.json"

def parse_excel(file_path):
    wb = load_workbook(file_path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))

    # Row 0 = dates (col 0 is None, col 1+ are dates)
    header = rows[0]
    dates = []
    for i in range(1, len(header)):
        val = header[i]
        if val is None:
            continue
        if isinstance(val, datetime.datetime):
            dates.append({"col": i, "label": val.strftime("%-d %b %Y")})
        elif isinstance(val, str):
            dates.append({"col": i, "label": val})

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

    # Tele Team rows by index
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

        doc_total   = (ti.get('Interested', 0) + tc.get('Interested', 0))
        collected   = val(rows[doc_idx + 1], col) if doc_idx >= 0 else 0
        partial     = val(rows[doc_idx + 2], col) if doc_idx >= 0 else 0
        wip         = val(rows[doc_idx + 3], col) if doc_idx >= 0 else 0
        dropout     = val(rows[doc_idx + 4], col) if doc_idx >= 0 else 0
        rejected    = val(rows[doc_idx + 5], col) if doc_idx >= 0 else 0
        disbursed   = val(rows[doc_idx + 6], col) if doc_idx >= 0 else 0

        return {
            "label": label,
            "totalAI": int(val(rTotal, col)),
            "connected": int(connected),
            "nmc": int(nmc),
            "mfc": int(mfc),
            "aiInt": int(val(rAiInt, col)),
            "aiCB": int(val(rAiCB, col)),
            "ti": {k: int(ti[k]) for k in dkeys},
            "tc": {k: int(tc[k]) for k in dkeys},
            "doc": {
                "docTotal": int(doc_total),
                "collected": int(collected),
                "partial": int(partial),
                "wip": int(wip),
                "dropout": int(dropout),
                "rejected": int(rejected),
                "disbursed": int(disbursed)
            }
        }

    days = [build_day(d["col"], d["label"]) for d in dates]

    # MTD = sum of all days
    def sum_days(days):
        if not days:
            return build_day(dates[0]["col"], "MTD")
        def s(key):
            return sum(d[key] for d in days)
        def s_ti(k):
            return sum(d["ti"][k] for d in days)
        def s_tc(k):
            return sum(d["tc"][k] for d in days)
        def s_doc(k):
            return sum(d["doc"][k] for d in days)

        ti = {k: int(s_ti(k)) for k in dkeys}
        tc = {k: int(s_tc(k)) for k in dkeys}
        return {
            "label": "MTD",
            "totalAI": int(s("totalAI")),
            "connected": int(s("connected")),
            "nmc": int(s("nmc")),
            "mfc": int(s("mfc")),
            "aiInt": int(s("aiInt")),
            "aiCB": int(s("aiCB")),
            "ti": ti,
            "tc": tc,
            "doc": {
                "docTotal": int(s_doc("docTotal")),
                "collected": int(s_doc("collected")),
                "partial": int(s_doc("partial")),
                "wip": int(s_doc("wip")),
                "dropout": int(s_doc("dropout")),
                "rejected": int(s_doc("rejected")),
                "disbursed": int(s_doc("disbursed"))
            }
        }

    mtd = sum_days(days)
    last_day = days[-1] if days else None

    return {
        "generated": datetime.datetime.now().strftime("%d %b %Y, %I:%M %p"),
        "mtd": mtd,
        "lastDay": last_day,
        "allDays": days
    }

@app.route("/")
def home():
    return "CarmaOne MIS Backend is running"

@app.route("/upload", methods=["POST"])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    file = request.files['file']
    if not file.filename.endswith('.xlsx'):
        return jsonify({"error": "Only .xlsx files allowed"}), 400

    tmp = "upload_tmp.xlsx"
    file.save(tmp)

    try:
        data = parse_excel(tmp)
        with open(DATA_FILE, 'w') as f:
            json.dump(data, f)
        os.remove(tmp)
        return jsonify({"success": True, "message": "File uploaded and processed successfully"})
    except Exception as e:
        if os.path.exists(tmp):
            os.remove(tmp)
        return jsonify({"error": str(e)}), 500

@app.route("/data", methods=["GET"])
def get_data():
    if not os.path.exists(DATA_FILE):
        return jsonify({"error": "No data available"}), 404
    with open(DATA_FILE, 'r') as f:
        data = json.load(f)
    response = make_response(jsonify(data))
    response.headers["Cache-Control"] = "no-store"
    response.headers["Last-Updated"] = data.get("generated", "")
    return response

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)