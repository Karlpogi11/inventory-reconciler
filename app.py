from flask import Flask, request, jsonify, send_from_directory, send_file
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import pandas as pd
import os, io, json

app = Flask(__name__, static_folder="static")
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

state = {"fixably": None, "actual": None, "stockedout": None}
site_state = {}

NORM = lambda s: str(s).strip().upper().replace(" ", "")

def _find_sheet(names, *keywords):
    for n in names:
        u = n.upper().replace(" ", "")
        if all(k in u for k in keywords):
            return n
    return None

def _header_row(ws, key):
    for row in ws.iter_rows(min_row=1, max_row=10):
        for c in row:
            if c.value and str(c.value).strip().upper().startswith(key):
                return c.row
    return 1

def read_file(path):
    if path.endswith(".xlsx"):
        return pd.read_excel(path, dtype=str).fillna("")
    return pd.read_csv(path, dtype=str).fillna("")

def norm(series):
    return series.str.replace(r'\s+', '', regex=True).str.upper()

@app.route("/")
def index():
    return send_from_directory("static", "index.html")

@app.route("/upload", methods=["POST"])
def upload():
    kind = request.form["kind"]  # "fixably", "actual", "stockedout"
    f = request.files["file"]
    path = os.path.join(UPLOAD_DIR, f"{kind}_{f.filename}")
    f.save(path)
    df = read_file(path)
    state[kind] = {"path": path, "columns": list(df.columns), "df": df}
    return jsonify({"columns": list(df.columns), "preview": df.head(3).to_dict(orient="records")})

@app.route("/step1", methods=["POST"])
def step1():
    body = request.json
    scanned = set(s.strip().upper() for s in body["serials"] if s.strip())
    m = body["actual_map"]
    df = state["actual"]["df"].rename(columns={m["serial"]: "serial", m["code"]: "code"})
    df["serial"] = norm(df["serial"])
    actual_serials = set(df[df["serial"] != ""]["serial"])
    not_in_actual = sorted(scanned - actual_serials)
    not_scanned = df[df["serial"].isin(actual_serials - scanned)][["serial", "code"]].to_dict(orient="records")
    return jsonify({
        "matched": len(scanned & actual_serials),
        "not_in_actual": not_in_actual,
        "not_scanned": not_scanned,
    })

@app.route("/reconcile", methods=["POST"])
def reconcile():
    body = request.json
    fx_map = body["fixably_map"]
    ac_map = body["actual_map"]

    fx = state["fixably"]["df"].rename(columns={
        fx_map["serial"]: "serial", fx_map["code"]: "code", fx_map["quantity"]: "quantity"
    })
    ac = state["actual"]["df"].rename(columns={
        ac_map["serial"]: "serial", ac_map["code"]: "code", ac_map["quantity"]: "quantity"
    })

    fx["serial"] = norm(fx["serial"])
    ac["serial"] = norm(ac["serial"])

    fx_serials = set(fx[fx["serial"] != ""]["serial"])
    ac_serials = set(ac[ac["serial"] != ""]["serial"])

    # Stocked-out serials: in Fixably but already used — exclude from discrepancy
    so_serials = set()
    if state["stockedout"] and body.get("stockedout_map"):
        so_df = state["stockedout"]["df"].rename(columns={body["stockedout_map"]["serial"]: "serial"})
        so_df["serial"] = norm(so_df["serial"])
        so_serials = set(so_df[so_df["serial"] != ""]["serial"])

    in_actual_not_fixably = sorted(ac_serials - fx_serials)
    in_fixably_not_actual = sorted((fx_serials - ac_serials) - so_serials)
    stocked_out_excluded  = sorted((fx_serials - ac_serials) & so_serials)

    # Quantity mismatch by part code
    def qty_sum(df):
        return df.groupby("code")["quantity"].apply(
            lambda x: sum(int(v) for v in x if str(v).isdigit())
        ).reset_index()

    merged = pd.merge(
        qty_sum(fx).rename(columns={"quantity": "fx_qty"}),
        qty_sum(ac).rename(columns={"quantity": "ac_qty"}),
        on="code", how="outer"
    ).fillna(0)
    merged["fx_qty"] = merged["fx_qty"].astype(int)
    merged["ac_qty"] = merged["ac_qty"].astype(int)
    qty_mismatch = merged[merged["fx_qty"] != merged["ac_qty"]].to_dict(orient="records")

    state["serial_lookup"] = dict(zip(fx["serial"], fx["code"]))

    return jsonify({
        "in_actual_not_fixably": in_actual_not_fixably,
        "in_fixably_not_actual": in_fixably_not_actual,
        "stocked_out_excluded": stocked_out_excluded,
        "qty_mismatch": qty_mismatch,
        "summary": {
            "fx_total": len(fx_serials),
            "ac_total": len(ac_serials),
            "matched": len(fx_serials & ac_serials),
            "_debug": {
                "fx_col": fx_map["serial"], "fx_sample": sorted(fx_serials)[:5],
                "ac_col": ac_map["serial"], "ac_sample": sorted(ac_serials)[:5],
            }
        }
    })

@app.route("/scan", methods=["POST"])
def scan():
    serial = request.json.get("serial", "").strip().upper()
    lookup = state.get("serial_lookup", {})
    code = lookup.get(serial)
    return jsonify({"serial": serial, "found": code is not None, "code": code or ""})

@app.route("/site/upload", methods=["POST"])
def site_upload():
    f = request.files["file"]
    path = os.path.join(UPLOAD_DIR, f"site_{f.filename}")
    f.save(path)
    wb = load_workbook(path, read_only=True)
    names = wb.sheetnames
    master = _find_sheet(names, "MASTER")
    total = _find_sheet(names, "TOTAL", "PART")
    outtake = _find_sheet(names, "OUTTAKE")
    if not (master and total and outtake):
        return jsonify({"error": f"Could not detect all 3 sheets. Found: {names}"}), 400
    ml = wb[master]
    hdr = [str(c.value).strip().lower() if c.value else "" for c in next(ml.iter_rows(max_row=1))]
    ci = hdr.index("serial number") if "serial number" in hdr else 3
    serials = sum(1 for r in ml.iter_rows(min_row=2, values_only=True) if r[ci])
    site_state.update({"path": path, "master": master, "total": total, "outtake": outtake})
    return jsonify({"file": f.filename, "sheets": {"masterlist": master, "total_parts": total, "outtake": outtake},
                    "masterlist_serials": serials})

@app.route("/site/generate", methods=["POST"])
def site_generate():
    if not site_state.get("path"):
        return jsonify({"error": "Upload a site workbook first."}), 400
    scanned_raw = [s.strip() for s in request.json["serials"] if s.strip()]
    scanned = {NORM(s): s for s in scanned_raw}
    scanned_set = set(scanned)

    wb = load_workbook(site_state["path"])
    ml, tp, out = wb[site_state["master"]], wb[site_state["total"]], wb[site_state["outtake"]]

    rows = list(ml.iter_rows(values_only=True))
    hdr = [str(h).strip().lower() if h else "" for h in rows[0]]
    gi = lambda k, d: hdr.index(k) if k in hdr else d
    ci_code, ci_name, ci_ser, ci_qty = gi("code", 0), gi("name", 1), gi("serial number", 3), gi("quantity", 6)

    parts, ml_serials = {}, set()
    for r in rows[1:]:
        code = r[ci_code]
        if code is None:
            continue
        p = parts.setdefault(code, {"name": r[ci_name], "scanned": [], "unscanned": [], "nonser": 0})
        sn = r[ci_ser]
        if sn:
            ml_serials.add(NORM(sn))
            (p["scanned"] if NORM(sn) in scanned_set else p["unscanned"]).append(str(sn).strip())
        else:
            q = r[ci_qty] or 0
            p["nonser"] += int(q) if str(q).strip().lstrip("-").isdigit() else 0
    new_serials = [scanned[n] for n in scanned_set - ml_serials]

    # --- Total Parts: on-hand only, recomputed count, + Serial Numbers column ---
    thr = _header_row(tp, "PART")
    tp.cell(row=thr, column=4, value="Serial Numbers")
    if tp.max_row > thr:
        tp.delete_rows(thr + 1, tp.max_row - thr)
    kept = 0
    r = thr + 1
    for code, p in parts.items():
        ser_total = len(p["scanned"]) + len(p["unscanned"])
        if ser_total:
            count, serials_str = len(p["scanned"]), ", ".join(p["scanned"])
        else:
            count, serials_str = p["nonser"], ""
        if count <= 0:
            continue
        tp.cell(r, 1, code); tp.cell(r, 2, p["name"]); tp.cell(r, 3, count); tp.cell(r, 4, serials_str)
        r += 1; kept += 1
    for s in new_serials:
        tp.cell(r, 1, ""); tp.cell(r, 2, "NEW STOCK"); tp.cell(r, 3, 1); tp.cell(r, 4, s)
        r += 1

    # --- FOR OUTTAKE: one row per unscanned serial, remark = serial ---
    ohr = _header_row(out, "PART")
    if out.max_row > ohr:
        out.delete_rows(ohr + 1, out.max_row - ohr)
    r = ohr + 1
    outtake_n = 0
    for code, p in parts.items():
        for sn in p["unscanned"]:
            out.cell(r, 1, code); out.cell(r, 2, p["name"]); out.cell(r, 3, sn)
            r += 1; outtake_n += 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    name = os.path.basename(site_state["path"]).replace("site_", "").replace(".xlsx", "")
    resp = send_file(buf, as_attachment=True, download_name=f"{name}_reconciled.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    resp.headers["X-Stats"] = json.dumps({"on_hand_parts": kept, "new_stock": len(new_serials), "outtake": outtake_n})
    return resp

@app.route("/site/annotate", methods=["POST"])
def site_annotate():
    if not site_state.get("path"):
        return jsonify({"error": "Upload a site workbook first."}), 400
    scanned_raw = [s.strip() for s in request.json["serials"] if s.strip()]
    scanned = {NORM(s): s for s in scanned_raw}
    scanned_set = set(scanned)

    wb = load_workbook(site_state["path"])
    ml = wb[site_state["master"]]
    rows = list(ml.iter_rows(values_only=True))
    hdr = [str(h).strip().lower() if h else "" for h in rows[0]]
    gi = lambda k, d: hdr.index(k) if k in hdr else d
    ci_ser, ci_order = gi("serial number", 3), (hdr.index("order") if "order" in hdr else -1)
    rem_col = len(rows[0]) + 1

    USED = PatternFill("solid", fgColor="FFD9A0")   # orange — believed used
    NEW = PatternFill("solid", fgColor="FFF275")    # yellow — not in export
    ml.cell(1, rem_col, "Remarks")

    used_n = 0
    ocn_filled = 0
    ml_serials = set()
    for i, r in enumerate(rows[1:], start=2):
        sn = r[ci_ser]
        if not sn:
            continue
        ml_serials.add(NORM(sn))
        if NORM(sn) not in scanned_set:  # in export but not on hand -> believed used
            ocn = r[ci_order] if ci_order >= 0 and r[ci_order] else None
            ml.cell(i, rem_col, ocn if ocn else "USED - enter OCN")
            if ocn:
                ocn_filled += 1
            for c in range(1, rem_col + 1):
                ml.cell(i, c).fill = USED
            used_n += 1

    # scanned serials not in the export -> append + highlight yellow
    new_serials = [scanned[n] for n in scanned_set - ml_serials]
    r = ml.max_row + 1
    for s in new_serials:
        ml.cell(r, ci_ser + 1, s)
        ml.cell(r, rem_col, "NOT IN EXPORT")
        for c in range(1, rem_col + 1):
            ml.cell(r, c).fill = NEW
        r += 1

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    name = os.path.basename(site_state["path"]).replace("site_", "").replace(".xlsx", "")
    resp = send_file(buf, as_attachment=True, download_name=f"{name}_annotated.xlsx",
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    resp.headers["X-Stats"] = json.dumps({"believed_used": used_n, "not_in_export": len(new_serials),
                                          "ocn_source": f"{ocn_filled} auto from Order, {used_n - ocn_filled} manual"})
    return resp

if __name__ == "__main__":
    app.run(debug=True, port=5050)
