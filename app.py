from flask import Flask, request, jsonify, send_from_directory
import pandas as pd
import os

app = Flask(__name__, static_folder="static")
UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

state = {"fixably": None, "actual": None, "stockedout": None}

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

if __name__ == "__main__":
    app.run(debug=True, port=5050)
