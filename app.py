import io
import re
import os

import pandas as pd
from flask import Flask, jsonify, render_template, request

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # 16 MB


# ── Flexible column detection ────────────────────────────────────────────────
ALIASES = {
    "Bid_ID":        ["bid_id", "bid id", "bidid", "id", "ref", "reference",
                      "quote_id", "quoteid", "rfq", "tender_id"],
    "Supplier":      ["supplier", "vendor", "company", "bidder", "contractor",
                      "supplier_name", "vendor_name", "firm"],
    "Category":      ["category", "cat", "type", "group", "division",
                      "department", "commodity"],
    "Product":       ["product", "item", "description", "material", "goods",
                      "service", "item_description", "product_name"],
    "Quantity":      ["quantity", "qty", "units", "volume", "count",
                      "amount", "no_of_units", "num"],
    "Unit_Price":    ["unit_price", "unit price", "price", "rate", "cost",
                      "unit_cost", "per_unit", "unitprice", "unit rate"],
    "Total_Amount":  ["total_amount", "total amount", "total", "total_cost",
                      "grand_total", "subtotal", "line_total", "ext_price"],
    "Delivery_Days": ["delivery_days", "delivery days", "delivery",
                      "lead_time", "leadtime", "days", "turnaround",
                      "ship_days", "eta_days"],
    "Proposal_Type": ["proposal_type", "proposal type", "bid_type",
                      "proposal", "type", "quote_type"],
}


def _norm(s: str) -> str:
    return re.sub(r"[\s_\-]+", "_", s.strip().lower())


def detect_columns(df: pd.DataFrame) -> dict:
    """Return {field: original_col_name} for every matched field."""
    df_cols = {_norm(c): c for c in df.columns}
    mapping = {}
    for field, aliases in ALIASES.items():
        for alias in aliases:
            if _norm(alias) in df_cols:
                mapping[field] = df_cols[_norm(alias)]
                break
    return mapping


def clean_number(val) -> float:
    try:
        return float(re.sub(r"[£$,\s]", "", str(val)))
    except (ValueError, TypeError):
        return 0.0


def parse_df(df: pd.DataFrame, col_map: dict) -> list:
    records = []
    for i, row in df.iterrows():
        def get(field, default=""):
            col = col_map.get(field)
            return row[col] if col else default

        unit_price = clean_number(get("Unit_Price", 0))
        quantity   = clean_number(get("Quantity", 1)) or 1
        total      = clean_number(get("Total_Amount", 0))
        if total == 0 and unit_price > 0:
            total = unit_price * quantity

        records.append({
            "Bid_ID":        str(get("Bid_ID", f"BID{i+1:03d}")) or f"BID{i+1:03d}",
            "Supplier":      str(get("Supplier", "Unknown")),
            "Category":      str(get("Category", "General")),
            "Product":       str(get("Product", f"Item {i+1}")),
            "Quantity":      quantity,
            "Unit_Price":    unit_price,
            "Total_Amount":  total,
            "Delivery_Days": int(clean_number(get("Delivery_Days", 0))),
            "Proposal_Type": str(get("Proposal_Type", "Original")).strip() or "Original",
        })
    return [r for r in records if r["Supplier"] not in ("Unknown", "", "nan")]


def read_csv_content(content: bytes) -> pd.DataFrame:
    text = content.decode("utf-8-sig", errors="replace")
    return pd.read_csv(io.StringIO(text))


def read_any_table_file(f) -> pd.DataFrame:
    filename = f.filename or ""
    content = f.read()
    ext = os.path.splitext(filename.lower())[1]

    if ext in {".xls", ".xlsx", ".xlsm", ".xlsb", ".ods"}:
        return pd.read_excel(io.BytesIO(content))

    if ext == ".csv" or not ext:
        return read_csv_content(content)

    try:
        return read_csv_content(content)
    except Exception as csv_exc:
        try:
            return pd.read_excel(io.BytesIO(content))
        except Exception as excel_exc:
            raise ValueError(
                f"Could not parse file as CSV or Excel. CSV error: {csv_exc}; Excel error: {excel_exc}"
            ) from excel_exc


# ── Routes ───────────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    f = request.files["file"]

    try:
        df = read_any_table_file(f)
    except Exception as exc:
        return jsonify({"error": f"Could not parse file: {exc}"}), 400

    df.dropna(how="all", inplace=True)
    col_map = detect_columns(df)

    if "Supplier" not in col_map and "Product" not in col_map:
        # Return preview so the user can remap manually
        return jsonify({
            "error": "no_match",
            "columns": list(df.columns),
            "preview": df.head(3).fillna("").to_dict(orient="records"),
        }), 422

    records = parse_df(df, col_map)
    detected = {k: v for k, v in col_map.items()}

    return jsonify({
        "records":          records,
        "columns_detected": detected,
        "total_rows":       len(records),
        "unmatched_cols":   [c for c in df.columns if c not in col_map.values()],
    })


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
