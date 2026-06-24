import io
import json
import sqlite3
import zipfile
import re
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import streamlit as st

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image

APP_DIR = Path(__file__).parent
LOGO_PATH = APP_DIR / "western_excavation_logo.png"
DATA_DIR = Path.home() / "Western_Payroll_Data"
DATA_DIR.mkdir(parents=True, exist_ok=True)
DB_PATH = DATA_DIR / "western_payroll.db"
YTD_BACKUP_PATH = DATA_DIR / "ytd_totals_backup.json"
BUSINESS_NAME = "Western Excavation"
BUSINESS_ADDRESS = "1546 Austinville Road, Max Meadows, VA 24360"
BUSINESS_PHONE = "276-613-3854"
DEFAULT_TRANSFER_FEE = 35.00

SOCIAL_SECURITY_RATE = 0.062
SOCIAL_SECURITY_WAGE_BASE = 176100.00
MEDICARE_RATE = 0.0145
ADDITIONAL_MEDICARE_RATE = 0.009

PAY_PERIODS = {
    "Weekly": 52,
    "Biweekly": 26,
    "Semimonthly": 24,
    "Monthly": 12,
}

# Simplified annual federal income tax brackets used for paycheck estimating.
# This is intended to behave like PaycheckCity-style withholding inputs, not to replace official payroll software.
FEDERAL_BRACKETS = {
    "Single": [
        (0, 0.10), (11925, 0.12), (48475, 0.22), (103350, 0.24),
        (197300, 0.32), (250525, 0.35), (626350, 0.37),
    ],
    "Married Filing Jointly": [
        (0, 0.10), (23850, 0.12), (96950, 0.22), (206700, 0.24),
        (394600, 0.32), (501050, 0.35), (751600, 0.37),
    ],
    "Head of Household": [
        (0, 0.10), (17000, 0.12), (64850, 0.22), (103350, 0.24),
        (197300, 0.32), (250500, 0.35), (626350, 0.37),
    ],
}

STANDARD_DEDUCTION = {
    "Single": 15000.00,
    "Married Filing Jointly": 30000.00,
    "Head of Household": 22500.00,
}

st.set_page_config(page_title="Western Excavation Payroll", page_icon="🚛", layout="wide")


def db_connect():
    return sqlite3.connect(DB_PATH)


def init_db():
    with db_connect() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS employees (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                hourly_rate REAL NOT NULL DEFAULT 0,
                tax_status TEXT NOT NULL DEFAULT '1099',
                transfer_fee_enabled INTEGER NOT NULL DEFAULT 0,
                ssn_last4 TEXT,
                active INTEGER NOT NULL DEFAULT 1,
                created_at TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS payroll_runs (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                run_date TEXT NOT NULL,
                pay_period_start TEXT NOT NULL,
                pay_period_end TEXT NOT NULL,
                total_gross REAL NOT NULL,
                total_deductions REAL NOT NULL,
                total_net REAL NOT NULL,
                created_at TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS payroll_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                run_id INTEGER NOT NULL,
                employee_id INTEGER,
                employee_name TEXT NOT NULL,
                tax_status TEXT NOT NULL,
                hours REAL NOT NULL DEFAULT 0,
                hourly_rate REAL NOT NULL DEFAULT 0,
                gross_pay REAL NOT NULL,
                federal_tax REAL NOT NULL DEFAULT 0,
                virginia_tax REAL NOT NULL DEFAULT 0,
                social_security REAL NOT NULL DEFAULT 0,
                medicare REAL NOT NULL DEFAULT 0,
                transfer_fee REAL NOT NULL DEFAULT 0,
                total_deductions REAL NOT NULL,
                net_pay REAL NOT NULL,
                ssn_last4 TEXT,
                check_number TEXT,
                FOREIGN KEY(run_id) REFERENCES payroll_runs(id)
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS ytd_adjustments (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                employee_id INTEGER,
                employee_name TEXT NOT NULL,
                ssn_last4 TEXT,
                ytd_gross REAL NOT NULL DEFAULT 0,
                ytd_deductions REAL NOT NULL DEFAULT 0,
                ytd_net REAL NOT NULL DEFAULT 0,
                ytd_federal REAL NOT NULL DEFAULT 0,
                ytd_ss REAL NOT NULL DEFAULT 0,
                ytd_medicare REAL NOT NULL DEFAULT 0,
                ytd_virginia REAL NOT NULL DEFAULT 0,
                imported_at TEXT NOT NULL
            )
        """)
        # Safe migrations for existing databases from older app versions.
        existing_cols = [row[1] for row in conn.execute("PRAGMA table_info(payroll_items)").fetchall()]
        if "check_number" not in existing_cols:
            conn.execute("ALTER TABLE payroll_items ADD COLUMN check_number TEXT")
        conn.commit()


def read_df(query, params=()):
    with db_connect() as conn:
        return pd.read_sql_query(query, conn, params=params)


def add_employee(name, hourly_rate, tax_status, transfer_fee_enabled, ssn_last4):
    with db_connect() as conn:
        conn.execute(
            "INSERT INTO employees (name, hourly_rate, tax_status, transfer_fee_enabled, ssn_last4, created_at) VALUES (?, ?, ?, ?, ?, ?)",
            (name, hourly_rate, tax_status, int(transfer_fee_enabled), ssn_last4, datetime.now().isoformat(timespec="seconds")),
        )
        conn.commit()


def update_employee(emp_id, name, hourly_rate, tax_status, transfer_fee_enabled, ssn_last4, active):
    with db_connect() as conn:
        conn.execute(
            "UPDATE employees SET name=?, hourly_rate=?, tax_status=?, transfer_fee_enabled=?, ssn_last4=?, active=? WHERE id=?",
            (name, hourly_rate, tax_status, int(transfer_fee_enabled), ssn_last4, int(active), emp_id),
        )
        conn.commit()

def canonical_col_name(value):
    """Normalize Excel column names so many payroll/YTD spreadsheet styles import correctly."""
    text = str(value).strip().lower()
    text = text.replace("&", "and")
    text = re.sub(r"[^a-z0-9]+", "_", text)
    text = re.sub(r"_+", "_", text).strip("_")
    return text


def normalize_excel_columns(df):
    renamed = {}
    seen = {}
    for col in df.columns:
        key = canonical_col_name(col)
        if not key:
            key = "blank"
        if key in seen:
            seen[key] += 1
            key = f"{key}_{seen[key]}"
        else:
            seen[key] = 0
        renamed[col] = key
    return df.rename(columns=renamed)


def first_available(row, names, default=""):
    """Return the first usable column value. Supports many aliases and ignores blank/NaN values."""
    row_keys = {canonical_col_name(k): k for k in row.index}
    for name in names:
        clean = canonical_col_name(name)
        candidates = [clean]
        # Also accept duplicate columns like federal_tax_ytd_1 created by pandas.
        candidates += [k for k in row_keys if k == clean or k.startswith(clean + "_")]
        for cand in candidates:
            if cand in row_keys:
                val = row[row_keys[cand]]
                try:
                    if pd.isna(val):
                        continue
                except Exception:
                    pass
                if str(val).strip() != "":
                    return val
    return default



def normalize_name_key(value):
    """Normalize employee names so Excel imports still match paystubs with small spacing/case differences."""
    text = str(value or "").strip().lower()
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


def sync_ytd_backup_from_db():
    """Write a JSON backup of saved YTD totals beside the local database.

    The SQLite database is the main storage. The JSON backup gives the app a second
    local copy so YTD totals remain recoverable after app folder updates and are easy
    to inspect if needed.
    """
    try:
        ytd = read_df("""
            SELECT employee_id, employee_name, ssn_last4, ytd_gross, ytd_deductions, ytd_net,
                   ytd_federal, ytd_ss, ytd_medicare, ytd_virginia, imported_at
            FROM ytd_adjustments
            ORDER BY imported_at
        """)
        records = ytd.to_dict(orient="records") if not ytd.empty else []
        with open(YTD_BACKUP_PATH, "w", encoding="utf-8") as f:
            json.dump(records, f, indent=2, default=str)
    except Exception:
        pass


def load_ytd_backup_rows():
    """Load YTD rows from JSON backup if available."""
    try:
        if YTD_BACKUP_PATH.exists():
            with open(YTD_BACKUP_PATH, "r", encoding="utf-8") as f:
                data = json.load(f)
            if isinstance(data, list):
                return pd.DataFrame(data)
    except Exception:
        pass
    return pd.DataFrame(columns=[
        "employee_id", "employee_name", "ssn_last4", "ytd_gross", "ytd_deductions", "ytd_net",
        "ytd_federal", "ytd_ss", "ytd_medicare", "ytd_virginia", "imported_at"
    ])


def get_all_saved_ytd_rows_raw():
    """Return saved YTD rows from SQLite, falling back to JSON backup if needed."""
    try:
        ytd = read_df("""
            SELECT employee_id, employee_name, ssn_last4, ytd_gross, ytd_deductions, ytd_net,
                   ytd_federal, ytd_ss, ytd_medicare, ytd_virginia, imported_at
            FROM ytd_adjustments
            ORDER BY imported_at
        """)
    except Exception:
        ytd = pd.DataFrame()
    if ytd.empty:
        ytd = load_ytd_backup_rows()
        # Restore backup into SQLite so future paystubs can use it normally.
        if not ytd.empty:
            for _, r in ytd.iterrows():
                save_ytd_record(
                    r.get("employee_id", None), r.get("employee_name", ""), r.get("ssn_last4", ""),
                    r.get("ytd_gross", 0), r.get("ytd_deductions", 0), r.get("ytd_net", 0),
                    r.get("ytd_federal", 0), r.get("ytd_ss", 0), r.get("ytd_medicare", 0), r.get("ytd_virginia", 0),
                    _sync_backup=False,
                )
    return ytd

def clear_ytd_adjustments():
    with db_connect() as conn:
        conn.execute("DELETE FROM ytd_adjustments")
        conn.commit()
    try:
        if YTD_BACKUP_PATH.exists():
            YTD_BACKUP_PATH.unlink()
    except Exception:
        pass


def save_ytd_record(employee_id, employee_name, ssn_last4, ytd_gross, ytd_deductions, ytd_net, ytd_federal, ytd_ss, ytd_medicare, ytd_virginia, _sync_backup=True):
    """Save one permanent imported/base YTD row. Replaces that employee's previous saved YTD."""
    with db_connect() as conn:
        conditions = []
        params = []
        if employee_id is not None:
            conditions.append("employee_id=?")
            params.append(int(employee_id))
        if employee_name:
            conditions.append("LOWER(employee_name)=LOWER(?)")
            params.append(str(employee_name).strip())
        if ssn_last4:
            conditions.append("ssn_last4=?")
            params.append(normalize_ssn_last4(ssn_last4))
        if conditions:
            conn.execute(f"DELETE FROM ytd_adjustments WHERE {' OR '.join(conditions)}", tuple(params))
        conn.execute("""
            INSERT INTO ytd_adjustments (
                employee_id, employee_name, ssn_last4, ytd_gross, ytd_deductions, ytd_net,
                ytd_federal, ytd_ss, ytd_medicare, ytd_virginia, imported_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (
            employee_id, str(employee_name).strip(), normalize_ssn_last4(ssn_last4),
            safe_num(ytd_gross), safe_num(ytd_deductions), safe_num(ytd_net),
            safe_num(ytd_federal), safe_num(ytd_ss), safe_num(ytd_medicare), safe_num(ytd_virginia),
            datetime.now().isoformat(timespec="seconds"),
        ))
        conn.commit()
    if _sync_backup:
        sync_ytd_backup_from_db()


def import_ytd_adjustments(uploaded_file, replace_existing=True):
    """Import and permanently save base YTD totals.

    Matches employees by employee_id, employee_name/name, or SSN last 4. Supports the
    app template plus common payroll export column names. Imported totals are stored
    in ~/Western_Payroll_Data/western_payroll.db on Mac, so they only need to be
    entered/imported one time.
    """
    df = pd.read_excel(uploaded_file)
    if df.empty:
        return 0, ["The uploaded Excel file was empty."]
    df = normalize_excel_columns(df)

    employees = read_df("SELECT id, name, ssn_last4 FROM employees")
    employee_by_id = {int(r["id"]): r for _, r in employees.iterrows()}
    employee_by_name = {str(r["name"]).strip().lower(): r for _, r in employees.iterrows()}
    employee_by_ssn = {normalize_ssn_last4(r.get("ssn_last4")): r for _, r in employees.iterrows() if normalize_ssn_last4(r.get("ssn_last4"))}

    if replace_existing:
        clear_ytd_adjustments()

    inserted = 0
    skipped = []
    for idx, row in df.iterrows():
        emp_id_val = first_available(row, ["employee_id", "id", "emp_id", "employee_number", "employee_no"], "")
        employee_id = None
        try:
            if str(emp_id_val).strip() != "":
                employee_id = int(float(str(emp_id_val).strip()))
        except Exception:
            employee_id = None

        name = str(first_available(row, [
            "employee_name", "name", "employee", "employee_full_name", "full_name",
            "worker", "driver", "employee name"
        ], "")).strip()
        ssn_last4 = normalize_ssn_last4(first_available(row, [
            "ssn_last4", "last_4", "last4", "last_four", "social_sec_id", "social_sec_id",
            "social_security_id", "social_security_last4", "social_security", "ssn",
            "employee_ssn", "tax_id_last4", "social sec id"
        ], ""))

        matched = None
        if employee_id is not None and employee_id in employee_by_id:
            matched = employee_by_id[employee_id]
        elif ssn_last4 and ssn_last4 in employee_by_ssn:
            matched = employee_by_ssn[ssn_last4]
        elif name and name.lower() in employee_by_name:
            matched = employee_by_name[name.lower()]

        if matched is not None:
            employee_id = int(matched["id"])
            employee_name = str(matched["name"]).strip()
            ssn_last4 = normalize_ssn_last4(matched.get("ssn_last4") or ssn_last4)
        elif name:
            employee_name = name
        else:
            skipped.append(f"Row {idx + 2}: missing employee name, employee ID, or SSN last 4")
            continue

        ytd_gross = safe_num(first_available(row, [
            "ytd_gross", "gross_ytd", "gross_pay_ytd", "gross", "gross_pay",
            "year_to_date_gross", "year_to_date_gross_pay", "ytd_pay", "pay_ytd",
            "total_gross", "ytd_total_gross", "ytd_earnings", "earnings_ytd",
            "ytd_regular_plus_overtime", "gross_earnings_ytd"
        ], 0))
        ytd_federal = safe_num(first_available(row, [
            "ytd_federal", "federal_ytd", "ytd_federal_tax", "federal_tax_ytd",
            "federal_tax", "fed_tax", "fed_ytd", "federal_withholding",
            "federal_withholding_ytd", "federal_income_tax_ytd", "fit_ytd"
        ], 0))
        ytd_ss = safe_num(first_available(row, [
            "ytd_ss", "ss_ytd", "ytd_social_security", "social_security_ytd",
            "social_security", "social_security_tax", "social_security_tax_ytd",
            "fica_ss", "fica_social_security", "fica_social_security_ytd"
        ], 0))
        ytd_medicare = safe_num(first_available(row, [
            "ytd_medicare", "medicare_ytd", "medicare", "medicare_tax",
            "medicare_tax_ytd", "fica_medicare", "fica_medicare_ytd"
        ], 0))
        ytd_virginia = safe_num(first_available(row, [
            "ytd_virginia", "virginia_ytd", "state_withholding_ytd",
            "state_withholding", "state_tax", "state_tax_ytd", "virginia_tax_ytd",
            "va_tax", "va_tax_ytd", "virginia_withholding", "virginia_withholding_ytd",
            "sit_ytd", "state"
        ], 0))
        ytd_deductions = safe_num(first_available(row, [
            "ytd_deductions", "deductions_ytd", "total_deductions_ytd", "deductions",
            "ytd_total_deductions", "total_deductions", "taxes_ytd", "withholding_ytd"
        ], 0))
        if ytd_deductions == 0:
            ytd_deductions = ytd_federal + ytd_ss + ytd_medicare + ytd_virginia
        ytd_net = safe_num(first_available(row, [
            "ytd_net", "net_ytd", "net_pay_ytd", "net_pay", "year_to_date_net",
            "year_to_date_net_pay", "ytd_net_pay", "take_home_ytd", "total_net", "net_earnings_ytd"
        ], 0))
        if ytd_net == 0 and ytd_gross:
            ytd_net = ytd_gross - ytd_deductions

        if ytd_gross == 0 and ytd_deductions == 0 and ytd_net == 0 and ytd_federal == 0 and ytd_ss == 0 and ytd_medicare == 0 and ytd_virginia == 0:
            skipped.append(f"Row {idx + 2}: no YTD dollar amounts found")
            continue

        save_ytd_record(employee_id, employee_name, ssn_last4, ytd_gross, ytd_deductions, ytd_net, ytd_federal, ytd_ss, ytd_medicare, ytd_virginia)
        inserted += 1

    return inserted, skipped


def get_saved_ytd_rows():
    employees = read_df("SELECT id, name, ssn_last4 FROM employees ORDER BY name")
    ytd = get_all_saved_ytd_rows_raw()
    if employees.empty:
        return ytd
    rows = []
    for _, emp in employees.iterrows():
        match = pd.DataFrame()
        if not ytd.empty:
            emp_ssn = normalize_ssn_last4(emp.get("ssn_last4"))
            emp_name_key = normalize_name_key(emp.get("name"))
            emp_id = int(emp["id"])
            ytd_ids = pd.to_numeric(ytd.get("employee_id", pd.Series(dtype=object)), errors="coerce")
            mask = (ytd_ids.fillna(-999999).astype(int) == emp_id)
            mask = mask | (ytd["employee_name"].fillna("").map(normalize_name_key) == emp_name_key)
            if emp_ssn:
                mask = mask | (ytd["ssn_last4"].fillna("").astype(str).map(normalize_ssn_last4) == emp_ssn)
            match = ytd[mask]
        if not match.empty:
            r = match.iloc[-1].to_dict()
            # Always attach the current employee ID/name so future paystub exports match directly.
            r["employee_id"] = int(emp["id"])
            r["employee_name"] = str(emp["name"])
            if not normalize_ssn_last4(r.get("ssn_last4")) and normalize_ssn_last4(emp.get("ssn_last4")):
                r["ssn_last4"] = normalize_ssn_last4(emp.get("ssn_last4"))
        else:
            r = {
                "employee_id": int(emp["id"]), "employee_name": emp["name"], "ssn_last4": emp.get("ssn_last4", ""),
                "ytd_gross": 0.0, "ytd_deductions": 0.0, "ytd_net": 0.0, "ytd_federal": 0.0,
                "ytd_ss": 0.0, "ytd_medicare": 0.0, "ytd_virginia": 0.0, "imported_at": ""
            }
        rows.append(r)
    return pd.DataFrame(rows)

def save_ytd_editor_df(df):
    count = 0
    for _, row in df.iterrows():
        employee_name = str(row.get("employee_name", "")).strip()
        if not employee_name:
            continue
        employee_id = row.get("employee_id", None)
        try:
            employee_id = int(float(employee_id)) if str(employee_id).strip() not in ("", "nan", "None") else None
        except Exception:
            employee_id = None
        ytd_gross = safe_num(row.get("ytd_gross", 0))
        ytd_federal = safe_num(row.get("ytd_federal", 0))
        ytd_ss = safe_num(row.get("ytd_ss", 0))
        ytd_medicare = safe_num(row.get("ytd_medicare", 0))
        ytd_virginia = safe_num(row.get("ytd_virginia", 0))
        ytd_deductions = safe_num(row.get("ytd_deductions", 0))
        if ytd_deductions == 0:
            ytd_deductions = ytd_federal + ytd_ss + ytd_medicare + ytd_virginia
        ytd_net = safe_num(row.get("ytd_net", 0))
        if ytd_net == 0 and ytd_gross:
            ytd_net = ytd_gross - ytd_deductions
        save_ytd_record(employee_id, employee_name, row.get("ssn_last4", ""), ytd_gross, ytd_deductions, ytd_net, ytd_federal, ytd_ss, ytd_medicare, ytd_virginia)
        count += 1
    return count


def build_ytd_template():
    employees = read_df("SELECT name AS employee_name, ssn_last4 FROM employees ORDER BY name")
    if employees.empty:
        employees = pd.DataFrame(columns=["employee_name", "ssn_last4"])
    for col in ["ytd_gross", "ytd_federal", "ytd_social_security", "ytd_medicare", "ytd_virginia", "ytd_deductions", "ytd_net"]:
        employees[col] = 0.00
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        employees.to_excel(writer, sheet_name="YTD Totals", index=False)
    output.seek(0)
    return output.getvalue()


def annual_tax_from_brackets(taxable, brackets):
    taxable = max(0.0, float(taxable))
    tax = 0.0
    for i, (start, rate) in enumerate(brackets):
        end = brackets[i + 1][0] if i + 1 < len(brackets) else None
        if end is None:
            if taxable > start:
                tax += (taxable - start) * rate
        else:
            if taxable > start:
                tax += (min(taxable, end) - start) * rate
    return max(0.0, tax)


def virginia_annual_tax(taxable):
    taxable = max(0.0, float(taxable))
    tax = 0.0
    tax += min(taxable, 3000) * 0.02
    if taxable > 3000:
        tax += min(taxable - 3000, 2000) * 0.03
    if taxable > 5000:
        tax += min(taxable - 5000, 12000) * 0.05
    if taxable > 17000:
        tax += (taxable - 17000) * 0.0575
    return max(0.0, tax)


def calc_payroll(
    gross_pay,
    tax_status,
    transfer_fee_enabled,
    pay_frequency="Weekly",
    filing_status="Single",
    federal_dependents=0.0,
    federal_other_income=0.0,
    federal_extra_deductions=0.0,
    federal_extra_withholding=0.0,
    va_exemptions=0,
    va_age_blind_exemptions=0,
    va_extra_withholding=0.0,
):
    gross_pay = float(gross_pay)
    if tax_status == "W-2":
        periods = PAY_PERIODS.get(pay_frequency, 52)
        annual_gross = gross_pay * periods

        standard_deduction = STANDARD_DEDUCTION.get(filing_status, STANDARD_DEDUCTION["Single"])
        annual_federal_taxable = annual_gross + float(federal_other_income) - standard_deduction - float(federal_extra_deductions)
        annual_federal_tax = annual_tax_from_brackets(annual_federal_taxable, FEDERAL_BRACKETS.get(filing_status, FEDERAL_BRACKETS["Single"]))
        federal = max(0.0, (annual_federal_tax - float(federal_dependents)) / periods + float(federal_extra_withholding))

        # Virginia method: annualize the check, subtract the Virginia standard deduction and VA-4 exemptions,
        # apply Virginia's progressive withholding brackets, then divide back to the pay period.
        # This follows the method PaycheckCity exposes for Virginia inputs and the Virginia withholding formula.
        va_standard_deduction = 17500.00 if filing_status == "Married Filing Jointly" else 8750.00
        va_personal_exemption_value = 930.00
        va_age_blind_exemption_value = 800.00
        annual_va_taxable = annual_gross - va_standard_deduction - (max(0, int(va_exemptions)) * va_personal_exemption_value) - (max(0, int(va_age_blind_exemptions)) * va_age_blind_exemption_value)
        virginia = max(0.0, virginia_annual_tax(annual_va_taxable) / periods + float(va_extra_withholding))

        ss_wages = min(gross_pay, SOCIAL_SECURITY_WAGE_BASE / periods)
        ss = ss_wages * SOCIAL_SECURITY_RATE
        medicare = gross_pay * MEDICARE_RATE
        addl_threshold = 200000 / periods
        if gross_pay > addl_threshold:
            medicare += (gross_pay - addl_threshold) * ADDITIONAL_MEDICARE_RATE
    else:
        federal = virginia = ss = medicare = 0.0
    transfer = DEFAULT_TRANSFER_FEE if transfer_fee_enabled else 0.0
    deductions = federal + virginia + ss + medicare + transfer
    net = gross_pay - deductions
    return {
        "federal_tax": round(federal, 2),
        "virginia_tax": round(virginia, 2),
        "social_security": round(ss, 2),
        "medicare": round(medicare, 2),
        "transfer_fee": round(transfer, 2),
        "total_deductions": round(deductions, 2),
        "net_pay": round(net, 2),
    }


def save_payroll_run(pay_period_start, pay_period_end, pay_date, rows):
    total_gross = sum(r["gross_pay"] for r in rows)
    total_deductions = sum(r["total_deductions"] for r in rows)
    total_net = sum(r["net_pay"] for r in rows)
    with db_connect() as conn:
        cur = conn.cursor()
        cur.execute(
            "INSERT INTO payroll_runs (run_date, pay_period_start, pay_period_end, total_gross, total_deductions, total_net, created_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (str(pay_date), str(pay_period_start), str(pay_period_end), total_gross, total_deductions, total_net, datetime.now().isoformat(timespec="seconds")),
        )
        run_id = cur.lastrowid
        for r in rows:
            cur.execute("""
                INSERT INTO payroll_items (
                    run_id, employee_id, employee_name, tax_status, hours, hourly_rate, gross_pay,
                    federal_tax, virginia_tax, social_security, medicare, transfer_fee,
                    total_deductions, net_pay, ssn_last4, check_number
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, (
                run_id, r["employee_id"], r["employee_name"], r["tax_status"], r["hours"], r["hourly_rate"], r["gross_pay"],
                r["federal_tax"], r["virginia_tax"], r["social_security"], r["medicare"], r["transfer_fee"],
                r["total_deductions"], r["net_pay"], r.get("ssn_last4", ""), r.get("check_number", ""),
            ))
        conn.commit()
    return run_id


def money(x):
    return f"${float(x):,.2f}"


def safe_num(x):
    """Parse normal Excel/payroll money values safely. Handles $1,234.56, (123.45), blanks, and numeric cells."""
    try:
        if pd.isna(x):
            return 0.0
    except Exception:
        pass
    try:
        return float(x)
    except Exception:
        try:
            text = str(x).strip()
            if not text:
                return 0.0
            neg = text.startswith("(") and text.endswith(")")
            text = text.replace("$", "").replace(",", "").replace(" ", "").replace("(", "").replace(")", "")
            val = float(text)
            return -val if neg else val
        except Exception:
            return 0.0


def normalize_ssn_last4(value):
    """Normalize SSN last 4 from Excel text or numbers like 7682, 7682.0, xxx-xx-7682."""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    text = str(value).strip()
    if not text:
        return ""
    # Excel often stores 1234 as 1234.0
    if text.endswith(".0"):
        text = text[:-2]
    digits = "".join(ch for ch in text if ch.isdigit())
    if len(digits) >= 4:
        return digits[-4:]
    return digits.zfill(4) if digits else ""



def split_regular_overtime(hours, hourly_rate):
    """Return regular/overtime hour and pay amounts. Overtime is anything over 40 hours at time-and-a-half."""
    total_hours = safe_num(hours)
    rate = safe_num(hourly_rate)
    regular_hours = min(total_hours, 40.0)
    overtime_hours = max(total_hours - 40.0, 0.0)
    overtime_rate = rate * 1.5
    regular_pay = regular_hours * rate
    overtime_pay = overtime_hours * overtime_rate
    gross_pay = regular_pay + overtime_pay
    return {
        "regular_hours": round(regular_hours, 2),
        "overtime_hours": round(overtime_hours, 2),
        "regular_rate": round(rate, 2),
        "overtime_rate": round(overtime_rate, 2),
        "regular_pay": round(regular_pay, 2),
        "overtime_pay": round(overtime_pay, 2),
        "gross_pay": round(gross_pay, 2),
    }

def get_ytd_for_item(item, run_id):
    """Return paystub YTD totals as imported/saved beginning YTD + payroll runs through this check.

    This is the main fix: paystub PDFs now explicitly pull the permanently saved
    imported YTD totals and add the current app payroll history. Matching is done
    by employee ID first, then SSN last 4, then normalized employee name.
    """
    employee_id_raw = item.get("employee_id")
    employee_id = None
    try:
        if employee_id_raw is not None and not pd.isna(employee_id_raw) and str(employee_id_raw).strip() != "":
            employee_id = int(float(employee_id_raw))
    except Exception:
        employee_id = None

    employee_name = str(item.get("employee_name", "") or "").strip()
    employee_name_key = normalize_name_key(employee_name)
    ssn_last4 = normalize_ssn_last4(item.get("ssn_last4", ""))

    # Payroll totals created inside this app through the selected run, including this check.
    if employee_id is not None:
        params = (employee_id, int(run_id))
        where = "employee_id=? AND run_id<=?"
    else:
        params = (employee_name, int(run_id))
        where = "LOWER(employee_name)=LOWER(?) AND run_id<=?"

    df = read_df(f"""
        SELECT
            COALESCE(SUM(gross_pay),0) AS ytd_gross,
            COALESCE(SUM(total_deductions),0) AS ytd_deductions,
            COALESCE(SUM(net_pay),0) AS ytd_net,
            COALESCE(SUM(federal_tax),0) AS ytd_federal,
            COALESCE(SUM(social_security),0) AS ytd_ss,
            COALESCE(SUM(medicare),0) AS ytd_medicare,
            COALESCE(SUM(virginia_tax),0) AS ytd_virginia
        FROM payroll_items
        WHERE {where}
    """, params)
    payroll_totals = df.iloc[0].to_dict() if not df.empty else {}

    totals = {k: safe_num(payroll_totals.get(k, 0)) for k in [
        "ytd_gross", "ytd_deductions", "ytd_net", "ytd_federal", "ytd_ss", "ytd_medicare", "ytd_virginia"
    ]}

    # Permanently saved/imported beginning YTD totals.
    ytd_rows = get_all_saved_ytd_rows_raw()
    base = {k: 0.0 for k in totals}
    if not ytd_rows.empty:
        # Normalize types/fields.
        ytd_rows = ytd_rows.copy()
        if "employee_id" in ytd_rows.columns:
            ytd_rows["_employee_id_num"] = pd.to_numeric(ytd_rows["employee_id"], errors="coerce")
        else:
            ytd_rows["_employee_id_num"] = pd.NA
        ytd_rows["_ssn"] = ytd_rows.get("ssn_last4", pd.Series([""] * len(ytd_rows))).astype(str).map(normalize_ssn_last4)
        ytd_rows["_name_key"] = ytd_rows.get("employee_name", pd.Series([""] * len(ytd_rows))).astype(str).map(normalize_name_key)

        matched = pd.DataFrame()
        if employee_id is not None:
            matched = ytd_rows[ytd_rows["_employee_id_num"].fillna(-999999).astype(int) == int(employee_id)]
        if matched.empty and ssn_last4:
            matched = ytd_rows[ytd_rows["_ssn"] == ssn_last4]
        if matched.empty and employee_name_key:
            matched = ytd_rows[ytd_rows["_name_key"] == employee_name_key]

        if not matched.empty:
            # Use the latest saved/imported row for this employee. This avoids accidentally doubling
            # the same employee's YTD if they imported the file more than once.
            r = matched.iloc[-1]
            base = {
                "ytd_gross": safe_num(r.get("ytd_gross", 0)),
                "ytd_deductions": safe_num(r.get("ytd_deductions", 0)),
                "ytd_net": safe_num(r.get("ytd_net", 0)),
                "ytd_federal": safe_num(r.get("ytd_federal", 0)),
                "ytd_ss": safe_num(r.get("ytd_ss", 0)),
                "ytd_medicare": safe_num(r.get("ytd_medicare", 0)),
                "ytd_virginia": safe_num(r.get("ytd_virginia", 0)),
            }
            if base["ytd_deductions"] == 0:
                base["ytd_deductions"] = base["ytd_federal"] + base["ytd_ss"] + base["ytd_medicare"] + base["ytd_virginia"]
            if base["ytd_net"] == 0 and base["ytd_gross"]:
                base["ytd_net"] = base["ytd_gross"] - base["ytd_deductions"]

    for key in totals:
        totals[key] = round(safe_num(base.get(key, 0)) + safe_num(totals.get(key, 0)), 2)

    return totals

def create_paystub_pdf(item, run):
    """Create a Western Excavation earning-statement paystub PDF.

    The layout is intentionally matched to the uploaded sample earning statement:
    centered Western Excavation logo at the top, EARNING STATEMENT title,
    blue section bars, employee/check info row, earnings/deductions row,
    totals row, and centered business address/footer.
    """
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer,
        pagesize=landscape(letter),
        rightMargin=18,
        leftMargin=18,
        topMargin=14,
        bottomMargin=18,
    )
    styles = getSampleStyleSheet()
    normal = ParagraphStyle("normal", parent=styles["Normal"], fontName="Times-Roman", fontSize=10, leading=12)
    small = ParagraphStyle("small", parent=styles["Normal"], fontName="Times-Roman", fontSize=9, leading=11)
    title = ParagraphStyle("paystub_title", parent=styles["Title"], fontName="Times-Bold", fontSize=12, leading=14, alignment=1)
    footer = ParagraphStyle("footer", parent=styles["Normal"], fontName="Times-Bold", fontSize=10, leading=12, alignment=1)

    ytd = get_ytd_for_item(item, int(run["id"])) if "id" in run else {}

    def _fmt_stub_date(value):
        """Format paystub dates compactly so header columns stay straight."""
        if not value:
            return ""
        if isinstance(value, (date, datetime)):
            return f"{value.month}/{value.day}/{value.year}"
        text = str(value).strip()
        for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m/%d/%y"):
            try:
                dt = datetime.strptime(text, fmt)
                return f"{dt.month}/{dt.day}/{dt.year}"
            except Exception:
                pass
        return text

    story = []

    # Logo centered at top, like the sample paystub.
    if LOGO_PATH.exists():
        logo = Image(str(LOGO_PATH))
        logo.drawWidth = 5.65 * inch
        logo.drawHeight = logo.drawWidth * (356 / 975)
        logo.hAlign = "CENTER"
        story.append(logo)
    else:
        story.append(Paragraph(f"<b>{BUSINESS_NAME}</b>", title))
    story.append(Spacer(1, 2))
    story.append(Paragraph("<b>EARNING STATEMENT</b>", title))
    story.append(Spacer(1, 24))

    check_no = str(item.get("check_number") or "").strip()
    if not check_no:
        check_no = f"{int(run['id']):04d}{int(item.get('id', 0) or 0):02d}" if "id" in run else "----"
    ssn = f"xxx-xx-{item.get('ssn_last4') or '----'}"
    pay_record = f"{_fmt_stub_date(run['pay_period_start'])} - {_fmt_stub_date(run['pay_period_end'])}"
    pay_date = _fmt_stub_date(run.get("run_date", ""))
    blue = colors.HexColor("#9dccf3")

    # Employee/check section with blue header bar.
    header_data = [
        ["Employee Name", "Social Sec. ID", "Check No.", "Pay Record", "Pay Date"],
        [item["employee_name"], ssn, check_no, pay_record, pay_date],
    ]
    # Fixed widths keep Social Sec. ID, Check No., Pay Record, and Pay Date aligned.
    header = Table(
        header_data,
        colWidths=[2.45*inch, 1.55*inch, 1.15*inch, 3.10*inch, 1.25*inch],
        rowHeights=[0.28*inch, 0.44*inch],
        hAlign="CENTER",
    )
    header.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), blue),
        ("FONTNAME", (0, 0), (-1, 0), "Times-Bold"),
        ("FONTNAME", (0, 1), (-1, 1), "Times-Roman"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("ALIGN", (0, 1), (-1, 1), "CENTER"),
        ("ALIGN", (0, 1), (0, 1), "LEFT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 4),
        ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
        ("LINEBELOW", (0, 0), (-1, 0), 0.25, colors.white),
    ]))
    story.append(header)
    story.append(Spacer(1, 18))

    # Earnings/deductions section with blue header bar.
    ot = split_regular_overtime(item.get("hours", 0), item.get("hourly_rate", 0))
    earnings_deductions = [
        ["Earnings", "Rate", "Hours", "Current", "Deductions", "Current", "Year to Date"],
        ["Regular Hours", f"{ot['regular_rate']:.2f}", f"{ot['regular_hours']:.2f}", f"{ot['regular_pay']:.2f}", "Federal Tax", f"{safe_num(item['federal_tax']):.2f}", f"{safe_num(ytd.get('ytd_federal', item['federal_tax'])):.2f}"],
        ["Overtime Hours", f"{ot['overtime_rate']:.2f}", f"{ot['overtime_hours']:.2f}", f"{ot['overtime_pay']:.2f}", "Social Security", f"{safe_num(item['social_security']):.2f}", f"{safe_num(ytd.get('ytd_ss', item['social_security'])):.2f}"],
        ["", "", "", "", "Medicare", f"{safe_num(item['medicare']):.2f}", f"{safe_num(ytd.get('ytd_medicare', item['medicare'])):.2f}"],
        ["", "", "", "", "State Withholding", f"{safe_num(item['virginia_tax']):.2f}", f"{safe_num(ytd.get('ytd_virginia', item['virginia_tax'])):.2f}"],
    ]
    if safe_num(item.get("transfer_fee", 0)):
        earnings_deductions.append(["", "", "", "", "Transfer Fee", f"{safe_num(item['transfer_fee']):.2f}", f"{safe_num(item['transfer_fee']):.2f}"])

    # Fixed column widths and matching alignment keep Rate and Hours straight on exported paystubs.
    ed = Table(
        earnings_deductions,
        colWidths=[1.75*inch, .95*inch, .95*inch, 1.05*inch, 1.65*inch, 1.05*inch, 1.2*inch],
        rowHeights=[0.28*inch] + [0.27*inch]*(len(earnings_deductions)-1),
        hAlign="CENTER",
    )
    ed.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), blue),
        ("FONTNAME", (0, 0), (-1, 0), "Times-Bold"),
        ("FONTNAME", (0, 1), (-1, -1), "Times-Roman"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("ALIGN", (0, 1), (0, -1), "LEFT"),
        ("ALIGN", (1, 0), (3, -1), "CENTER"),
        ("ALIGN", (4, 1), (4, -1), "LEFT"),
        ("ALIGN", (5, 0), (6, -1), "RIGHT"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
    ]))
    story.append(ed)
    story.append(Spacer(1, 18))

    # Totals section with blue header bar.
    totals = [
        ["YTD Gross", "YTD Deductions", "YTD Net Pay", "Current Total", "Current Deductions", "Net Pay"],
        [
            f"{safe_num(ytd.get('ytd_gross', item['gross_pay'])):,.2f}",
            f"{safe_num(ytd.get('ytd_deductions', item['total_deductions'])):,.2f}",
            f"{safe_num(ytd.get('ytd_net', item['net_pay'])):,.2f}",
            f"{safe_num(item['gross_pay']):,.2f}",
            f"{safe_num(item['total_deductions']):,.2f}",
            f"{safe_num(item['net_pay']):,.2f}",
        ],
    ]
    tt = Table(totals, colWidths=[1.5*inch, 1.52*inch, 1.52*inch, 1.5*inch, 1.65*inch, 1.25*inch], rowHeights=[0.28*inch, 0.43*inch])
    tt.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), blue),
        ("FONTNAME", (0, 0), (-1, 0), "Times-Bold"),
        ("FONTNAME", (0, 1), (-1, 1), "Times-Roman"),
        ("FONTNAME", (-1, 0), (-1, 0), "Times-Bold"),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("ALIGN", (0, 0), (-1, -1), "CENTER"),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 3),
        ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("TOPPADDING", (0, 0), (-1, -1), 1),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
    ]))
    story.append(tt)
    story.append(Spacer(1, 22))
    story.append(Paragraph(f"{BUSINESS_ADDRESS}", footer))
    story.append(Paragraph(f"{BUSINESS_PHONE}", small))

    doc.build(story)
    buffer.seek(0)
    return buffer.getvalue()


def export_run_excel(run_id):
    run = read_df("SELECT * FROM payroll_runs WHERE id=?", (run_id,))
    items = read_df("SELECT * FROM payroll_items WHERE run_id=? ORDER BY employee_name", (run_id,))
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        run.to_excel(writer, sheet_name="Run Summary", index=False)
        items.to_excel(writer, sheet_name="Payroll Items", index=False)
    output.seek(0)
    return output.getvalue()


def export_paystubs_zip(run_id):
    run_df = read_df("SELECT * FROM payroll_runs WHERE id=?", (run_id,))
    items = read_df("SELECT * FROM payroll_items WHERE run_id=? ORDER BY employee_name", (run_id,))
    if run_df.empty:
        return None
    run = run_df.iloc[0].to_dict()
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
        for _, row in items.iterrows():
            item = row.to_dict()
            safe_name = "".join(c for c in item["employee_name"] if c.isalnum() or c in (" ", "_", "-")).strip().replace(" ", "_")
            zf.writestr(f"paystub_{safe_name}_run_{run_id}.pdf", create_paystub_pdf(item, run))
    zip_buffer.seek(0)
    return zip_buffer.getvalue()


def df_money_columns(df, cols):
    out = df.copy()
    for c in cols:
        if c in out.columns:
            out[c] = out[c].apply(money)
    return out


init_db()

st.title("Western Excavation Payroll")
st.caption("Streamlit payroll app. No hauling features. Uses PaycheckCity-style Virginia paycheck inputs, YTD imports, check numbers, and sample-style earning statement paystubs.")
st.info("Payroll taxes are estimates. Virginia withholding now uses the Virginia annualized wage formula: gross pay × pay periods, less standard deduction and VA-4 exemptions, then Virginia brackets divided back to the paycheck. Compare with PaycheckCity and review with a payroll professional/accountant before using for official payroll.")

tab_employees, tab_payroll, tab_history = st.tabs(["Employees", "Run Payroll", "Payroll History"])

with tab_employees:
    st.subheader("Add New Employee")
    with st.form("add_employee_form", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        name = c1.text_input("Employee name")
        hourly_rate = c2.number_input("Hourly rate", min_value=0.0, step=1.0, format="%.2f")
        tax_status = c3.selectbox("Tax status", ["1099", "W-2"])
        c4, c5 = st.columns(2)
        transfer_fee = c4.checkbox("Deduct $35 transfer fee")
        ssn_last4 = c5.text_input("Last 4 of SSN", max_chars=4)
        submitted = st.form_submit_button("Add Employee")
        if submitted:
            if not name.strip():
                st.error("Employee name is required.")
            elif ssn_last4 and (not ssn_last4.isdigit() or len(ssn_last4) != 4):
                st.error("SSN last 4 must be exactly 4 digits.")
            else:
                add_employee(name.strip(), hourly_rate, tax_status, transfer_fee, ssn_last4.strip())
                st.success(f"Added {name.strip()}.")
                st.rerun()

    st.subheader("Import Year-to-Date Totals")
    st.caption("Upload an Excel file to add previous YTD totals to paystubs. Match employees by employee_name/name or ssn_last4. Money values can include $ signs/commas, and SSN last 4 can be typed as 1234, 1234.0, or xxx-xx-1234. These imported totals are added to payroll runs saved inside this app.")
    c_ytd1, c_ytd2 = st.columns(2)
    c_ytd1.download_button(
        "Download YTD Excel Template",
        data=build_ytd_template(),
        file_name="western_payroll_ytd_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    uploaded_ytd = c_ytd2.file_uploader("Upload YTD Excel", type=["xlsx", "xls"], key="ytd_upload")
    replace_ytd = st.checkbox("Replace existing imported YTD totals", value=True)
    if uploaded_ytd is not None and st.button("Import YTD Totals"):
        try:
            count, skipped = import_ytd_adjustments(uploaded_ytd, replace_existing=replace_ytd)
            st.success(f"Imported YTD totals for {count} employee(s).")
            if skipped:
                st.warning("Skipped rows: " + "; ".join(skipped[:10]))
            st.info("YTD totals were saved and will be used on paystub exports from now on.")
            st.rerun()
        except Exception as e:
            st.error(f"Could not import YTD Excel file: {e}")
    st.markdown("**Saved YTD totals**")
    st.caption("These totals are saved in ~/Western_Payroll_Data/western_payroll.db and backed up to ~/Western_Payroll_Data/ytd_totals_backup.json, so you only have to import or type them once. Review them here before making paystubs.")
    ytd_saved = get_saved_ytd_rows()
    if not ytd_saved.empty:
        edit_cols = ["employee_id", "employee_name", "ssn_last4", "ytd_gross", "ytd_federal", "ytd_ss", "ytd_medicare", "ytd_virginia", "ytd_deductions", "ytd_net"]
        ytd_editor = st.data_editor(
            ytd_saved[edit_cols],
            width="stretch",
            hide_index=True,
            disabled=["employee_id", "employee_name"],
            column_config={
                "employee_id": st.column_config.NumberColumn("Employee ID", disabled=True),
                "employee_name": st.column_config.TextColumn("Employee", disabled=True),
                "ssn_last4": st.column_config.TextColumn("SSN Last 4"),
                "ytd_gross": st.column_config.NumberColumn("YTD Gross", format="$%.2f"),
                "ytd_federal": st.column_config.NumberColumn("YTD Federal", format="$%.2f"),
                "ytd_ss": st.column_config.NumberColumn("YTD Social Security", format="$%.2f"),
                "ytd_medicare": st.column_config.NumberColumn("YTD Medicare", format="$%.2f"),
                "ytd_virginia": st.column_config.NumberColumn("YTD Virginia", format="$%.2f"),
                "ytd_deductions": st.column_config.NumberColumn("YTD Deductions", format="$%.2f"),
                "ytd_net": st.column_config.NumberColumn("YTD Net", format="$%.2f"),
            },
            key="saved_ytd_editor",
        )
        c_save_ytd, c_clear_ytd = st.columns(2)
        if c_save_ytd.button("Save Edited YTD Totals"):
            saved_count = save_ytd_editor_df(ytd_editor)
            st.success(f"Saved YTD totals for {saved_count} employee(s).")
            st.rerun()
        if c_clear_ytd.button("Clear Imported/Saved YTD Totals"):
            clear_ytd_adjustments()
            st.success("Saved YTD totals cleared.")
            st.rerun()

    st.subheader("Current Employees")
    employees = read_df("SELECT * FROM employees ORDER BY active DESC, name")
    if employees.empty:
        st.info("No employees saved yet.")
    else:
        display = employees[["id", "name", "hourly_rate", "tax_status", "transfer_fee_enabled", "ssn_last4", "active"]].copy()
        display["transfer_fee_enabled"] = display["transfer_fee_enabled"].map({1: "Yes", 0: "No"})
        display["active"] = display["active"].map({1: "Active", 0: "Inactive"})
        st.dataframe(display, width="stretch", hide_index=True)

        st.subheader("Edit Employee")
        selected_id = st.selectbox("Select employee to edit", employees["id"], format_func=lambda x: employees.loc[employees["id"] == x, "name"].iloc[0])
        emp = employees.loc[employees["id"] == selected_id].iloc[0]
        with st.form("edit_employee_form"):
            c1, c2, c3 = st.columns(3)
            edit_name = c1.text_input("Name", value=emp["name"])
            edit_rate = c2.number_input("Hourly rate", min_value=0.0, value=float(emp["hourly_rate"]), step=1.0, format="%.2f")
            edit_tax = c3.selectbox("Tax status", ["1099", "W-2"], index=["1099", "W-2"].index(emp["tax_status"]))
            c4, c5, c6 = st.columns(3)
            edit_fee = c4.checkbox("Deduct $35 transfer fee", value=bool(emp["transfer_fee_enabled"]))
            edit_ssn = c5.text_input("Last 4 of SSN", value=emp["ssn_last4"] or "", max_chars=4)
            edit_active = c6.checkbox("Active", value=bool(emp["active"]))
            save_edit = st.form_submit_button("Save Changes")
            if save_edit:
                if edit_ssn and (not edit_ssn.isdigit() or len(edit_ssn) != 4):
                    st.error("SSN last 4 must be exactly 4 digits.")
                else:
                    update_employee(int(selected_id), edit_name.strip(), edit_rate, edit_tax, edit_fee, edit_ssn.strip(), edit_active)
                    st.success("Employee updated.")
                    st.rerun()

with tab_payroll:
    st.subheader("Run Payroll")
    active = read_df("SELECT * FROM employees WHERE active=1 ORDER BY name")
    if active.empty:
        st.warning("Add at least one active employee first.")
    else:
        c1, c2, c3, c4 = st.columns(4)
        period_start = c1.date_input("Pay period start", value=date.today())
        period_end = c2.date_input("Pay period end", value=date.today())
        pay_date = c3.date_input("Pay date", value=date.today())
        pay_frequency = c4.selectbox("Pay frequency", list(PAY_PERIODS.keys()), index=0)

        with st.expander("W-2 tax setup — PaycheckCity-style inputs", expanded=False):
            st.caption("These inputs apply to W-2 employees. 1099 employees do not have taxes deducted.")
            f1, f2, f3 = st.columns(3)
            filing_status = f1.selectbox("Federal filing status", ["Single", "Married Filing Jointly", "Head of Household"])
            federal_dependents = f2.number_input("Federal dependent credits per year", min_value=0.0, value=0.0, step=100.0, format="%.2f")
            federal_extra_withholding = f3.number_input("Extra federal withholding per check", min_value=0.0, value=0.0, step=5.0, format="%.2f")
            f4, f5, f6 = st.columns(3)
            federal_other_income = f4.number_input("Other annual income", min_value=0.0, value=0.0, step=100.0, format="%.2f")
            federal_extra_deductions = f5.number_input("Extra annual deductions", min_value=0.0, value=0.0, step=100.0, format="%.2f")
            va_exemptions = f6.number_input("Virginia VA-4 personal/dependent exemptions", min_value=0, value=0, step=1)
            f7, f8 = st.columns(2)
            va_age_blind_exemptions = f7.number_input("Virginia age 65/blind exemptions", min_value=0, value=0, step=1)
            va_extra_withholding = f8.number_input("Extra Virginia withholding per check", min_value=0.0, value=0.0, step=5.0, format="%.2f")

        st.markdown("Enter hours for each employee. Leave hours at 0 to skip that employee.")
        payroll_rows = []
        with st.form("payroll_hours_form"):
            for _, emp in active.iterrows():
                cols = st.columns([2, 1, 1, 1, 1])
                cols[0].write(f"**{emp['name']}**")
                hours = cols[1].number_input("Hours", min_value=0.0, step=1.0, format="%.2f", key=f"hours_{emp['id']}")
                rate = cols[2].number_input("Rate", min_value=0.0, value=float(emp["hourly_rate"]), step=1.0, format="%.2f", key=f"rate_{emp['id']}")
                check_number = cols[3].text_input("Check No.", key=f"check_{emp['id']}")
                cols[4].write(emp["tax_status"])
                if hours > 0:
                    ot = split_regular_overtime(hours, rate)
                    gross = ot["gross_pay"]
                    calc = calc_payroll(
                        gross,
                        emp["tax_status"],
                        bool(emp["transfer_fee_enabled"]),
                        pay_frequency=pay_frequency,
                        filing_status=filing_status,
                        federal_dependents=federal_dependents,
                        federal_other_income=federal_other_income,
                        federal_extra_deductions=federal_extra_deductions,
                        federal_extra_withholding=federal_extra_withholding,
                        va_exemptions=va_exemptions,
                        va_age_blind_exemptions=va_age_blind_exemptions,
                        va_extra_withholding=va_extra_withholding,
                    )
                    payroll_rows.append({
                        "employee_id": int(emp["id"]),
                        "employee_name": emp["name"],
                        "tax_status": emp["tax_status"],
                        "hours": hours,
                        "regular_hours": ot["regular_hours"],
                        "overtime_hours": ot["overtime_hours"],
                        "hourly_rate": rate,
                        "overtime_rate": ot["overtime_rate"],
                        "regular_pay": ot["regular_pay"],
                        "overtime_pay": ot["overtime_pay"],
                        "gross_pay": gross,
                        "ssn_last4": emp.get("ssn_last4", ""),
                        "check_number": check_number.strip(),
                        **calc,
                    })
            preview = st.form_submit_button("Preview Payroll")

        if preview:
            if not payroll_rows:
                st.warning("Enter hours for at least one employee.")
            else:
                st.session_state["payroll_preview"] = payroll_rows
                st.session_state["period_start"] = str(period_start)
                st.session_state["period_end"] = str(period_end)
                st.session_state["pay_date"] = str(pay_date)

        if "payroll_preview" in st.session_state:
            rows = st.session_state["payroll_preview"]
            st.subheader("Payroll Preview")
            preview_df = pd.DataFrame(rows)
            show_cols = ["employee_name", "check_number", "tax_status", "hours", "regular_hours", "overtime_hours", "hourly_rate", "overtime_rate", "regular_pay", "overtime_pay", "gross_pay", "federal_tax", "virginia_tax", "social_security", "medicare", "transfer_fee", "total_deductions", "net_pay"]
            st.dataframe(df_money_columns(preview_df[show_cols], ["hourly_rate", "overtime_rate", "regular_pay", "overtime_pay", "gross_pay", "federal_tax", "virginia_tax", "social_security", "medicare", "transfer_fee", "total_deductions", "net_pay"]), width="stretch", hide_index=True)
            c1, c2, c3 = st.columns(3)
            c1.metric("Total Gross", money(sum(r["gross_pay"] for r in rows)))
            c2.metric("Total Deductions", money(sum(r["total_deductions"] for r in rows)))
            c3.metric("Total Net", money(sum(r["net_pay"] for r in rows)))
            if st.button("Save Payroll Run"):
                run_id = save_payroll_run(st.session_state["period_start"], st.session_state["period_end"], st.session_state.get("pay_date", str(date.today())), rows)
                st.session_state.pop("payroll_preview", None)
                st.success(f"Payroll run #{run_id} saved.")
                st.rerun()

with tab_history:
    st.subheader("Payroll History")
    runs = read_df("SELECT * FROM payroll_runs ORDER BY id DESC")
    if runs.empty:
        st.info("No payroll runs saved yet.")
    else:
        runs_display = df_money_columns(runs[["id", "run_date", "pay_period_start", "pay_period_end", "total_gross", "total_deductions", "total_net"]], ["total_gross", "total_deductions", "total_net"])
        st.dataframe(runs_display, width="stretch", hide_index=True)
        run_id = st.selectbox("Select payroll run", runs["id"], format_func=lambda x: f"Run #{x}")
        run = runs.loc[runs["id"] == run_id].iloc[0]
        items = read_df("SELECT * FROM payroll_items WHERE run_id=? ORDER BY employee_name", (int(run_id),))
        if not items.empty:
            st.subheader(f"Run #{run_id} Details")
            detail_cols = ["employee_name", "check_number", "tax_status", "hours", "hourly_rate", "gross_pay", "federal_tax", "virginia_tax", "social_security", "medicare", "transfer_fee", "total_deductions", "net_pay"]
            st.dataframe(df_money_columns(items[detail_cols], ["hourly_rate", "gross_pay", "federal_tax", "virginia_tax", "social_security", "medicare", "transfer_fee", "total_deductions", "net_pay"]), width="stretch", hide_index=True)
            c1, c2 = st.columns(2)
            c1.download_button(
                "Download Excel Report",
                data=export_run_excel(int(run_id)),
                file_name=f"western_payroll_run_{run_id}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            c2.download_button(
                "Download Paystubs ZIP",
                data=export_paystubs_zip(int(run_id)),
                file_name=f"western_paystubs_run_{run_id}.zip",
                mime="application/zip",
            )
