# merge_first_reconciliation_v12.py
# ================================================================
# MERGE-FIRST KYC RECONCILIATION
# ================================================================
# Required business sequence:
#   0. Clean Paymeter first.
#   1. Merge Paymeter <-> Macron.
#   2. Merge VPS <-> Bank credits.
#   3. Build KYC starting from VPS virtual account numbers, using BOTH merges.
#   4. Attach KYC to Paymeter-Macron merge.
#   5. Attach KYC to VPS-Bank merge.
#   6. Reconcile the two KYC-enriched merged reports.
#   7. Compare VPS against Macron using KYC and:
#          Expected Macron Token = VPS transaction_amount_minor - N100.
#   8. Build complete Customer Summary and hard source tie-out.
# ================================================================

import csv
import io
import re
import gc
import os
import tempfile
from datetime import datetime, date
from collections import defaultdict, deque
from pathlib import Path

import numpy as np
import pandas as pd

# ================================================================
# STREAMLIT CLOUD EXCEL DEPENDENCY CHECK
# ================================================================
# pandas requires openpyxl to read .xlsx files.
try:
    import openpyxl  # noqa: F401
except ImportError as exc:
    raise ImportError(
        "Missing dependency: openpyxl. "
        "For Streamlit Community Cloud, create a file named exactly "
        "'requirements.txt' in the GitHub repository (same folder as this app) "
        "and add 'openpyxl>=3.1.5'. Then reboot/redeploy the app."
    ) from exc


try:
    import streamlit as st
except Exception:
    st = None


AMOUNT_TOLERANCE = 0.01
DATE_TOLERANCE_DAYS = 2
MACRON_TOKEN_CHARGE = 100.00

PAYMETER_HEADER = [
    "Transaction ID", "Created At", "Updated Date", "Customer Name",
    "Phone Number", "Bank Name", "Disco Name", "District Name",
    "Account Number", "Disco Account Number", "System Fee Per Disco",
    "Meter Number", "Address", "Transaction Amount", "Input Amount",
    "System Charge", "Disco Commission Type", "Disco Amount",
    "Disco System Commission Fee", "Disco Commission Fee Value",
    "Disco System Commission Cap Fee", "RRN", "Reference", "User Type",
    "Earning Partner ID", "Earning Partner Name", "Earning Partner Fee",
    "Earning Partner Commission Type", "Earning Fee Value",
    "Earning Partner Cap Fee", "District Manager Fee",
    "Agent Commission Type", "Agent Commission Fee",
    "Agent Commission Percentage Type Fee Value", "Agent Commission Cap Fee",
    "Bank Commission Type", "Bank Commission Fee", "Bank Commission Fee Value",
    "Bank Commission Fee Cap", "Bank Tax Fee", "Fee Due To System",
    "Message ID", "SMS Delivery Status", "Status Checked",
]


def clean_id(value):
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass
    s = str(value).strip()
    if s.lower() in {"nan", "none", "nat"}:
        return ""
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s


def normalize_meter(value):
    s = clean_id(value)
    if not s:
        return ""
    return (s.lstrip("0") or "0") if s.isdigit() else s.upper()


def to_number(value):
    s = clean_id(value).replace(",", "")
    if not s:
        return np.nan
    s = re.sub(r"[^0-9.\-]", "", s)
    return pd.to_numeric(s, errors="coerce")


def numeric_series(series):
    s = pd.Series(series, dtype="string")
    s = s.str.replace(",", "", regex=False)
    s = s.str.replace(r"[^0-9.\-]", "", regex=True)
    return pd.to_numeric(s, errors="coerce")


def amount_equal(a, b, tolerance=AMOUNT_TOLERANCE):
    try:
        return pd.notna(a) and pd.notna(b) and abs(float(a) - float(b)) <= tolerance
    except Exception:
        return False


def parse_datetime(series, utc=False, dayfirst=True):
    """Parse source dates without swapping ISO YYYY-MM-DD month/day values.

    VPS uses timezone-aware ISO timestamps and is handled with utc=True.
    Paymeter uses ISO YYYY-MM-DD HH:MM:SS and must use dayfirst=False.
    Bank/Macron use DD/MM/YYYY and use the default dayfirst=True.
    """
    if utc:
        d = pd.to_datetime(series, errors="coerce", utc=True)
        try:
            return d.dt.tz_convert("Africa/Lagos").dt.tz_localize(None)
        except Exception:
            return d.dt.tz_localize(None)
    return pd.to_datetime(series, errors="coerce", dayfirst=dayfirst)


def unique_join(values):
    out, seen = [], set()
    for v in values:
        s = clean_id(v)
        if s and s not in seen:
            seen.add(s)
            out.append(s)
    return " | ".join(out)


def most_common_text(values):
    counts = defaultdict(int)
    order = []
    for v in values:
        s = clean_id(v)
        if not s:
            continue
        if s not in counts:
            order.append(s)
        counts[s] += 1
    if not counts:
        return ""
    return max(order, key=lambda x: counts[x])


def split_multi(value):
    return [x.strip() for x in clean_id(value).split("|") if x.strip()]


def prefixed_record(row, prefix):
    return {f"{prefix}{k}": v for k, v in row.items()}


def read_bytes(file_obj):
    if isinstance(file_obj, (str, Path)):
        return Path(file_obj).read_bytes()
    if isinstance(file_obj, (bytes, bytearray)):
        return bytes(file_obj)
    if hasattr(file_obj, "getvalue"):
        return file_obj.getvalue()
    try:
        file_obj.seek(0)
    except Exception:
        pass
    return file_obj.read()


def excel_input(file_obj):
    if isinstance(file_obj, (str, Path)):
        return file_obj
    return io.BytesIO(read_bytes(file_obj))


def suffix_of(file_obj):
    if isinstance(file_obj, (str, Path)):
        return Path(file_obj).suffix.lower()
    return Path(getattr(file_obj, "name", "")).suffix.lower()


def require_columns(df, columns, source):
    missing = [c for c in columns if c not in df.columns]
    if missing:
        raise ValueError(f"{source} missing required column(s): {', '.join(missing)}")


def progress_call(callback, pct, message):
    if callback:
        callback(pct, message)


# ================================================================
# 0. CLEAN PAYMETER BEFORE ANY MERGE
# ================================================================

def _blank_header(value):
    s = clean_id(value)
    return not s or bool(re.fullmatch(r"Unnamed:\s*\d+", s, flags=re.I))


def _normalise_paymeter_header(raw_header):
    header = [str(x).replace("\ufeff", "").strip() for x in raw_header]
    ignored = 0
    while header and _blank_header(header[-1]):
        header.pop()
        ignored += 1
    if header != PAYMETER_HEADER:
        missing = [c for c in PAYMETER_HEADER if c not in header]
        if missing:
            raise ValueError("Paymeter missing required column(s): " + ", ".join(missing))
        diffs = []
        for i, expected in enumerate(PAYMETER_HEADER):
            actual = header[i] if i < len(header) else "<missing>"
            if actual != expected:
                diffs.append(f"column {i+1}: expected '{expected}', found '{actual}'")
            if len(diffs) >= 5:
                break
        raise ValueError(
            "Paymeter named business columns differ from the expected 44-column layout. "
            + "; ".join(diffs)
        )
    return header, ignored


def _find_true_pm_amount_start(row):
    """Find the true Transaction Amount position after an Address spill."""
    tx_idx = PAYMETER_HEADER.index("Transaction Amount")
    rrn_offset = PAYMETER_HEADER.index("RRN") - tx_idx
    ref_offset = PAYMETER_HEADER.index("Reference") - tx_idx
    user_offset = PAYMETER_HEADER.index("User Type") - tx_idx
    addr_idx = PAYMETER_HEADER.index("Address")
    candidates = []

    for pos in range(addr_idx + 1, max(addr_idx + 2, len(row) - 2)):
        if pos + 2 >= len(row):
            break
        tx, inp, charge = to_number(row[pos]), to_number(row[pos + 1]), to_number(row[pos + 2])
        if pd.isna(tx) or pd.isna(inp) or pd.isna(charge):
            continue
        score = 5
        if amount_equal(float(inp) - float(charge), tx):
            score += 10
        if pos + rrn_offset < len(row) and clean_id(row[pos + rrn_offset]):
            score += 2
        if pos + ref_offset < len(row) and clean_id(row[pos + ref_offset]):
            score += 2
        if pos + user_offset < len(row) and clean_id(row[pos + user_offset]):
            score += 1
        candidates.append((score, pos))

    if not candidates:
        return None, 0
    candidates.sort(key=lambda x: (-x[0], x[1]))
    return candidates[0][1], candidates[0][0]


def _repair_paymeter_row(raw_row):
    expected = len(PAYMETER_HEADER)
    addr_idx = PAYMETER_HEADER.index("Address")
    tx_idx = PAYMETER_HEADER.index("Transaction Amount")

    # A perfect-width row can be used directly if the amount triplet is valid.
    if len(raw_row) == expected:
        tx, inp, chg = map(to_number, raw_row[tx_idx:tx_idx + 3])
        if pd.notna(tx) and pd.notna(inp) and pd.notna(chg):
            return list(raw_row), "CLEAN", 0, "HIGH"

    true_pos, score = _find_true_pm_amount_start(raw_row)
    if true_pos is not None:
        first_part = list(raw_row[:addr_idx])
        original_address = raw_row[addr_idx] if addr_idx < len(raw_row) else ""
        tail = list(raw_row[true_pos:])
        tail_needed = expected - tx_idx
        tail = (tail + [""] * tail_needed)[:tail_needed]
        rebuilt = first_part + [original_address] + tail
        rebuilt = (rebuilt + [""] * expected)[:expected]
        deleted = max(true_pos - (addr_idx + 1), 0)
        status = "REPAIRED - ADDRESS SPILL DELETED" if deleted else "CLEAN - EXPORT PADDING NORMALISED"
        return rebuilt, status, deleted, "HIGH" if score >= 15 else "MEDIUM"

    # Last-resort preservation. Keep first/original address and retain row.
    left = list(raw_row[:addr_idx + 1])
    rest = list(raw_row[addr_idx + 1:])
    rebuilt = (left + rest + [""] * expected)[:expected]
    return rebuilt, "REVIEW REQUIRED", max(len(raw_row) - expected, 0), "LOW"


def clean_paymeter(file_obj):
    if suffix_of(file_obj) in {".xlsx", ".xls"}:
        df = pd.read_excel(excel_input(file_obj), dtype=str)
        require_columns(df, PAYMETER_HEADER, "Paymeter")
        df = df[PAYMETER_HEADER].copy()
        df["PM_Cleaning_Status"] = "CLEAN XLS/XLSX"
        df["PM_Address_Spill_Cells_Deleted"] = 0
        df["PM_Cleaning_Confidence"] = "HIGH"
        df["PM_Source_Line"] = np.arange(2, len(df) + 2)
        ignored_header = 0
    else:
        rows = list(csv.reader(io.StringIO(read_bytes(file_obj).decode("utf-8-sig", errors="replace"))))
        if not rows:
            raise ValueError("Paymeter report is empty.")
        _, ignored_header = _normalise_paymeter_header(rows[0])
        records = []
        for line, raw_row in enumerate(rows[1:], start=2):
            if not any(clean_id(x) for x in raw_row):
                continue
            rebuilt, status, deleted, confidence = _repair_paymeter_row(raw_row)
            rec = dict(zip(PAYMETER_HEADER, rebuilt))
            rec["PM_Cleaning_Status"] = status
            rec["PM_Address_Spill_Cells_Deleted"] = deleted
            rec["PM_Cleaning_Confidence"] = confidence
            rec["PM_Source_Line"] = line
            records.append(rec)
        df = pd.DataFrame(records)

    for c in [
        "Transaction Amount", "Input Amount", "System Charge", "System Fee Per Disco",
        "Disco Amount", "Disco System Commission Fee", "Disco Commission Fee Value",
        "Earning Partner Fee", "Bank Commission Fee", "Bank Commission Fee Value",
        "Bank Tax Fee", "Fee Due To System",
    ]:
        if c in df.columns:
            df[c] = numeric_series(df[c])
    for c in ["Transaction ID", "Account Number", "Meter Number", "RRN", "Reference"]:
        df[c] = df[c].map(clean_id)

    # Paymeter exports ISO dates (YYYY-MM-DD), so dayfirst must be False.
    df["PM_DateTime"] = parse_datetime(df["Created At"], dayfirst=False)
    df["PM_Date"] = df["PM_DateTime"].dt.normalize()
    df["PM_Row"] = np.arange(len(df))
    df.attrs["ignored_blank_header_columns"] = ignored_header
    df.attrs["repaired_rows"] = int((pd.to_numeric(df["PM_Address_Spill_Cells_Deleted"], errors="coerce").fillna(0) > 0).sum())
    df.attrs["spill_cells_deleted"] = int(pd.to_numeric(df["PM_Address_Spill_Cells_Deleted"], errors="coerce").fillna(0).sum())
    return df


# ================================================================
# SOURCE READERS
# ================================================================

def read_vps(file_obj):
    df = pd.read_excel(excel_input(file_obj), dtype=str)
    require_columns(df, [
        "session_id", "settlement_ref", "transaction_amount_minor", "settled_amount_minor",
        "charge_amount_minor", "source_acct_name", "source_acct_no", "virtual_acct_no", "created_at",
    ], "VPS")
    for c in ["transaction_amount_minor", "settled_amount_minor", "charge_amount_minor", "vat_amount_minor"]:
        if c in df.columns:
            df[c] = numeric_series(df[c])
    for c in ["session_id", "settlement_ref", "virtual_acct_no", "source_acct_no"]:
        if c in df.columns:
            df[c] = df[c].map(clean_id)
    df["VPS_DateTime"] = parse_datetime(df["created_at"], utc=True)
    df["VPS_Date"] = df["VPS_DateTime"].dt.normalize()
    df["VPS_Row"] = np.arange(len(df))
    return df


def read_macron(file_obj):
    df = pd.read_excel(excel_input(file_obj), dtype=str)
    require_columns(df, ["DATE & TIME", "TRANSACTION ID", "METER NUMBER", "ACCOUNT NUMBER", "AMOUNT", "STATUS", "REFERENCE ID"], "Macron")
    df["AMOUNT"] = numeric_series(df["AMOUNT"])
    for c in ["TRANSACTION ID", "METER NUMBER", "ACCOUNT NUMBER", "REFERENCE ID"]:
        df[c] = df[c].map(clean_id)
    df["MAC_DateTime"] = parse_datetime(df["DATE & TIME"])
    df["MAC_Date"] = df["MAC_DateTime"].dt.normalize()
    df["MAC_Row"] = np.arange(len(df))
    return df


def read_bank(file_obj):
    raw = pd.read_excel(excel_input(file_obj), header=None, dtype=str)
    header_row = None
    for i in range(min(100, len(raw))):
        vals = [clean_id(x) for x in raw.iloc[i].tolist()]
        if "Narration" in vals and "Credit" in vals:
            header_row = i
            break
    if header_row is None:
        raise ValueError("Could not locate Bank transaction header containing Narration and Credit.")
    headers = [clean_id(x) or f"Unnamed_{j}" for j, x in enumerate(raw.iloc[header_row].tolist())]
    df = raw.iloc[header_row + 1:].copy()
    df.columns = headers
    df = df.dropna(how="all").reset_index(drop=True)
    require_columns(df, ["Narration", "Credit"], "Bank")
    for c in ["Debit", "Credit", "Current Balance"]:
        if c in df.columns:
            df[c] = numeric_series(df[c])
    date_col = "Actual Transaction Date" if "Actual Transaction Date" in df.columns else "Post Date"
    df["BANK_DateTime"] = parse_datetime(df[date_col])
    df["BANK_Date"] = df["BANK_DateTime"].dt.normalize()
    df["BANK_Row"] = np.arange(len(df))
    return df


# ================================================================
# GENERIC EXACT ONE-TO-ONE MATCH
# ================================================================

def exact_queue_map(left_series, right_series):
    queues = defaultdict(deque)
    for ri, v in right_series.items():
        key = clean_id(v)
        if key:
            queues[key].append(ri)
    mapping, used = {}, set()
    for li, v in left_series.items():
        key = clean_id(v)
        if key and queues.get(key):
            ri = queues[key].popleft()
            mapping[li] = ri
            used.add(ri)
    return mapping, used


# ================================================================
# 1. MERGE PAYMETER WITH MACRON FIRST
# ================================================================

def merge_paymeter_macron(pay, mac):
    mapping, used_mac = exact_queue_map(pay["Reference"], mac["REFERENCE ID"])
    pay_records = pay.to_dict("index")
    mac_records = mac.to_dict("index")
    rows = []

    for pi, p in pay_records.items():
        mi = mapping.get(pi)
        m = mac_records.get(mi) if mi is not None else None
        rec = {
            "PMMAC_Row": len(rows),
            "PMMAC_Match_Status": "MATCHED" if m is not None else "PAYMETER ONLY - NO MACRON",
            "PMMAC_Match_Method": "Paymeter Reference = Macron REFERENCE ID" if m is not None else "",
        }
        rec.update(prefixed_record(p, "PM_"))
        if m is not None:
            rec.update(prefixed_record(m, "MAC_"))
        rows.append(rec)

    for mi, m in mac_records.items():
        if mi in used_mac:
            continue
        rec = {
            "PMMAC_Row": len(rows),
            "PMMAC_Match_Status": "MACRON ONLY - NO PAYMETER",
            "PMMAC_Match_Method": "",
        }
        rec.update(prefixed_record(m, "MAC_"))
        rows.append(rec)

    out = pd.DataFrame(rows)
    # Ensure key prefixed columns exist even if a source is absent in all rows.
    for c in ["PM_RRN", "PM_Reference", "PM_Account Number", "PM_Meter Number", "PM_Customer Name", "PM_Phone Number", "PM_Input Amount", "PM_Transaction Amount", "PM_System Charge", "PM_PM_Date", "PM_PM_DateTime", "MAC_ACCOUNT NUMBER", "MAC_METER NUMBER", "MAC_AMOUNT", "MAC_REFERENCE ID", "MAC_MAC_Date", "MAC_MAC_DateTime"]:
        if c not in out.columns:
            out[c] = np.nan
    return out


# ================================================================
# 2. MERGE VPS WITH BANK STATEMENT
# ================================================================

def merge_vps_bank(vps, bank, date_tolerance=DATE_TOLERANCE_DAYS):
    credits = bank[bank["Credit"].notna()].copy()
    bank_records = credits.to_dict("index")
    vps_records = vps.to_dict("index")

    # Primary exact session ID from Bank Narration.
    token_queues = defaultdict(deque)
    for bi, b in bank_records.items():
        narration = clean_id(b.get("Narration"))
        for token in re.findall(r"\d{12,}", narration):
            token_queues[token].append(bi)

    vps_to_bank, used_bank = {}, set()
    match_method = {}
    for vi, v in vps_records.items():
        sid = clean_id(v.get("session_id"))
        q = token_queues.get(sid)
        while q and q[0] in used_bank:
            q.popleft()
        if q:
            bi = q.popleft()
            vps_to_bank[vi] = bi
            used_bank.add(bi)
            match_method[vi] = "VPS session_id found in Bank Narration"

    # Secondary unique date + settled amount.
    amount_date = defaultdict(list)
    for bi, b in bank_records.items():
        if bi in used_bank or pd.isna(b.get("Credit")) or pd.isna(b.get("BANK_Date")):
            continue
        amount_date[(pd.Timestamp(b["BANK_Date"]), round(float(b["Credit"]), 2))].append(bi)

    for vi, v in vps_records.items():
        if vi in vps_to_bank or pd.isna(v.get("settled_amount_minor")) or pd.isna(v.get("VPS_Date")):
            continue
        base_date = pd.Timestamp(v["VPS_Date"])
        amount = round(float(v["settled_amount_minor"]), 2)
        candidates = []
        for d in range(-date_tolerance, date_tolerance + 1):
            candidates.extend([x for x in amount_date.get((base_date + pd.Timedelta(days=d), amount), []) if x not in used_bank])
        candidates = list(dict.fromkeys(candidates))
        if len(candidates) == 1:
            bi = candidates[0]
            vps_to_bank[vi] = bi
            used_bank.add(bi)
            match_method[vi] = "Unique VPS settled amount + date"

    rows = []
    for vi, v in vps_records.items():
        bi = vps_to_bank.get(vi)
        b = bank_records.get(bi) if bi is not None else None
        rec = {
            "VPSBANK_Row": len(rows),
            "VPSBANK_Match_Status": "MATCHED" if b is not None else "VPS ONLY - NO BANK",
            "VPSBANK_Match_Method": match_method.get(vi, ""),
        }
        rec.update(prefixed_record(v, "VPS_"))
        if b is not None:
            rec.update(prefixed_record(b, "BANK_"))
        rows.append(rec)

    for bi, b in bank_records.items():
        if bi in used_bank:
            continue
        rec = {
            "VPSBANK_Row": len(rows),
            "VPSBANK_Match_Status": "BANK ONLY - NO VPS",
            "VPSBANK_Match_Method": "",
        }
        rec.update(prefixed_record(b, "BANK_"))
        rows.append(rec)

    out = pd.DataFrame(rows)
    for c in ["VPS_virtual_acct_no", "VPS_settlement_ref", "VPS_session_id", "VPS_transaction_amount_minor", "VPS_settled_amount_minor", "VPS_charge_amount_minor", "VPS_VPS_Date", "VPS_VPS_DateTime", "BANK_Credit", "BANK_Narration", "BANK_BANK_Date", "BANK_BANK_DateTime"]:
        if c not in out.columns:
            out[c] = np.nan
    return out


# ================================================================
# 3. EXTRACT KYC FROM BOTH MERGED REPORTS — VPS ACCOUNT FIRST
# ================================================================

def build_kyc(vps_bank_merge, pm_mac_merge):
    vps_rows = vps_bank_merge[vps_bank_merge["VPS_virtual_acct_no"].map(clean_id).ne("")].copy()
    pm_rows = pm_mac_merge.copy()

    # Maps from PM-Mac merge for fast history retrieval.
    rrn_to_rows = defaultdict(list)
    pm_account_to_rows = defaultdict(list)
    for idx, r in pm_rows.iterrows():
        rrn = clean_id(r.get("PM_RRN"))
        pa = clean_id(r.get("PM_Account Number"))
        if rrn:
            rrn_to_rows[rrn].append(idx)
        if pa:
            pm_account_to_rows[pa].append(idx)

    records = []
    for va, vg in vps_rows.groupby(vps_rows["VPS_virtual_acct_no"].map(clean_id), sort=True):
        linked_pm_rows = set()
        proven_pm_accounts = set()
        methods = []

        # Primary KYC evidence: VPS settlement_ref = Paymeter RRN.
        for ref in vg["VPS_settlement_ref"].map(clean_id):
            if not ref:
                continue
            for pidx in rrn_to_rows.get(ref, []):
                linked_pm_rows.add(pidx)
                pa = clean_id(pm_rows.at[pidx, "PM_Account Number"])
                if pa:
                    proven_pm_accounts.add(pa)
        if linked_pm_rows:
            methods.append("VPS settlement_ref = Paymeter RRN")

        # Fallback identity: VPS virtual account = Paymeter account.
        if not proven_pm_accounts and va in pm_account_to_rows:
            proven_pm_accounts.add(va)
            methods.append("VPS virtual account = Paymeter Account Number")

        # Once an account is proven, use its full PM/Macron history to enrich KYC.
        for pa in list(proven_pm_accounts):
            linked_pm_rows.update(pm_account_to_rows.get(pa, []))

        hist = pm_rows.loc[sorted(linked_pm_rows)] if linked_pm_rows else pm_rows.iloc[0:0]

        pay_accounts = unique_join(hist.get("PM_Account Number", pd.Series(dtype=object)))
        customer_names = unique_join(hist.get("PM_Customer Name", pd.Series(dtype=object)))
        customer_name = most_common_text(hist.get("PM_Customer Name", pd.Series(dtype=object)))
        phones = unique_join(hist.get("PM_Phone Number", pd.Series(dtype=object)))
        pay_meters = unique_join(hist.get("PM_Meter Number", pd.Series(dtype=object)))
        mac_accounts = unique_join(hist.get("MAC_ACCOUNT NUMBER", pd.Series(dtype=object)))
        mac_meters = unique_join(hist.get("MAC_METER NUMBER", pd.Series(dtype=object)))

        conflicts = []
        if len(split_multi(pay_accounts)) > 1:
            conflicts.append("Multiple Paymeter accounts")
        if len(split_multi(customer_names)) > 1:
            conflicts.append("Multiple Paymeter customer names")
        if len(split_multi(mac_accounts)) > 1:
            conflicts.append("Multiple Macron accounts")
        if len(split_multi(mac_meters)) > 1:
            conflicts.append("Multiple Macron meters")

        records.append({
            "KYC_ID": va,
            "Customer Name": customer_name or most_common_text(vg.get("VPS_source_acct_name", pd.Series(dtype=object))),
            "VPS Virtual Account Number": va,
            "VPS Source Account Name(s)": unique_join(vg.get("VPS_source_acct_name", pd.Series(dtype=object))),
            "VPS Source Account Number(s)": unique_join(vg.get("VPS_source_acct_no", pd.Series(dtype=object))),
            "Paymeter Account Number(s)": pay_accounts,
            "Paymeter Customer Name(s)": customer_names,
            "Customer Phone Number(s)": phones,
            "Paymeter Meter Number(s)": pay_meters,
            "Macron Account Number(s)": mac_accounts,
            "Macron Meter Number(s)": mac_meters,
            "VPS Transaction Count": len(vg),
            "Linked PM-Macron Row Count": len(linked_pm_rows),
            "KYC Link Method": " | ".join(methods),
            "KYC Conflict": "; ".join(conflicts),
            "KYC Status": "COMPLETE CORE KYC" if pay_accounts and mac_accounts else "INCOMPLETE KYC",
        })

    return pd.DataFrame(records)


def _unique_map_from_multivalue(kyc, column, normalizer=clean_id):
    holder = defaultdict(set)
    for _, r in kyc.iterrows():
        kid = clean_id(r.get("KYC_ID"))
        for item in split_multi(r.get(column)):
            key = normalizer(item)
            if key and kid:
                holder[key].add(kid)
    return {k: next(iter(v)) for k, v in holder.items() if len(v) == 1}


def build_kyc_maps(kyc, vps_bank_merge):
    kyc_by_id = {clean_id(r["KYC_ID"]): r for _, r in kyc.iterrows() if clean_id(r["KYC_ID"])}
    pay_account = _unique_map_from_multivalue(kyc, "Paymeter Account Number(s)")
    mac_account = _unique_map_from_multivalue(kyc, "Macron Account Number(s)")
    mac_meter = _unique_map_from_multivalue(kyc, "Macron Meter Number(s)", normalize_meter)

    rrn_holder = defaultdict(set)
    for _, r in vps_bank_merge.iterrows():
        ref = clean_id(r.get("VPS_settlement_ref"))
        kid = clean_id(r.get("VPS_virtual_acct_no"))
        if ref and kid:
            rrn_holder[ref].add(kid)
    rrn = {k: next(iter(v)) for k, v in rrn_holder.items() if len(v) == 1}
    return kyc_by_id, rrn, pay_account, mac_account, mac_meter


# ================================================================
# 4. MERGE KYC INTO PAYMETER-MACRON MERGE
# ================================================================

def attach_kyc_to_pm_mac(pm_mac_merge, kyc, vps_bank_merge):
    kyc_by_id, rrn_map, pay_map, mac_acct_map, mac_meter_map = build_kyc_maps(kyc, vps_bank_merge)
    rows = []
    for r in pm_mac_merge.to_dict("records"):
        kid, method = "", ""
        rrn = clean_id(r.get("PM_RRN"))
        pa = clean_id(r.get("PM_Account Number"))
        ma = clean_id(r.get("MAC_ACCOUNT NUMBER"))
        mm = normalize_meter(r.get("MAC_METER NUMBER"))

        if rrn and rrn in rrn_map:
            kid, method = rrn_map[rrn], "PM RRN -> VPS settlement_ref -> KYC"
        elif pa and pa in pay_map:
            kid, method = pay_map[pa], "Paymeter Account Number -> KYC"
        elif ma and ma in mac_acct_map:
            kid, method = mac_acct_map[ma], "Macron Account Number -> KYC"
        elif mm and mm in mac_meter_map:
            kid, method = mac_meter_map[mm], "Macron Meter Number -> KYC"

        kr = kyc_by_id.get(kid)
        r["KYC_ID"] = kid
        r["KYC_Match_Method"] = method or "NOT LINKED"
        r["KYC_Customer_Name"] = clean_id(kr.get("Customer Name")) if kr is not None else ""
        r["KYC_VPS_Virtual_Account"] = clean_id(kr.get("VPS Virtual Account Number")) if kr is not None else ""
        r["KYC_Paymeter_Account(s)"] = clean_id(kr.get("Paymeter Account Number(s)")) if kr is not None else ""
        r["KYC_Macron_Account(s)"] = clean_id(kr.get("Macron Account Number(s)")) if kr is not None else ""
        r["KYC_Meter_Number(s)"] = clean_id(kr.get("Macron Meter Number(s)")) if kr is not None else ""
        rows.append(r)
    return pd.DataFrame(rows)


# ================================================================
# 5. MERGE KYC INTO VPS-BANK MERGE
# ================================================================

def attach_kyc_to_vps_bank(vps_bank_merge, kyc):
    kyc_by_id = {clean_id(r["KYC_ID"]): r for _, r in kyc.iterrows() if clean_id(r["KYC_ID"])}
    rows = []
    for r in vps_bank_merge.to_dict("records"):
        kid = clean_id(r.get("VPS_virtual_acct_no"))
        kr = kyc_by_id.get(kid)
        r["KYC_ID"] = kid if kr is not None else ""
        r["KYC_Match_Method"] = "VPS virtual_acct_no -> KYC" if kr is not None else "NOT LINKED"
        r["KYC_Customer_Name"] = clean_id(kr.get("Customer Name")) if kr is not None else ""
        r["KYC_Paymeter_Account(s)"] = clean_id(kr.get("Paymeter Account Number(s)")) if kr is not None else ""
        r["KYC_Macron_Account(s)"] = clean_id(kr.get("Macron Account Number(s)")) if kr is not None else ""
        r["KYC_Meter_Number(s)"] = clean_id(kr.get("Macron Meter Number(s)")) if kr is not None else ""
        rows.append(r)
    return pd.DataFrame(rows)


# ================================================================
# 6. RECONCILE THE TWO KYC-ENRICHED MERGES
# ================================================================

def _candidate_date_pm_mac(r):
    for key in ["PM_PM_Date", "MAC_MAC_Date"]:
        v = r.get(key)
        if pd.notna(v):
            return pd.Timestamp(v)
    return pd.NaT


def reconcile_merged_reports(vps_bank_kyc, pm_mac_kyc, date_tolerance=DATE_TOLERANCE_DAYS):
    left = vps_bank_kyc.copy()
    right = pm_mac_kyc.copy()
    left_records = left.to_dict("index")
    right_records = right.to_dict("index")

    # Primary exact transactional bridge: VPS settlement_ref = Paymeter RRN.
    mapping, used_right = exact_queue_map(left["VPS_settlement_ref"], right["PM_RRN"])
    method = {li: "VPS settlement_ref = Paymeter RRN" for li in mapping}

    # Secondary KYC+amount+date index. It is used ONLY for remaining rows.
    pm_gross_index = defaultdict(set)
    mac_token_index = defaultdict(set)
    for ri, r in right_records.items():
        if ri in used_right:
            continue
        kid = clean_id(r.get("KYC_ID"))
        if not kid:
            continue
        pm_input = r.get("PM_Input Amount")
        mac_amt = r.get("MAC_AMOUNT")
        if pd.notna(pm_input):
            pm_gross_index[(kid, round(float(pm_input), 2))].add(ri)
        if pd.notna(mac_amt):
            mac_token_index[(kid, round(float(mac_amt), 2))].add(ri)

    for li, l in left_records.items():
        if li in mapping:
            continue
        kid = clean_id(l.get("KYC_ID"))
        gross = l.get("VPS_transaction_amount_minor")
        ldate = l.get("VPS_VPS_Date")
        if not kid or pd.isna(gross) or pd.isna(ldate):
            continue

        gross_key = (kid, round(float(gross), 2))
        expected_token = float(gross) - MACRON_TOKEN_CHARGE
        token_key = (kid, round(expected_token, 2))
        c_pm = set(pm_gross_index.get(gross_key, set()))
        c_mac = set(mac_token_index.get(token_key, set()))
        candidates = (c_pm & c_mac) if c_pm and c_mac else (c_pm | c_mac)
        candidates = {ri for ri in candidates if ri not in used_right}

        dated = []
        for ri in candidates:
            rdate = _candidate_date_pm_mac(right_records[ri])
            if pd.isna(rdate):
                continue
            dd = abs((pd.Timestamp(ldate) - rdate).days)
            if dd <= date_tolerance:
                dated.append((dd, ri))
        dated.sort()
        if len(dated) == 1:
            ri = dated[0][1]
            mapping[li] = ri
            used_right.add(ri)
            method[li] = "KYC + amount + date fallback"

    rows = []
    used_left = set()
    for li, l in left_records.items():
        used_left.add(li)
        ri = mapping.get(li)
        r = right_records.get(ri) if ri is not None else None
        rows.append(_final_recon_row(l, r, method.get(li, "")))

    for ri, r in right_records.items():
        if ri not in used_right:
            rows.append(_final_recon_row(None, r, ""))

    out = pd.DataFrame(rows)
    out.insert(0, "FINAL_Row_ID", [f"REC-{i+1:07d}" for i in range(len(out))])
    return out


def _safe_diff(a, b):
    return float(a) - float(b) if pd.notna(a) and pd.notna(b) else np.nan


def _final_recon_row(l, r, match_method):
    has_vps = l is not None and clean_id(l.get("VPS_VPS_Row")) != ""
    has_bank = l is not None and clean_id(l.get("BANK_BANK_Row")) != ""
    has_pm = r is not None and clean_id(r.get("PM_PM_Row")) != ""
    has_mac = r is not None and clean_id(r.get("MAC_MAC_Row")) != ""

    left_kyc = clean_id(l.get("KYC_ID")) if l else ""
    right_kyc = clean_id(r.get("KYC_ID")) if r else ""
    kyc = left_kyc or right_kyc
    kyc_status = "NO KYC"
    if left_kyc and right_kyc:
        kyc_status = "KYC ALIGNED" if left_kyc == right_kyc else "KYC MISMATCH"
    elif kyc:
        kyc_status = "KYC ON ONE SIDE"

    vps_gross = l.get("VPS_transaction_amount_minor") if has_vps else np.nan
    vps_settled = l.get("VPS_settled_amount_minor") if has_vps else np.nan
    bank_credit = l.get("BANK_Credit") if has_bank else np.nan
    pm_input = r.get("PM_Input Amount") if has_pm else np.nan
    pm_token = r.get("PM_Transaction Amount") if has_pm else np.nan
    mac_token = r.get("MAC_AMOUNT") if has_mac else np.nan
    expected_mac = float(vps_gross) - MACRON_TOKEN_CHARGE if pd.notna(vps_gross) else np.nan

    d_bank_vps = _safe_diff(bank_credit, vps_settled)
    d_vps_pm = _safe_diff(vps_gross, pm_input)
    d_pm_mac = _safe_diff(pm_token, mac_token)
    d_expected_mac = _safe_diff(expected_mac, mac_token)

    amount_checks = []
    if has_bank and has_vps and pd.notna(d_bank_vps):
        amount_checks.append(abs(d_bank_vps) <= AMOUNT_TOLERANCE)
    if has_vps and has_pm and pd.notna(d_vps_pm):
        amount_checks.append(abs(d_vps_pm) <= AMOUNT_TOLERANCE)
    if has_vps and has_mac and pd.notna(d_expected_mac):
        amount_checks.append(abs(d_expected_mac) <= AMOUNT_TOLERANCE)
    if has_pm and has_mac and pd.notna(d_pm_mac):
        amount_checks.append(abs(d_pm_mac) <= AMOUNT_TOLERANCE)
    amount_status = "AMOUNTS ALIGNED" if amount_checks and all(amount_checks) else ("AMOUNT DIFFERENCE" if amount_checks else "NOT COMPARABLE")

    if has_vps and not has_mac:
        status = "CUSTOMER PAID / VPS RECEIVED - NO MACRON TOKEN"
    elif has_mac and not has_vps:
        status = "MACRON TOKEN - NO VPS PAYMENT"
    elif has_vps and has_mac and kyc_status == "KYC MISMATCH":
        status = "KYC MISMATCH"
    elif has_vps and has_bank and has_pm and has_mac and amount_status == "AMOUNTS ALIGNED":
        status = "FULLY RECONCILED"
    elif has_vps and has_mac and amount_status == "AMOUNT DIFFERENCE":
        status = "VPS / MACRON VALUE DIFFERENCE"
    elif has_vps and not has_bank:
        status = "VPS PRESENT - BANK MISSING"
    elif has_vps and not has_pm:
        status = "VPS PRESENT - PAYMETER MISSING"
    else:
        status = "PARTIAL / REVIEW"

    rec = {
        "FINAL_Reconciliation_Status": status,
        "FINAL_Match_Method": match_method,
        "FINAL_KYC_ID": kyc,
        "FINAL_KYC_Status": kyc_status,
        "FINAL_Customer_Name": (clean_id(l.get("KYC_Customer_Name")) if l else "") or (clean_id(r.get("KYC_Customer_Name")) if r else "") or (clean_id(r.get("PM_Customer Name")) if r else ""),
        "Bank Present?": "YES" if has_bank else "NO",
        "VPS Present?": "YES" if has_vps else "NO",
        "Paymeter Present?": "YES" if has_pm else "NO",
        "Macron Present?": "YES" if has_mac else "NO",
        "VPS Customer Amount Paid": vps_gross,
        "Expected Macron Token (VPS Paid - 100)": expected_mac,
        "Bank Credit": bank_credit,
        "VPS Settled Amount": vps_settled,
        "Paymeter Input Amount": pm_input,
        "Paymeter Token/Net Amount": pm_token,
        "Macron Token Amount": mac_token,
        "Difference Bank Credit - VPS Settled": d_bank_vps,
        "Difference VPS Paid - Paymeter Input": d_vps_pm,
        "Difference Paymeter Token - Macron Token": d_pm_mac,
        "Difference Expected Macron - Actual Macron": d_expected_mac,
        "FINAL_Amount_Status": amount_status,
    }
    if l:
        rec.update(prefixed_record(l, "LEFT_"))
    if r:
        rec.update(prefixed_record(r, "RIGHT_"))
    return rec


# ================================================================
# 7. VPS AGAINST MACRON WITH KYC ATTACHED
# ================================================================

def build_vps_macron_comparison(final_recon):
    cols = [
        "FINAL_Row_ID", "FINAL_KYC_ID", "FINAL_Customer_Name", "FINAL_KYC_Status",
        "LEFT_VPS_virtual_acct_no", "LEFT_VPS_session_id", "LEFT_VPS_settlement_ref",
        "RIGHT_PM_Account Number", "RIGHT_PM_RRN", "RIGHT_PM_Reference",
        "RIGHT_MAC_ACCOUNT NUMBER", "RIGHT_MAC_METER NUMBER", "RIGHT_MAC_REFERENCE ID",
        "VPS Customer Amount Paid", "Expected Macron Token (VPS Paid - 100)",
        "Macron Token Amount", "Difference Expected Macron - Actual Macron",
        "VPS Present?", "Macron Present?", "FINAL_Match_Method", "FINAL_Reconciliation_Status",
    ]
    for c in cols:
        if c not in final_recon.columns:
            final_recon[c] = np.nan
    out = final_recon[(final_recon["VPS Present?"] == "YES") | (final_recon["Macron Present?"] == "YES")][cols].copy()
    out["VPS_vs_Macron_Status"] = np.where(
        (out["VPS Present?"] == "YES") & (out["Macron Present?"] == "NO"),
        "PAID / NO TOKEN",
        np.where(
            (out["VPS Present?"] == "NO") & (out["Macron Present?"] == "YES"),
            "TOKEN / NO VPS PAYMENT",
            np.where(
                pd.to_numeric(out["Difference Expected Macron - Actual Macron"], errors="coerce").abs().fillna(np.inf) <= AMOUNT_TOLERANCE,
                "VALUE MATCH",
                "VALUE DIFFERENCE",
            ),
        ),
    )
    return out


# ================================================================
# 8. COMPLETE CUSTOMER SUMMARY — NO SOURCE FIGURE MAY DISAPPEAR
# ================================================================

def _summary_key(r):
    kid = clean_id(r.get("FINAL_KYC_ID"))
    if kid:
        return f"KYC::{kid}", "KYC CUSTOMER"
    pa = clean_id(r.get("RIGHT_PM_Account Number"))
    if pa:
        return f"UNLINKED-PAYMETER::{pa}", "UNLINKED PAYMETER"
    ma = clean_id(r.get("RIGHT_MAC_ACCOUNT NUMBER"))
    if ma:
        return f"UNLINKED-MACRON::{ma}", "UNLINKED MACRON"
    mm = normalize_meter(r.get("RIGHT_MAC_METER NUMBER"))
    if mm:
        return f"UNLINKED-MACRON-METER::{mm}", "UNLINKED MACRON METER"
    if clean_id(r.get("Bank Present?")) == "YES":
        return "UNIDENTIFIED-BANK-CREDIT", "UNIDENTIFIED BANK CREDIT"
    return f"UNIDENTIFIED::{clean_id(r.get('FINAL_Row_ID'))}", "UNIDENTIFIED"


def build_customer_summary(final_recon, kyc):
    work = final_recon.copy()
    keys = work.apply(_summary_key, axis=1)
    work["Summary Customer Key"] = [x[0] for x in keys]
    work["Customer Classification"] = [x[1] for x in keys]
    kyc_lookup = {clean_id(r["KYC_ID"]): r for _, r in kyc.iterrows() if clean_id(r["KYC_ID"])}
    rows = []

    for key, g in work.groupby("Summary Customer Key", sort=True):
        kid = most_common_text(g["FINAL_KYC_ID"])
        kr = kyc_lookup.get(kid)
        vps_count = int((g["VPS Present?"] == "YES").sum())
        row = {
            "Summary Customer Key": key,
            "Customer Classification": most_common_text(g["Customer Classification"]),
            "KYC_ID": kid,
            "Customer Name": clean_id(kr.get("Customer Name")) if kr is not None else (most_common_text(g["FINAL_Customer_Name"]) or "UNIDENTIFIED CUSTOMER"),
            "VPS Virtual Account Number": clean_id(kr.get("VPS Virtual Account Number")) if kr is not None else unique_join(g.get("LEFT_VPS_virtual_acct_no", pd.Series(dtype=object))),
            "Paymeter Account Number(s)": clean_id(kr.get("Paymeter Account Number(s)")) if kr is not None else unique_join(g.get("RIGHT_PM_Account Number", pd.Series(dtype=object))),
            "Paymeter Meter Number(s)": clean_id(kr.get("Paymeter Meter Number(s)")) if kr is not None else unique_join(g.get("RIGHT_PM_Meter Number", pd.Series(dtype=object))),
            "Macron Account Number(s)": clean_id(kr.get("Macron Account Number(s)")) if kr is not None else unique_join(g.get("RIGHT_MAC_ACCOUNT NUMBER", pd.Series(dtype=object))),
            "Macron Meter Number(s)": clean_id(kr.get("Macron Meter Number(s)")) if kr is not None else unique_join(g.get("RIGHT_MAC_METER NUMBER", pd.Series(dtype=object))),
            "Total Bank Credit Count": int((g["Bank Present?"] == "YES").sum()),
            "Total VPS Payment Count": vps_count,
            "Total Paymeter Transaction Count": int((g["Paymeter Present?"] == "YES").sum()),
            "Total Macron Vend Count": int((g["Macron Present?"] == "YES").sum()),
            "Total Bank Credit": pd.to_numeric(g["Bank Credit"], errors="coerce").sum(),
            "Total VPS Customer Amount Paid": pd.to_numeric(g["VPS Customer Amount Paid"], errors="coerce").sum(),
            "Total VPS Settled Amount": pd.to_numeric(g["VPS Settled Amount"], errors="coerce").sum(),
            "Expected Total Macron Token": pd.to_numeric(g["Expected Macron Token (VPS Paid - 100)"], errors="coerce").sum(),
            "Total Paymeter Input Amount": pd.to_numeric(g["Paymeter Input Amount"], errors="coerce").sum(),
            "Total Paymeter Token/Net Amount": pd.to_numeric(g["Paymeter Token/Net Amount"], errors="coerce").sum(),
            "Total Macron Token Amount": pd.to_numeric(g["Macron Token Amount"], errors="coerce").sum(),
            "VPS Paid / No Token Count": int(((g["VPS Present?"] == "YES") & (g["Macron Present?"] == "NO")).sum()),
            "Token / No VPS Payment Count": int(((g["Macron Present?"] == "YES") & (g["VPS Present?"] == "NO")).sum()),
            "Difference Expected Macron - Actual Macron": pd.to_numeric(g["Expected Macron Token (VPS Paid - 100)"], errors="coerce").sum() - pd.to_numeric(g["Macron Token Amount"], errors="coerce").sum(),
            "Difference VPS Paid - Paymeter Input": pd.to_numeric(g["VPS Customer Amount Paid"], errors="coerce").sum() - pd.to_numeric(g["Paymeter Input Amount"], errors="coerce").sum(),
            "Difference Paymeter Token - Macron Token": pd.to_numeric(g["Paymeter Token/Net Amount"], errors="coerce").sum() - pd.to_numeric(g["Macron Token Amount"], errors="coerce").sum(),
            "KYC Status": clean_id(kr.get("KYC Status")) if kr is not None else "UNLINKED / NO FORMAL KYC",
        }
        rows.append(row)
    return pd.DataFrame(rows)


def build_summary_control(summary, bank, vps, pay, mac):
    def s(series):
        return pd.to_numeric(series, errors="coerce").sum()
    controls = [
        ("Bank", "Credit Count", int(bank["Credit"].notna().sum()), int(pd.to_numeric(summary["Total Bank Credit Count"], errors="coerce").fillna(0).sum()), "COUNT"),
        ("Bank", "Credit Amount", s(bank.loc[bank["Credit"].notna(), "Credit"]), s(summary["Total Bank Credit"]), "AMOUNT"),
        ("VPS", "Transaction Count", len(vps), int(pd.to_numeric(summary["Total VPS Payment Count"], errors="coerce").fillna(0).sum()), "COUNT"),
        ("VPS", "transaction_amount_minor", s(vps["transaction_amount_minor"]), s(summary["Total VPS Customer Amount Paid"]), "AMOUNT"),
        ("VPS", "settled_amount_minor", s(vps["settled_amount_minor"]), s(summary["Total VPS Settled Amount"]), "AMOUNT"),
        ("Paymeter", "Transaction Count", len(pay), int(pd.to_numeric(summary["Total Paymeter Transaction Count"], errors="coerce").fillna(0).sum()), "COUNT"),
        ("Paymeter", "Input Amount", s(pay["Input Amount"]), s(summary["Total Paymeter Input Amount"]), "AMOUNT"),
        ("Paymeter", "Transaction Amount", s(pay["Transaction Amount"]), s(summary["Total Paymeter Token/Net Amount"]), "AMOUNT"),
        ("Macron", "Vend Count", len(mac), int(pd.to_numeric(summary["Total Macron Vend Count"], errors="coerce").fillna(0).sum()), "COUNT"),
        ("Macron", "AMOUNT", s(mac["AMOUNT"]), s(summary["Total Macron Token Amount"]), "AMOUNT"),
        ("VPS Derived", "Expected Macron Token = transaction_amount_minor - 100", s(vps["transaction_amount_minor"]) - MACRON_TOKEN_CHARGE * int(vps["transaction_amount_minor"].notna().sum()), s(summary["Expected Total Macron Token"]), "AMOUNT"),
    ]
    rows = []
    for source, metric, source_total, summary_total, typ in controls:
        diff = float(summary_total) - float(source_total)
        okay = abs(diff) < 0.5 if typ == "COUNT" else abs(diff) <= AMOUNT_TOLERANCE
        rows.append({"Source": source, "Metric": metric, "Source Total": source_total, "Customer Summary Total": summary_total, "Difference": diff, "Type": typ, "Status": "MATCH" if okay else "MISMATCH"})
    return pd.DataFrame(rows)


def validate_tieout(control):
    bad = control[control["Status"] != "MATCH"]
    if len(bad):
        detail = "; ".join(f"{r['Source']} {r['Metric']} diff={r['Difference']}" for _, r in bad.iterrows())
        raise ValueError("Customer Summary does not tie to source reports: " + detail)


# ================================================================
# MANAGEMENT REPORT / EXCEPTION ANALYSIS
# ================================================================

def _num_sum(series):
    return pd.to_numeric(series, errors="coerce").sum()


def _pct(numerator, denominator):
    if denominator in (0, None) or pd.isna(denominator):
        return 0.0
    return (float(numerator) / float(denominator)) * 100.0


def _period_bounds(df, datetime_col, filter_mask=None):
    if datetime_col not in df.columns:
        return None, None
    s = pd.to_datetime(df[datetime_col], errors="coerce")
    if filter_mask is not None:
        s = s[filter_mask]
    s = s.dropna()
    if s.empty:
        return None, None
    return s.min(), s.max()


def _fmt_period(value):
    if value is None or pd.isna(value):
        return ""
    return pd.Timestamp(value).strftime("%d-%b-%Y")


def build_management_report(
    bank,
    vps,
    pay,
    mac,
    pm_mac_merge,
    vps_bank_merge,
    kyc,
    final_recon,
    vps_macron,
    customer_summary,
    summary_control,
    source_coverage,
):
    """
    Comprehensive management report covering:
      * source/reporting periods and volumes;
      * source financial totals;
      * Paymeter↔Macron and VPS↔Bank merge performance;
      * end-to-end reconciliation and amount alignment;
      * VPS↔Macron fulfilment/value accuracy;
      * KYC coverage;
      * operational/financial exceptions;
      * Paymeter/data-quality controls;
      * Customer Summary tie-out and source coverage.

    The report is deliberately long-form: one KPI per row with an explanation,
    calculation/source and management attention note.
    """
    rows = []

    def add(section, parameter, value, unit="", status="INFO", interpretation="", calculation="", attention=""):
        rows.append({
            "Section": section,
            "Parameter": parameter,
            "Value": value,
            "Unit": unit,
            "Status": status,
            "Interpretation": interpretation,
            "Calculation / Source": calculation,
            "Management Attention / Action": attention,
        })

    # ------------------------------------------------------------
    # REPORTING PERIOD / SOURCE COVERAGE DATES
    # ------------------------------------------------------------
    bank_credit_mask = bank["Credit"].notna()
    b_start, b_end = _period_bounds(bank, "BANK_DateTime", bank_credit_mask)
    v_start, v_end = _period_bounds(vps, "VPS_DateTime")
    p_start, p_end = _period_bounds(pay, "PM_DateTime")
    m_start, m_end = _period_bounds(mac, "MAC_DateTime")

    periods = {
        "Bank Credits": (b_start, b_end),
        "VPS": (v_start, v_end),
        "Paymeter": (p_start, p_end),
        "Macron": (m_start, m_end),
    }
    valid_ends = [e for _, e in periods.values() if e is not None and not pd.isna(e)]
    common_end = min(valid_ends) if valid_ends else None
    end_dates = {_fmt_period(e) for _, e in periods.values() if e is not None and not pd.isna(e)}

    for source, (start_dt, end_dt) in periods.items():
        add(
            "01 Reporting Period",
            f"{source} period",
            f"{_fmt_period(start_dt)} to {_fmt_period(end_dt)}" if start_dt is not None else "Not available",
            "date range",
            "INFO",
            f"Transaction date coverage detected in the {source} source.",
            f"Minimum and maximum parsed transaction dates in {source}.",
        )

    add(
        "01 Reporting Period",
        "Latest common date across all four sources",
        _fmt_period(common_end),
        "date",
        "INFO",
        "Latest date up to which all four reports have data coverage.",
        "Minimum of the four source ending dates.",
        "Use this date for strict like-for-like period analysis when source reports end on different dates.",
    )
    add(
        "01 Reporting Period",
        "Source ending dates aligned",
        "YES" if len(end_dates) <= 1 else "NO",
        "control",
        "GOOD" if len(end_dates) <= 1 else "ATTENTION",
        "Shows whether all four source reports end on the same calendar date.",
        "Comparison of Bank, VPS, Paymeter and Macron maximum transaction dates.",
        "If NO, some unmatched transactions may be timing/coverage exceptions rather than processing failures.",
    )

    # ------------------------------------------------------------
    # SOURCE VOLUMES / FINANCIAL TOTALS
    # ------------------------------------------------------------
    bank_count = int(bank_credit_mask.sum())
    bank_total = _num_sum(bank.loc[bank_credit_mask, "Credit"])
    vps_count = len(vps)
    vps_gross = _num_sum(vps["transaction_amount_minor"])
    vps_settled = _num_sum(vps["settled_amount_minor"])
    vps_charges = _num_sum(vps["charge_amount_minor"])
    valid_vps_amount_count = int(pd.to_numeric(vps["transaction_amount_minor"], errors="coerce").notna().sum())
    expected_macron_total = vps_gross - (MACRON_TOKEN_CHARGE * valid_vps_amount_count)
    pay_count = len(pay)
    pay_input = _num_sum(pay["Input Amount"])
    pay_charge = _num_sum(pay["System Charge"])
    pay_token = _num_sum(pay["Transaction Amount"])
    mac_count = len(mac)
    mac_amount = _num_sum(mac["AMOUNT"])

    source_kpis = [
        ("Bank credit transactions", bank_count, "count", "Number of credit rows in Providus Bank."),
        ("Total Bank credits", bank_total, "amount", "Sum of Bank Credit rows."),
        ("VPS transactions", vps_count, "count", "Number of VPS transactions."),
        ("Total customer amount paid on VPS", vps_gross, "amount", "Sum of VPS transaction_amount_minor."),
        ("Total VPS settled amount", vps_settled, "amount", "Sum of VPS settled_amount_minor."),
        ("Total VPS charges", vps_charges, "amount", "Sum of VPS charge_amount_minor."),
        ("Expected Macron token from VPS", expected_macron_total, "amount", "Σ(transaction_amount_minor - ₦100) for VPS rows with an amount."),
        ("Paymeter transactions", pay_count, "count", "Number of cleaned Paymeter rows."),
        ("Total Paymeter input amount", pay_input, "amount", "Sum of cleaned Paymeter Input Amount."),
        ("Total Paymeter system charges", pay_charge, "amount", "Sum of cleaned Paymeter System Charge."),
        ("Total Paymeter token/net amount", pay_token, "amount", "Sum of cleaned Paymeter Transaction Amount."),
        ("Macron vend transactions", mac_count, "count", "Number of Macron rows."),
        ("Total Macron token amount", mac_amount, "amount", "Sum of Macron AMOUNT."),
    ]
    for parameter, value, unit, calc in source_kpis:
        add("02 Source Volume & Financial Totals", parameter, value, unit, "INFO", calc, calc)

    add(
        "02 Source Volume & Financial Totals",
        "VPS gross less VPS settled",
        vps_gross - vps_settled,
        "amount",
        "INFO",
        "Aggregate difference between customer amount paid on VPS and VPS settled amount.",
        "Total VPS transaction_amount_minor - Total VPS settled_amount_minor.",
    )
    add(
        "02 Source Volume & Financial Totals",
        "Paymeter input less Paymeter token/net",
        pay_input - pay_token,
        "amount",
        "INFO",
        "Aggregate reduction between Paymeter gross input and Paymeter token/net value.",
        "Total Paymeter Input Amount - Total Paymeter Transaction Amount.",
    )

    # ------------------------------------------------------------
    # FIRST MERGE: PAYMETER ↔ MACRON
    # ------------------------------------------------------------
    pm_status = pm_mac_merge["PMMAC_Match_Status"].value_counts()
    pm_matched = int(pm_status.get("MATCHED", 0))
    pm_only = int(pm_status.get("PAYMETER ONLY - NO MACRON", 0))
    mac_only_pm = int(pm_status.get("MACRON ONLY - NO PAYMETER", 0))
    pm_match_rate = _pct(pm_matched, pay_count)
    mac_pm_match_rate = _pct(pm_matched, mac_count)

    add("03 Paymeter ↔ Macron Merge", "Exact Paymeter-Macron matches", pm_matched, "count", "GOOD" if pm_only == 0 and mac_only_pm == 0 else "ATTENTION", "Transactions linked by exact reference.", "Paymeter Reference = Macron REFERENCE ID.")
    add("03 Paymeter ↔ Macron Merge", "Paymeter matched to Macron rate", pm_match_rate, "percent", "GOOD" if pm_match_rate >= 99.99 else "ATTENTION", "Share of Paymeter transactions with a Macron counterpart.", "Exact matches / cleaned Paymeter rows × 100.")
    add("03 Paymeter ↔ Macron Merge", "Macron matched to Paymeter rate", mac_pm_match_rate, "percent", "GOOD" if mac_pm_match_rate >= 99.99 else "ATTENTION", "Share of Macron vends with a Paymeter counterpart.", "Exact matches / Macron rows × 100.")
    add("03 Paymeter ↔ Macron Merge", "Paymeter transactions with no Macron", pm_only, "count", "GOOD" if pm_only == 0 else "ATTENTION", "Paymeter-side transactions not found in Macron.", "PMMAC_Match_Status = PAYMETER ONLY - NO MACRON.", "Investigate whether token/vend was delayed, failed or outside the Macron report period.")
    add("03 Paymeter ↔ Macron Merge", "Macron transactions with no Paymeter", mac_only_pm, "count", "GOOD" if mac_only_pm == 0 else "ATTENTION", "Macron-side transactions not found in Paymeter.", "PMMAC_Match_Status = MACRON ONLY - NO PAYMETER.", "Review for missing upstream transaction, report-period difference or reference-quality issue.")

    pm_only_rows = pm_mac_merge[pm_mac_merge["PMMAC_Match_Status"] == "PAYMETER ONLY - NO MACRON"]
    mac_only_rows = pm_mac_merge[pm_mac_merge["PMMAC_Match_Status"] == "MACRON ONLY - NO PAYMETER"]
    add("03 Paymeter ↔ Macron Merge", "Paymeter token/net value with no Macron counterpart", _num_sum(pm_only_rows.get("PM_Transaction Amount", pd.Series(dtype=float))), "amount", "GOOD" if pm_only == 0 else "ATTENTION", "Token/net value recorded by Paymeter without a Macron match.", "Sum PM_Transaction Amount where Paymeter has no Macron.")
    add("03 Paymeter ↔ Macron Merge", "Macron token value with no Paymeter counterpart", _num_sum(mac_only_rows.get("MAC_AMOUNT", pd.Series(dtype=float))), "amount", "GOOD" if mac_only_pm == 0 else "ATTENTION", "Macron token value without a Paymeter match.", "Sum MAC_AMOUNT where Macron has no Paymeter.")

    # ------------------------------------------------------------
    # SECOND MERGE: VPS ↔ BANK
    # ------------------------------------------------------------
    vb_status = vps_bank_merge["VPSBANK_Match_Status"].value_counts()
    vb_matched = int(vb_status.get("MATCHED", 0))
    vps_only_bank = int(vb_status.get("VPS ONLY - NO BANK", 0))
    bank_only_vps = int(vb_status.get("BANK ONLY - NO VPS", 0))
    exact_bank = int((vps_bank_merge["VPSBANK_Match_Method"] == "VPS session_id found in Bank Narration").sum())
    fallback_bank = int((vps_bank_merge["VPSBANK_Match_Method"] == "Unique VPS settled amount + date").sum())

    add("04 VPS ↔ Bank Merge", "VPS-Bank matched transactions", vb_matched, "count", "GOOD" if vps_only_bank == 0 and bank_only_vps == 0 else "ATTENTION", "Transactions linked between VPS and Bank.", "Exact narration/session match plus controlled unique amount/date fallback.")
    add("04 VPS ↔ Bank Merge", "VPS matched to Bank rate", _pct(vb_matched, vps_count), "percent", "GOOD" if vps_only_bank == 0 else "ATTENTION", "Share of VPS rows with a Bank credit counterpart.", "VPS-Bank matched / VPS rows × 100.")
    add("04 VPS ↔ Bank Merge", "Bank credits matched to VPS rate", _pct(vb_matched, bank_count), "percent", "GOOD" if bank_only_vps == 0 else "ATTENTION", "Share of Bank credits linked to VPS.", "VPS-Bank matched / Bank credit rows × 100.")
    add("04 VPS ↔ Bank Merge", "Exact Bank matches using VPS session ID", exact_bank, "count", "INFO", "Strongest Bank↔VPS match method.", "VPS session_id found in Bank Narration.")
    add("04 VPS ↔ Bank Merge", "Fallback Bank matches", fallback_bank, "count", "INFO", "Bank↔VPS matches made only when settled amount/date candidate was unique.", "Unique VPS settled amount + date.")
    add("04 VPS ↔ Bank Merge", "VPS transactions with no Bank match", vps_only_bank, "count", "GOOD" if vps_only_bank == 0 else "ATTENTION", "VPS transactions for which no Bank credit was matched.", "VPSBANK_Match_Status = VPS ONLY - NO BANK.", "Review Bank report coverage, narration/session IDs and settlement timing.")
    add("04 VPS ↔ Bank Merge", "Bank credits with no VPS match", bank_only_vps, "count", "GOOD" if bank_only_vps == 0 else "ATTENTION", "Bank credits not linked to VPS; these are not automatically assumed to be vending payments.", "VPSBANK_Match_Status = BANK ONLY - NO VPS.", "Investigate separately for unrelated receipts or missing VPS records.")
    bank_only_rows = vps_bank_merge[vps_bank_merge["VPSBANK_Match_Status"] == "BANK ONLY - NO VPS"]
    add("04 VPS ↔ Bank Merge", "Bank credit value with no VPS match", _num_sum(bank_only_rows.get("BANK_Credit", pd.Series(dtype=float))), "amount", "GOOD" if bank_only_vps == 0 else "ATTENTION", "Total unmatched Bank-credit value.", "Sum BANK_Credit where Bank has no VPS match.")

    # ------------------------------------------------------------
    # KYC COVERAGE
    # ------------------------------------------------------------
    kyc_total = len(kyc)
    complete_kyc = int((kyc["KYC Status"] == "COMPLETE CORE KYC").sum())
    incomplete_kyc = kyc_total - complete_kyc
    conflict_kyc = int(kyc["KYC Conflict"].map(clean_id).ne("").sum()) if "KYC Conflict" in kyc.columns else 0
    kyc_status_counts = final_recon["FINAL_KYC_Status"].value_counts()
    no_kyc_rows = int(kyc_status_counts.get("NO KYC", 0))
    one_side_kyc = int(kyc_status_counts.get("KYC ON ONE SIDE", 0))
    aligned_kyc = int(kyc_status_counts.get("KYC ALIGNED", 0))
    mismatch_kyc = int(kyc_status_counts.get("KYC MISMATCH", 0))

    add("05 KYC Coverage", "KYC master customers / VPS virtual accounts", kyc_total, "people", "INFO", "One KYC record per unique VPS virtual account.", "KYC base population from VPS virtual_acct_no.")
    add("05 KYC Coverage", "Complete core KYC", complete_kyc, "people", "GOOD" if incomplete_kyc == 0 else "ATTENTION", "KYC records with core customer, Paymeter and Macron identity populated.", "KYC Status = COMPLETE CORE KYC.")
    add("05 KYC Coverage", "Complete core KYC rate", _pct(complete_kyc, kyc_total), "percent", "GOOD" if incomplete_kyc == 0 else "ATTENTION", "Percentage of the KYC master that is core-complete.", "Complete core KYC / KYC master × 100.")
    add("05 KYC Coverage", "Incomplete KYC", incomplete_kyc, "people", "GOOD" if incomplete_kyc == 0 else "ATTENTION", "KYC records still missing one or more core identities.", "KYC Status != COMPLETE CORE KYC.", "Prioritise identity completion for repeated/high-value customers.")
    add("05 KYC Coverage", "KYC records with conflict flags", conflict_kyc, "people", "GOOD" if conflict_kyc == 0 else "ATTENTION", "KYC records where multiple identity values may need review.", "Non-blank KYC Conflict.")
    add("05 KYC Coverage", "Full-reconciliation rows with aligned KYC", aligned_kyc, "count", "INFO", "Rows where both merged sides carry the same KYC ID.", "FINAL_KYC_Status = KYC ALIGNED.")
    add("05 KYC Coverage", "Full-reconciliation rows with KYC on one side", one_side_kyc, "count", "GOOD" if one_side_kyc == 0 else "ATTENTION", "Only one side of the transaction journey carries KYC.", "FINAL_KYC_Status = KYC ON ONE SIDE.")
    add("05 KYC Coverage", "Full-reconciliation rows with no KYC", no_kyc_rows, "count", "GOOD" if no_kyc_rows == 0 else "ATTENTION", "Rows retained in reconciliation but not attributable to formal KYC.", "FINAL_KYC_Status = NO KYC.")
    add("05 KYC Coverage", "Full-reconciliation KYC mismatches", mismatch_kyc, "count", "GOOD" if mismatch_kyc == 0 else "CRITICAL", "Rows where left and right KYC IDs disagree.", "FINAL_KYC_Status = KYC MISMATCH.", "Review identity mapping before relying on customer-level attribution.")

    # ------------------------------------------------------------
    # END-TO-END RECONCILIATION
    # ------------------------------------------------------------
    total_final = len(final_recon)
    final_status = final_recon["FINAL_Reconciliation_Status"].value_counts()
    fully = int(final_status.get("FULLY RECONCILED", 0))
    amount_status = final_recon["FINAL_Amount_Status"].value_counts()
    aligned_amount = int(amount_status.get("AMOUNTS ALIGNED", 0))
    amount_diff = int(amount_status.get("AMOUNT DIFFERENCE", 0))
    not_comparable = int(amount_status.get("NOT COMPARABLE", 0))
    exact_final = int((final_recon["FINAL_Match_Method"] == "VPS settlement_ref = Paymeter RRN").sum())
    fallback_final = int((final_recon["FINAL_Match_Method"] == "KYC + amount + date fallback").sum())

    add("06 End-to-End Reconciliation", "Full reconciliation journey rows", total_final, "count", "INFO", "All transaction journeys after the two merged reports are reconciled.", "Rows in 07 Full Reconciliation.")
    add("06 End-to-End Reconciliation", "Fully reconciled rows", fully, "count", "GOOD" if fully == total_final else "ATTENTION", "Rows with Bank, VPS, Paymeter, Macron and aligned amounts.", "FINAL_Reconciliation_Status = FULLY RECONCILED.")
    add("06 End-to-End Reconciliation", "Fully reconciled rate", _pct(fully, total_final), "percent", "GOOD" if fully == total_final else "ATTENTION", "Share of full journey rows that completely reconcile.", "Fully reconciled rows / final reconciliation rows × 100.")
    add("06 End-to-End Reconciliation", "Exact VPS settlement_ref ↔ Paymeter RRN matches", exact_final, "count", "INFO", "Strongest bridge between the VPS-Bank and Paymeter-Macron merged reports.", "FINAL_Match_Method = VPS settlement_ref = Paymeter RRN.")
    add("06 End-to-End Reconciliation", "KYC + amount + date fallback matches", fallback_final, "count", "INFO", "Controlled fallback matches when exact transactional bridge is unavailable.", "FINAL_Match_Method = KYC + amount + date fallback.")
    add("06 End-to-End Reconciliation", "Rows with amounts aligned", aligned_amount, "count", "INFO", "Rows where all available financial comparisons agree within tolerance.", "FINAL_Amount_Status = AMOUNTS ALIGNED.")
    add("06 End-to-End Reconciliation", "Rows with financial amount difference", amount_diff, "count", "GOOD" if amount_diff == 0 else "CRITICAL", "Rows with at least one financial comparison outside tolerance.", "FINAL_Amount_Status = AMOUNT DIFFERENCE.", "Review value differences before management sign-off.")
    add("06 End-to-End Reconciliation", "Rows not financially comparable", not_comparable, "count", "GOOD" if not_comparable == 0 else "ATTENTION", "Rows missing one or more stages needed for amount comparison.", "FINAL_Amount_Status = NOT COMPARABLE.")

    # Status-level exception counts
    for status_name, count in final_status.items():
        if status_name == "FULLY RECONCILED":
            continue
        add(
            "06 End-to-End Reconciliation",
            f"Status: {status_name}",
            int(count),
            "count",
            "ATTENTION",
            f"Number of final reconciliation rows classified as {status_name}.",
            f"FINAL_Reconciliation_Status = {status_name}.",
        )

    # ------------------------------------------------------------
    # VPS ↔ MACRON TOKEN FULFILMENT / VALUE
    # ------------------------------------------------------------
    vm_status = vps_macron["VPS_vs_Macron_Status"].value_counts()
    value_match = int(vm_status.get("VALUE MATCH", 0))
    value_difference = int(vm_status.get("VALUE DIFFERENCE", 0))
    paid_no_token_count = int(vm_status.get("PAID / NO TOKEN", 0))
    token_no_vps_count = int(vm_status.get("TOKEN / NO VPS PAYMENT", 0))
    both_vps_mac = value_match + value_difference

    paid_no_token = final_recon[(final_recon["VPS Present?"] == "YES") & (final_recon["Macron Present?"] == "NO")]
    token_no_vps = final_recon[(final_recon["Macron Present?"] == "YES") & (final_recon["VPS Present?"] == "NO")]
    value_diff_rows = vps_macron[vps_macron["VPS_vs_Macron_Status"] == "VALUE DIFFERENCE"].copy()
    value_diff_series = pd.to_numeric(value_diff_rows["Difference Expected Macron - Actual Macron"], errors="coerce")

    add("07 VPS ↔ Macron Token Control", "Fixed Macron token charge per VPS payment", MACRON_TOKEN_CHARGE, "amount", "INFO", "Business rule used for every VPS→Macron comparison.", "Expected Macron Token = VPS transaction_amount_minor - ₦100.")
    add("07 VPS ↔ Macron Token Control", "VPS payments with a Macron token present", both_vps_mac, "count", "GOOD" if paid_no_token_count == 0 else "ATTENTION", "VPS payments that have a Macron counterpart, regardless of value accuracy.", "VALUE MATCH + VALUE DIFFERENCE.")
    add("07 VPS ↔ Macron Token Control", "Macron token presence rate for VPS payments", _pct(both_vps_mac, vps_count), "percent", "GOOD" if paid_no_token_count == 0 else "CRITICAL", "Share of VPS payments for which a Macron token/vend was found.", "VPS rows with Macron present / VPS transactions × 100.", "Paid/no-token cases require operational follow-up.")
    add("07 VPS ↔ Macron Token Control", "VPS-Macron exact value matches", value_match, "count", "GOOD" if value_difference == 0 else "ATTENTION", "VPS payments where actual Macron amount equals VPS payment less ₦100.", "Difference Expected Macron - Actual Macron within tolerance.")
    add("07 VPS ↔ Macron Token Control", "Value accuracy rate where both VPS and Macron exist", _pct(value_match, both_vps_mac), "percent", "GOOD" if value_difference == 0 else "ATTENTION", "Accuracy of token value among transactions where both VPS and Macron records exist.", "VALUE MATCH / (VALUE MATCH + VALUE DIFFERENCE) × 100.")
    add("07 VPS ↔ Macron Token Control", "VPS-Macron value differences", value_difference, "count", "GOOD" if value_difference == 0 else "CRITICAL", "Transactions where Macron value differs from VPS paid less ₦100.", "VPS_vs_Macron_Status = VALUE DIFFERENCE.", "Investigate pricing/token issuance discrepancy.")
    add("07 VPS ↔ Macron Token Control", "Absolute value variance on VPS-Macron differences", value_diff_series.abs().sum(), "amount", "GOOD" if value_difference == 0 else "CRITICAL", "Gross financial magnitude of VPS↔Macron value differences without netting positive and negative differences.", "Sum of absolute Difference Expected Macron - Actual Macron on VALUE DIFFERENCE rows.")
    add("07 VPS ↔ Macron Token Control", "Net VPS-Macron value variance", value_diff_series.sum(), "amount", "GOOD" if value_difference == 0 else "ATTENTION", "Net difference across VPS↔Macron value-difference rows.", "Sum Difference Expected Macron - Actual Macron on VALUE DIFFERENCE rows.")
    add("07 VPS ↔ Macron Token Control", "VPS payments with no Macron token", paid_no_token_count, "count", "GOOD" if paid_no_token_count == 0 else "CRITICAL", "Customer payments/VPS receipts for which no Macron token was found.", "VPS Present? = YES and Macron Present? = NO.", "Prioritise customer impact and outstanding token fulfilment.")
    add("07 VPS ↔ Macron Token Control", "Customers with VPS payment but no token", int(paid_no_token["FINAL_KYC_ID"].replace("", np.nan).nunique()), "people", "GOOD" if paid_no_token_count == 0 else "CRITICAL", "Distinct mapped KYC customers affected by paid/no-token exceptions.", "Distinct non-blank FINAL_KYC_ID among paid/no-token rows.")
    add("07 VPS ↔ Macron Token Control", "Customer amount paid on VPS with no Macron token", _num_sum(paid_no_token["VPS Customer Amount Paid"]), "amount", "GOOD" if paid_no_token_count == 0 else "CRITICAL", "Gross customer money represented by VPS paid/no-token rows.", "Sum VPS Customer Amount Paid where Macron is absent.")
    add("07 VPS ↔ Macron Token Control", "Expected token value not issued", _num_sum(paid_no_token["Expected Macron Token (VPS Paid - 100)"]), "amount", "GOOD" if paid_no_token_count == 0 else "CRITICAL", "Token value that should have been issued after deducting ₦100 per VPS transaction.", "Sum Expected Macron Token on paid/no-token rows.")
    add("07 VPS ↔ Macron Token Control", "Macron tokens with no VPS payment", token_no_vps_count, "count", "GOOD" if token_no_vps_count == 0 else "CRITICAL", "Macron vends for which no VPS payment exists in the reconciled population.", "Macron Present? = YES and VPS Present? = NO.", "Investigate potential unsupported vending, missing upstream payment or report-period timing.")
    add("07 VPS ↔ Macron Token Control", "Customers with Macron token but no VPS payment", int(token_no_vps["FINAL_KYC_ID"].replace("", np.nan).nunique()), "people", "GOOD" if token_no_vps_count == 0 else "CRITICAL", "Distinct mapped KYC customers in token/no-VPS population.", "Distinct non-blank FINAL_KYC_ID among token/no-VPS rows.")
    add("07 VPS ↔ Macron Token Control", "Macron token value with no VPS payment", _num_sum(token_no_vps["Macron Token Amount"]), "amount", "GOOD" if token_no_vps_count == 0 else "CRITICAL", "Actual Macron token value with no linked VPS payment.", "Sum Macron Token Amount where VPS is absent.")

    # ------------------------------------------------------------
    # CUSTOMER SUMMARY / SOURCE CONTROL
    # ------------------------------------------------------------
    unlinked_customers = int((customer_summary["Customer Classification"] != "KYC CUSTOMER").sum()) if "Customer Classification" in customer_summary.columns else 0
    tie_pass = int((summary_control["Status"] == "MATCH").sum())
    tie_fail = int((summary_control["Status"] != "MATCH").sum())
    coverage_complete = int((source_coverage["Status"] == "COMPLETE").sum())
    coverage_review = int((source_coverage["Status"] != "COMPLETE").sum())
    missing_source_rows = int(pd.to_numeric(source_coverage["Missing Rows"], errors="coerce").fillna(0).sum())
    duplicate_source_uses = int(pd.to_numeric(source_coverage["Duplicate Uses"], errors="coerce").fillna(0).sum())

    add("08 Customer Summary & Control", "Customer Summary rows", len(customer_summary), "count", "INFO", "Formal KYC plus retained unlinked source identities.", "Rows in 09 Customer Summary.")
    add("08 Customer Summary & Control", "Unlinked / non-KYC Customer Summary identities", unlinked_customers, "count", "GOOD" if unlinked_customers == 0 else "ATTENTION", "Synthetic identities retained so no source value disappears from the Customer Summary.", "Customer Classification != KYC CUSTOMER.")
    add("08 Customer Summary & Control", "Customer Summary tie-out controls passed", tie_pass, "count", "GOOD" if tie_fail == 0 else "CRITICAL", "Source totals that exactly tie to Customer Summary.", "10 Summary Control Status = MATCH.")
    add("08 Customer Summary & Control", "Customer Summary tie-out controls failed", tie_fail, "count", "GOOD" if tie_fail == 0 else "CRITICAL", "Any failure means Customer Summary does not fully account for a source total.", "10 Summary Control Status != MATCH.", "Workbook should not be signed off if any control fails.")
    add("08 Customer Summary & Control", "Source coverage controls complete", coverage_complete, "count", "GOOD" if coverage_review == 0 else "CRITICAL", "Source populations represented exactly once in final reconciliation.", "15 Source Coverage Status = COMPLETE.")
    add("08 Customer Summary & Control", "Source coverage controls requiring review", coverage_review, "count", "GOOD" if coverage_review == 0 else "CRITICAL", "Number of sources with missing or duplicate row usage.", "15 Source Coverage Status != COMPLETE.")
    add("08 Customer Summary & Control", "Total missing source rows", missing_source_rows, "count", "GOOD" if missing_source_rows == 0 else "CRITICAL", "Source rows not represented in final reconciliation.", "Sum Missing Rows in Source Coverage.")
    add("08 Customer Summary & Control", "Total duplicate source uses", duplicate_source_uses, "count", "GOOD" if duplicate_source_uses == 0 else "CRITICAL", "Source rows represented more than once in final reconciliation.", "Sum Duplicate Uses in Source Coverage.")

    # ------------------------------------------------------------
    # PAYMETER / SOURCE DATA QUALITY
    # ------------------------------------------------------------
    repaired_rows = int(pay.attrs.get("repaired_rows", 0))
    spill_cells = int(pay.attrs.get("spill_cells_deleted", 0))
    ignored_headers = int(pay.attrs.get("ignored_blank_header_columns", 0))
    pm_review = int(pay.get("PM_Cleaning_Status", pd.Series(dtype=object)).astype(str).str.contains("REVIEW", case=False, na=False).sum())

    quality_metrics = [
        ("Paymeter rows cleaned", len(pay), "count", "INFO", "All downstream Paymeter reconciliation uses this cleaned population."),
        ("Paymeter rows with Address spill repaired", repaired_rows, "count", "GOOD" if repaired_rows == 0 else "INFO", "Rows where extra Address fragments were deleted and later fields shifted back."),
        ("Paymeter Address spill cells deleted", spill_cells, "count", "INFO", "Total extra Address cells removed."),
        ("Blank Paymeter header columns ignored", ignored_headers, "count", "INFO", "Trailing blank export columns ignored after Status Checked."),
        ("Paymeter rows requiring cleaning review", pm_review, "count", "GOOD" if pm_review == 0 else "ATTENTION", "Rows not confidently restored by structural cleaning."),
        ("Paymeter rows missing Account Number", int(pay["Account Number"].map(clean_id).eq("").sum()), "count", "GOOD", "Missing Paymeter account identity."),
        ("Paymeter rows missing Meter Number", int(pay["Meter Number"].map(clean_id).eq("").sum()), "count", "GOOD", "Missing Paymeter meter identity."),
        ("Paymeter rows missing RRN", int(pay["RRN"].map(clean_id).eq("").sum()), "count", "GOOD", "Missing primary VPS↔Paymeter reference."),
        ("Paymeter rows missing Reference", int(pay["Reference"].map(clean_id).eq("").sum()), "count", "GOOD", "Missing primary Paymeter↔Macron reference."),
        ("VPS rows missing virtual account", int(vps["virtual_acct_no"].map(clean_id).eq("").sum()), "count", "GOOD", "Missing KYC base key."),
        ("VPS rows missing settlement_ref", int(vps["settlement_ref"].map(clean_id).eq("").sum()), "count", "GOOD", "Missing primary VPS↔Paymeter key."),
        ("VPS rows missing session_id", int(vps["session_id"].map(clean_id).eq("").sum()), "count", "GOOD", "Missing primary Bank↔VPS key."),
        ("Macron rows missing Account Number", int(mac["ACCOUNT NUMBER"].map(clean_id).eq("").sum()), "count", "GOOD", "Missing Macron account identity."),
        ("Macron rows missing Meter Number", int(mac["METER NUMBER"].map(clean_id).eq("").sum()), "count", "GOOD", "Missing Macron meter identity."),
        ("Macron rows missing Reference ID", int(mac["REFERENCE ID"].map(clean_id).eq("").sum()), "count", "GOOD", "Missing primary Paymeter↔Macron key."),
        ("Bank credit rows missing narration", int(bank.loc[bank_credit_mask, "Narration"].map(clean_id).eq("").sum()), "count", "GOOD", "Missing narration can prevent exact VPS session matching."),
    ]
    for parameter, value, unit, default_status, interpretation in quality_metrics:
        status = default_status
        if parameter.startswith(("Paymeter rows missing", "VPS rows missing", "Macron rows missing", "Bank credit rows missing")):
            status = "GOOD" if int(value) == 0 else "ATTENTION"
        add("09 Data Quality", parameter, value, unit, status, interpretation, parameter)

    # ------------------------------------------------------------
    # OVERALL CONTROL RESULT
    # ------------------------------------------------------------
    critical_count = sum(1 for r in rows if r["Status"] == "CRITICAL")
    attention_count = sum(1 for r in rows if r["Status"] == "ATTENTION")
    add(
        "10 Overall Control",
        "Critical management indicators",
        critical_count,
        "count",
        "GOOD" if critical_count == 0 else "CRITICAL",
        "Number of management KPIs currently marked CRITICAL.",
        "Count of management report rows with Status = CRITICAL.",
        "Review all CRITICAL items before final management sign-off.",
    )
    add(
        "10 Overall Control",
        "Attention indicators",
        attention_count,
        "count",
        "GOOD" if attention_count == 0 else "ATTENTION",
        "Number of management KPIs marked ATTENTION.",
        "Count of management report rows with Status = ATTENTION.",
    )
    add(
        "10 Overall Control",
        "Management report conclusion",
        "EXCEPTIONS REQUIRE REVIEW" if critical_count or attention_count else "ALL CONTROLS CLEAR",
        "status",
        "CRITICAL" if critical_count else ("ATTENTION" if attention_count else "GOOD"),
        "High-level conclusion based on reconciliation, KYC, exception, tie-out and data-quality controls.",
        "Derived from management KPI statuses.",
    )

    return pd.DataFrame(rows)


def build_management_breakdown(final_recon, pm_mac_merge, vps_bank_merge, vps_macron):
    """Structured count/value breakdown for management drill-down."""
    rows = []

    def add_group(dimension, series, df, vps_col=None, expected_col=None, mac_col=None, bank_col=None):
        total = len(df)
        for category, idx in series.groupby(series).groups.items():
            g = df.loc[idx]
            rows.append({
                "Dimension": dimension,
                "Category": clean_id(category) or "BLANK",
                "Count": len(g),
                "Rate %": _pct(len(g), total),
                "VPS Customer Amount Paid": _num_sum(g[vps_col]) if vps_col and vps_col in g.columns else np.nan,
                "Expected Macron Token": _num_sum(g[expected_col]) if expected_col and expected_col in g.columns else np.nan,
                "Macron Token Amount": _num_sum(g[mac_col]) if mac_col and mac_col in g.columns else np.nan,
                "Bank Credit": _num_sum(g[bank_col]) if bank_col and bank_col in g.columns else np.nan,
            })

    add_group(
        "Final Reconciliation Status",
        final_recon["FINAL_Reconciliation_Status"],
        final_recon,
        "VPS Customer Amount Paid",
        "Expected Macron Token (VPS Paid - 100)",
        "Macron Token Amount",
        "Bank Credit",
    )
    add_group(
        "Final Amount Status",
        final_recon["FINAL_Amount_Status"],
        final_recon,
        "VPS Customer Amount Paid",
        "Expected Macron Token (VPS Paid - 100)",
        "Macron Token Amount",
        "Bank Credit",
    )
    add_group(
        "Final KYC Status",
        final_recon["FINAL_KYC_Status"],
        final_recon,
        "VPS Customer Amount Paid",
        "Expected Macron Token (VPS Paid - 100)",
        "Macron Token Amount",
        "Bank Credit",
    )
    add_group("Paymeter-Macron Merge Status", pm_mac_merge["PMMAC_Match_Status"], pm_mac_merge)
    add_group("VPS-Bank Merge Status", vps_bank_merge["VPSBANK_Match_Status"], vps_bank_merge)
    add_group(
        "VPS-Macron Status",
        vps_macron["VPS_vs_Macron_Status"],
        vps_macron,
        "VPS Customer Amount Paid",
        "Expected Macron Token (VPS Paid - 100)",
        "Macron Token Amount",
    )
    return pd.DataFrame(rows)


def build_management_exception_register(final_recon):
    """Management-focused register of all non-fully-reconciled transaction journeys."""
    exceptions = final_recon[final_recon["FINAL_Reconciliation_Status"] != "FULLY RECONCILED"].copy()
    if exceptions.empty:
        return pd.DataFrame(columns=[
            "Priority", "Exception Type", "FINAL_Row_ID", "FINAL_KYC_ID",
            "FINAL_Customer_Name", "VPS Virtual Account", "Paymeter Account",
            "Macron Account", "VPS Customer Amount Paid", "Expected Macron Token",
            "Macron Token Amount", "Financial Exposure / Variance", "Bank Credit",
            "VPS Settlement Ref", "Paymeter RRN", "Paymeter Reference",
            "Macron Reference ID", "KYC Status", "Amount Status", "Match Method",
        ])

    def priority(r):
        s = clean_id(r.get("FINAL_Reconciliation_Status"))
        if s in {
            "CUSTOMER PAID / VPS RECEIVED - NO MACRON TOKEN",
            "MACRON TOKEN - NO VPS PAYMENT",
            "VPS / MACRON VALUE DIFFERENCE",
            "KYC MISMATCH",
        }:
            return "CRITICAL"
        return "ATTENTION"

    def exposure(r):
        s = clean_id(r.get("FINAL_Reconciliation_Status"))
        if s == "CUSTOMER PAID / VPS RECEIVED - NO MACRON TOKEN":
            return abs(float(r.get("Expected Macron Token (VPS Paid - 100)"))) if pd.notna(r.get("Expected Macron Token (VPS Paid - 100)")) else 0.0
        if s == "MACRON TOKEN - NO VPS PAYMENT":
            return abs(float(r.get("Macron Token Amount"))) if pd.notna(r.get("Macron Token Amount")) else 0.0
        if pd.notna(r.get("Difference Expected Macron - Actual Macron")):
            return abs(float(r.get("Difference Expected Macron - Actual Macron")))
        if pd.notna(r.get("Difference Bank Credit - VPS Settled")):
            return abs(float(r.get("Difference Bank Credit - VPS Settled")))
        if pd.notna(r.get("Difference VPS Paid - Paymeter Input")):
            return abs(float(r.get("Difference VPS Paid - Paymeter Input")))
        return 0.0

    out = pd.DataFrame({
        "Priority": exceptions.apply(priority, axis=1),
        "Exception Type": exceptions["FINAL_Reconciliation_Status"],
        "FINAL_Row_ID": exceptions["FINAL_Row_ID"],
        "FINAL_KYC_ID": exceptions["FINAL_KYC_ID"],
        "FINAL_Customer_Name": exceptions["FINAL_Customer_Name"],
        "VPS Virtual Account": exceptions.get("LEFT_VPS_virtual_acct_no", ""),
        "Paymeter Account": exceptions.get("RIGHT_PM_Account Number", ""),
        "Macron Account": exceptions.get("RIGHT_MAC_ACCOUNT NUMBER", ""),
        "VPS Customer Amount Paid": exceptions["VPS Customer Amount Paid"],
        "Expected Macron Token": exceptions["Expected Macron Token (VPS Paid - 100)"],
        "Macron Token Amount": exceptions["Macron Token Amount"],
        "Financial Exposure / Variance": exceptions.apply(exposure, axis=1),
        "Bank Credit": exceptions["Bank Credit"],
        "VPS Settlement Ref": exceptions.get("LEFT_VPS_settlement_ref", ""),
        "Paymeter RRN": exceptions.get("RIGHT_PM_RRN", ""),
        "Paymeter Reference": exceptions.get("RIGHT_PM_Reference", ""),
        "Macron Reference ID": exceptions.get("RIGHT_MAC_REFERENCE ID", ""),
        "KYC Status": exceptions["FINAL_KYC_Status"],
        "Amount Status": exceptions["FINAL_Amount_Status"],
        "Match Method": exceptions["FINAL_Match_Method"],
    })
    order = pd.Categorical(out["Priority"], categories=["CRITICAL", "ATTENTION"], ordered=True)
    out = out.assign(_priority_order=order).sort_values(
        ["_priority_order", "Financial Exposure / Variance"],
        ascending=[True, False],
    ).drop(columns=["_priority_order"]).reset_index(drop=True)
    return out


def build_management_customer_risk(customer_summary):
    """Customer-level exception/risk view for management follow-up."""
    c = customer_summary.copy()
    c["Absolute Expected-vs-Macron Difference"] = pd.to_numeric(
        c["Difference Expected Macron - Actual Macron"], errors="coerce"
    ).abs()
    mask = (
        (pd.to_numeric(c["VPS Paid / No Token Count"], errors="coerce").fillna(0) > 0)
        | (pd.to_numeric(c["Token / No VPS Payment Count"], errors="coerce").fillna(0) > 0)
        | (c["Absolute Expected-vs-Macron Difference"] > AMOUNT_TOLERANCE)
        | (c["Customer Classification"] != "KYC CUSTOMER")
    )
    out = c[mask].copy()
    out["Management Issue"] = out.apply(
        lambda r: " | ".join([
            x for x in [
                "PAID / NO TOKEN" if float(pd.to_numeric(r.get("VPS Paid / No Token Count"), errors="coerce") or 0) > 0 else "",
                "TOKEN / NO VPS" if float(pd.to_numeric(r.get("Token / No VPS Payment Count"), errors="coerce") or 0) > 0 else "",
                "VALUE DIFFERENCE" if float(pd.to_numeric(r.get("Absolute Expected-vs-Macron Difference"), errors="coerce") or 0) > AMOUNT_TOLERANCE else "",
                "UNLINKED IDENTITY" if clean_id(r.get("Customer Classification")) != "KYC CUSTOMER" else "",
            ] if x
        ]),
        axis=1,
    )
    out = out.sort_values("Absolute Expected-vs-Macron Difference", ascending=False)
    important = [
        "Management Issue", "Summary Customer Key", "Customer Classification", "KYC_ID",
        "Customer Name", "VPS Virtual Account Number", "Paymeter Account Number(s)",
        "Macron Account Number(s)", "Total VPS Payment Count", "Total Macron Vend Count",
        "VPS Paid / No Token Count", "Token / No VPS Payment Count",
        "Total VPS Customer Amount Paid", "Expected Total Macron Token",
        "Total Macron Token Amount", "Difference Expected Macron - Actual Macron",
        "Absolute Expected-vs-Macron Difference", "KYC Status",
    ]
    return out[[c for c in important if c in out.columns]].reset_index(drop=True)

def build_source_coverage(final_recon, bank, vps, pay, mac):
    checks = [
        ("Bank Credit", int(bank["Credit"].notna().sum()), "LEFT_BANK_BANK_Row"),
        ("VPS", len(vps), "LEFT_VPS_VPS_Row"),
        ("Paymeter", len(pay), "RIGHT_PM_PM_Row"),
        ("Macron", len(mac), "RIGHT_MAC_MAC_Row"),
    ]
    rows = []
    for source, expected, col in checks:
        vals = pd.to_numeric(final_recon.get(col, pd.Series(dtype=float)), errors="coerce").dropna().astype(int)
        represented = len(vals)
        unique = vals.nunique()
        rows.append({
            "Source": source,
            "Expected Rows": expected,
            "Rows Represented": represented,
            "Unique Rows Represented": unique,
            "Duplicate Uses": represented - unique,
            "Missing Rows": expected - unique,
            "Status": "COMPLETE" if expected == unique and represented == unique else "REVIEW",
        })
    return pd.DataFrame(rows)


# ================================================================
# COLUMN GUIDE
# ================================================================

def build_column_guide(sheet_pairs):
    purpose = {
        "01 Cleaned Paymeter": "Paymeter after Address spill correction. This is the only Paymeter dataset used downstream.",
        "02 PM Macron Merge": "First merge: Paymeter with Macron using Paymeter Reference = Macron REFERENCE ID.",
        "03 VPS Bank Merge": "Second merge: VPS with Bank credits using session_id in narration, with unique date/settled-amount fallback.",
        "04 KYC": "KYC master starting from every VPS virtual account number and attaching Paymeter/Macron identity learned from the first merge.",
        "05 PM Macron + KYC": "Paymeter-Macron merge after KYC has been attached.",
        "06 VPS Bank + KYC": "VPS-Bank merge after KYC has been attached.",
        "07 Full Reconciliation": "Reconciliation of the two KYC-enriched merged reports.",
        "08 VPS vs Macron": "Focused comparison of VPS customer payment against Macron token using KYC. Expected token = VPS transaction_amount_minor - 100.",
        "09 Customer Summary": "Complete customer/source summary. Unlinked source identities are retained so source totals do not disappear.",
        "10 Summary Control": "Hard tie-out proving Customer Summary totals equal respective source totals.",
        "11 Management Report": "Comprehensive management KPI report covering periods, source totals, merge rates, KYC, reconciliation, token controls, exceptions, tie-outs and data quality.",
        "11A Mgmt Breakdown": "Management drill-down by final reconciliation, amount, KYC and stage-level match statuses.",
        "11B Mgmt Exceptions": "Transaction-level register of all non-fully-reconciled journeys, prioritised by financial exposure.",
        "11C Customer Risk": "Customer-level exception view showing paid/no-token, token/no-VPS, value differences and unlinked identities.",
        "12 Paid No Token": "VPS payment rows with no Macron token.",
        "13 Token No VPS": "Macron token rows with no VPS payment.",
        "14 Missing KYC": "Full-reconciliation rows where no KYC could be attached.",
        "15 Source Coverage": "Checks that every Bank credit, VPS, cleaned Paymeter and Macron row is represented exactly once.",
    }
    rows = []
    for sheet, df in sheet_pairs:
        for col in df.columns:
            calc = "Copied from source / merged source row."
            meaning = col
            if col == "Expected Macron Token (VPS Paid - 100)":
                calc = "VPS transaction_amount_minor - 100. This is the authoritative expected Macron token value."
                meaning = "Token value Macron should issue for the VPS payment."
            elif col == "Difference Expected Macron - Actual Macron":
                calc = "(VPS transaction_amount_minor - 100) - Macron AMOUNT."
                meaning = "Primary VPS-to-Macron value difference."
            elif col == "PMMAC_Match_Status":
                calc = "MATCHED when Paymeter Reference = Macron REFERENCE ID; otherwise source-only status."
            elif col == "VPSBANK_Match_Status":
                calc = "MATCHED when VPS session_id is found in Bank Narration or unique settled amount/date fallback succeeds."
            elif col == "KYC_ID":
                calc = "Base KYC ID is VPS virtual_acct_no. On merged rows, KYC is attached through RRN/account/Macron identity mappings."
            elif col == "FINAL_Match_Method":
                calc = "Primary: VPS settlement_ref = Paymeter RRN. Fallback: same KYC + matching gross/token amount + date."
            elif col == "FINAL_Reconciliation_Status":
                calc = "Derived from Bank/VPS/Paymeter/Macron presence, KYC alignment and amount checks."
            elif col.startswith("Difference"):
                calc = "Arithmetic difference between the two values named in the header. Zero means exact alignment."
            elif col.startswith("Total "):
                calc = "Sum/count of the named value for the customer/source identity group."
            elif col.startswith("PM_"):
                calc = "Paymeter field after Paymeter cleaning, prefixed PM_."
            elif col.startswith("MAC_"):
                calc = "Macron source field, prefixed MAC_."
            elif col.startswith("VPS_"):
                calc = "VPS source field, prefixed VPS_."
            elif col.startswith("BANK_"):
                calc = "Bank source field, prefixed BANK_."
            elif sheet == "11 Management Report" and col == "Value":
                calc = "Calculated KPI value. See the Calculation / Source column on the same row for the exact formula/source."
                meaning = "Management KPI result."
            elif sheet == "11 Management Report" and col == "Status":
                calc = "GOOD when control/exception is clear; ATTENTION when follow-up is needed; CRITICAL for material reconciliation/customer-impact exceptions; INFO for descriptive KPIs."
                meaning = "Management attention classification."
            elif sheet == "11B Mgmt Exceptions" and col == "Financial Exposure / Variance":
                calc = "Paid/no-token: expected token value; token/no-VPS: actual Macron token value; value difference: absolute expected-vs-actual token difference; otherwise available financial difference."
                meaning = "Financial magnitude used to prioritise the exception."
            rows.append({"Sheet": sheet, "Sheet Purpose": purpose.get(sheet, "Reconciliation output."), "Column Header": col, "Meaning": meaning, "How Calculated / Arrived At": calc})
    return pd.DataFrame(rows)


# ================================================================
# EXCEL OUTPUT
# ================================================================

def _excel_safe_value(value):
    """Convert pandas/numpy values to XlsxWriter-friendly scalars."""
    if value is None:
        return None
    try:
        if pd.isna(value):
            return None
    except Exception:
        pass
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if isinstance(value, np.generic):
        return value.item()
    return value


def _workbook_sheet_pairs(result, management_only=False):
    if management_only:
        return [
            ("11 Management Report", result["management_report"]),
            ("11A Mgmt Breakdown", result["management_breakdown"]),
            ("11B Mgmt Exceptions", result["management_exceptions"]),
            ("11C Customer Risk", result["management_customer_risk"]),
            ("10 Summary Control", result["summary_control"]),
            ("15 Source Coverage", result["source_coverage"]),
        ]

    return [
        ("01 Cleaned Paymeter", result["pay_cleaned"]),
        ("02 PM Macron Merge", result["pm_mac_merge"]),
        ("03 VPS Bank Merge", result["vps_bank_merge"]),
        ("04 KYC", result["kyc"]),
        ("05 PM Macron + KYC", result["pm_mac_kyc"]),
        ("06 VPS Bank + KYC", result["vps_bank_kyc"]),
        ("07 Full Reconciliation", result["final_recon"]),
        ("08 VPS vs Macron", result["vps_macron"]),
        ("09 Customer Summary", result["customer_summary"]),
        ("10 Summary Control", result["summary_control"]),
        ("11 Management Report", result["management_report"]),
        ("11A Mgmt Breakdown", result["management_breakdown"]),
        ("11B Mgmt Exceptions", result["management_exceptions"]),
        ("11C Customer Risk", result["management_customer_risk"]),
        ("12 Paid No Token", result["paid_no_token"]),
        ("13 Token No VPS", result["token_no_vps"]),
        ("14 Missing KYC", result["missing_kyc"]),
        ("15 Source Coverage", result["source_coverage"]),
    ]


def make_excel(result, management_only=False):
    """
    Build the Excel workbook in STREAMING / constant-memory mode.

    V11 used pandas.to_excel() into an in-memory BytesIO object while all
    reconciliation DataFrames were already resident in RAM. On Streamlit
    Community Cloud, large reports can make that memory spike high enough for
    the process to be terminated and the browser to show the generic "Oh no"
    page.

    V12 writes directly to a temporary .xlsx file on disk with XlsxWriter
    constant_memory=True, then reads only the final compressed workbook bytes.
    """
    try:
        import xlsxwriter
    except ImportError as exc:
        raise ImportError(
            "Missing dependency: xlsxwriter. Add xlsxwriter>=3.2 to requirements.txt."
        ) from exc

    sheets = _workbook_sheet_pairs(result, management_only=management_only)
    guide = build_column_guide(sheets)
    sheets.insert(0, ("00 Column Guide", guide))

    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    tmp_path = tmp.name
    tmp.close()

    wb = None
    try:
        wb = xlsxwriter.Workbook(
            tmp_path,
            {
                "constant_memory": True,
                "strings_to_urls": False,
                "in_memory": False,
            },
        )

        fmt_header = wb.add_format({
            "bold": True,
            "font_color": "white",
            "bg_color": "#1F4E78",
            "border": 1,
            "text_wrap": True,
            "valign": "vcenter",
        })
        fmt_money = wb.add_format({
            "num_format": '#,##0.00;[Red](#,##0.00);-'
        })
        fmt_date = wb.add_format({
            "num_format": "dd-mmm-yyyy hh:mm:ss"
        })
        fmt_wrap = wb.add_format({
            "text_wrap": True,
            "valign": "top",
        })
        fmt_good = wb.add_format({
            "bg_color": "#C6EFCE",
            "font_color": "#006100",
        })
        fmt_bad = wb.add_format({
            "bg_color": "#FFC7CE",
            "font_color": "#9C0006",
        })
        fmt_warn = wb.add_format({
            "bg_color": "#FFF2CC",
            "font_color": "#7F6000",
        })
        fmt_info = wb.add_format({
            "bg_color": "#D9EAF7",
            "font_color": "#1F4E78",
        })

        for name, df in sheets:
            sheet_name = name[:31]
            ws = wb.add_worksheet(sheet_name)

            # Do NOT copy the DataFrame. Just choose visible columns.
            cols = [
                c for c in df.columns
                if not c.endswith("_Date")
                and not c.endswith("_DateTime")
            ]

            for j, c in enumerate(cols):
                ws.write(0, j, c, fmt_header)

                lower = c.lower()
                if any(
                    x in lower
                    for x in [
                        "name", "narration", "method", "status", "address",
                        "purpose", "meaning", "calculated", "attention",
                        "interpretation",
                    ]
                ):
                    width, cell_fmt = 30, fmt_wrap
                elif any(
                    x in lower
                    for x in [
                        "amount", "credit", "charge", "difference", "total",
                    ]
                ):
                    width, cell_fmt = 18, fmt_money
                else:
                    width, cell_fmt = min(max(len(c) + 2, 12), 24), None

                ws.set_column(j, j, width, cell_fmt)

            # Stream rows one at a time.
            if cols:
                col_positions = [df.columns.get_loc(c) for c in cols]
                for excel_row, row in enumerate(
                    df.itertuples(index=False, name=None),
                    start=1,
                ):
                    for excel_col, source_pos in enumerate(col_positions):
                        value = _excel_safe_value(row[source_pos])
                        if value is None:
                            continue
                        if isinstance(value, (datetime, date)):
                            ws.write_datetime(
                                excel_row,
                                excel_col,
                                value,
                                fmt_date,
                            )
                        else:
                            ws.write(excel_row, excel_col, value)

                ws.freeze_panes(1, 0)
                ws.autofilter(
                    0,
                    0,
                    max(len(df), 1),
                    len(cols) - 1,
                )

            # Conditional formatting.
            for status_col in [
                "Status",
                "FINAL_Reconciliation_Status",
                "FINAL_Amount_Status",
                "VPS_vs_Macron_Status",
                "PMMAC_Match_Status",
                "VPSBANK_Match_Status",
            ]:
                if status_col not in cols or not len(df):
                    continue
                j = cols.index(status_col)
                ws.conditional_format(
                    1, j, len(df), j,
                    {
                        "type": "text",
                        "criteria": "containing",
                        "value": "MATCH",
                        "format": fmt_good,
                    },
                )
                ws.conditional_format(
                    1, j, len(df), j,
                    {
                        "type": "text",
                        "criteria": "containing",
                        "value": "NO ",
                        "format": fmt_bad,
                    },
                )
                ws.conditional_format(
                    1, j, len(df), j,
                    {
                        "type": "text",
                        "criteria": "containing",
                        "value": "DIFFERENCE",
                        "format": fmt_bad,
                    },
                )

            if name == "11 Management Report" and len(df):
                for c, w in {
                    "Section": 30,
                    "Parameter": 48,
                    "Value": 22,
                    "Unit": 14,
                    "Status": 14,
                    "Interpretation": 58,
                    "Calculation / Source": 58,
                    "Management Attention / Action": 58,
                }.items():
                    if c in cols:
                        j = cols.index(c)
                        ws.set_column(
                            j,
                            j,
                            w,
                            fmt_wrap
                            if c not in {"Value", "Unit", "Status"}
                            else None,
                        )
                if "Status" in cols:
                    sj = cols.index("Status")
                    for value, fmt in [
                        ("GOOD", fmt_good),
                        ("ATTENTION", fmt_warn),
                        ("CRITICAL", fmt_bad),
                        ("INFO", fmt_info),
                    ]:
                        ws.conditional_format(
                            1, sj, len(df), sj,
                            {
                                "type": "text",
                                "criteria": "containing",
                                "value": value,
                                "format": fmt,
                            },
                        )

            if (
                name == "11B Mgmt Exceptions"
                and len(df)
                and "Priority" in cols
            ):
                pj = cols.index("Priority")
                ws.conditional_format(
                    1, pj, len(df), pj,
                    {
                        "type": "text",
                        "criteria": "containing",
                        "value": "CRITICAL",
                        "format": fmt_bad,
                    },
                )
                ws.conditional_format(
                    1, pj, len(df), pj,
                    {
                        "type": "text",
                        "criteria": "containing",
                        "value": "ATTENTION",
                        "format": fmt_warn,
                    },
                )

        wb.close()
        wb = None

        data = Path(tmp_path).read_bytes()
        return data

    finally:
        if wb is not None:
            try:
                wb.close()
            except Exception:
                pass
        try:
            os.remove(tmp_path)
        except OSError:
            pass
        gc.collect()


# ================================================================
# MASTER ENGINE — EXACT USER-REQUESTED ORDER
# ================================================================

def run_reconciliation(bank_file, vps_file, paymeter_file, macron_file, date_tolerance=DATE_TOLERANCE_DAYS, progress=None):
    progress_call(progress, 5, "Cleaning Paymeter before any reconciliation...")
    pay = clean_paymeter(paymeter_file)
    mac = read_macron(macron_file)
    vps = read_vps(vps_file)
    bank = read_bank(bank_file)

    progress_call(progress, 18, "1/6 Merging cleaned Paymeter with Macron...")
    pm_mac = merge_paymeter_macron(pay, mac)

    progress_call(progress, 32, "2/6 Merging VPS with Bank statement credits...")
    vps_bank = merge_vps_bank(vps, bank, date_tolerance)

    progress_call(progress, 45, "3/6 Building KYC from VPS accounts using both merged reports...")
    kyc = build_kyc(vps_bank, pm_mac)

    progress_call(progress, 56, "4/6 Attaching KYC to Paymeter-Macron merge...")
    pm_mac_kyc = attach_kyc_to_pm_mac(pm_mac, kyc, vps_bank)

    progress_call(progress, 64, "5/6 Attaching KYC to VPS-Bank merge...")
    vps_bank_kyc = attach_kyc_to_vps_bank(vps_bank, kyc)

    progress_call(progress, 75, "6/6 Reconciling the two KYC-enriched merged reports...")
    final_recon = reconcile_merged_reports(vps_bank_kyc, pm_mac_kyc, date_tolerance)

    progress_call(progress, 84, "Comparing VPS against Macron using KYC and VPS payment less N100...")
    vps_macron = build_vps_macron_comparison(final_recon)

    progress_call(progress, 90, "Building complete Customer Summary and source tie-out...")
    customer_summary = build_customer_summary(final_recon, kyc)
    summary_control = build_summary_control(customer_summary, bank, vps, pay, mac)
    validate_tieout(summary_control)

    paid_no_token = final_recon[(final_recon["VPS Present?"] == "YES") & (final_recon["Macron Present?"] == "NO")].copy()
    token_no_vps = final_recon[(final_recon["Macron Present?"] == "YES") & (final_recon["VPS Present?"] == "NO")].copy()
    missing_kyc = final_recon[final_recon["FINAL_KYC_ID"].map(clean_id).eq("")].copy()
    source_coverage = build_source_coverage(final_recon, bank, vps, pay, mac)

    progress_call(progress, 95, "Building comprehensive management report and exception analysis...")
    management_report = build_management_report(
        bank=bank,
        vps=vps,
        pay=pay,
        mac=mac,
        pm_mac_merge=pm_mac,
        vps_bank_merge=vps_bank,
        kyc=kyc,
        final_recon=final_recon,
        vps_macron=vps_macron,
        customer_summary=customer_summary,
        summary_control=summary_control,
        source_coverage=source_coverage,
    )
    management_breakdown = build_management_breakdown(
        final_recon, pm_mac, vps_bank, vps_macron
    )
    management_exceptions = build_management_exception_register(final_recon)
    management_customer_risk = build_management_customer_risk(customer_summary)

    progress_call(progress, 100, "Reconciliation and management reporting complete.")
    return {
        "pay_cleaned": pay,
        "pm_mac_merge": pm_mac,
        "vps_bank_merge": vps_bank,
        "kyc": kyc,
        "pm_mac_kyc": pm_mac_kyc,
        "vps_bank_kyc": vps_bank_kyc,
        "final_recon": final_recon,
        "vps_macron": vps_macron,
        "customer_summary": customer_summary,
        "summary_control": summary_control,
        "management_report": management_report,
        "management_breakdown": management_breakdown,
        "management_exceptions": management_exceptions,
        "management_customer_risk": management_customer_risk,
        # Backward-compatible alias used by older UI/code.
        "management": management_report,
        "paid_no_token": paid_no_token,
        "token_no_vps": token_no_vps,
        "missing_kyc": missing_kyc,
        "source_coverage": source_coverage,
    }


# ================================================================
# STREAMLIT
# ================================================================

def main():
    if st is None:
        raise RuntimeError("Install Streamlit using requirements.txt")

    st.set_page_config(
        page_title="Merge-First KYC Reconciliation V12",
        page_icon="🧾",
        layout="wide",
    )
    st.title("🧾 Merge-First KYC Reconciliation V12")
    st.caption(
        "Cloud-safe management reporting: Paymeter+Macron → VPS+Bank → "
        "KYC → reconcile both → VPS vs Macron"
    )

    st.info(
        "V12 keeps the management report but reduces Streamlit Cloud memory "
        "pressure. The full detailed Excel workbook is now optional and is "
        "written in streaming mode instead of being built entirely in RAM."
    )

    with st.sidebar:
        st.header("Upload four reports")
        bank_file = st.file_uploader(
            "Providus Bank Statement",
            type=["xlsx", "xls"],
        )
        vps_file = st.file_uploader(
            "VPS Report",
            type=["xlsx", "xls"],
        )
        pay_file = st.file_uploader(
            "Paymeter Report",
            type=["csv", "xlsx", "xls"],
        )
        mac_file = st.file_uploader(
            "Macron Report",
            type=["xlsx", "xls"],
        )

        st.header("Settings")
        date_tol = st.slider(
            "Fallback date tolerance (days)",
            0,
            7,
            DATE_TOLERANCE_DAYS,
        )
        st.metric("Fixed Macron token charge", "₦100.00")

        build_full_workbook = st.checkbox(
            "Build full detailed Excel workbook",
            value=False,
            help=(
                "Leave this OFF for fastest/lowest-memory cloud processing. "
                "A smaller Management workbook is always produced. Turn it ON "
                "when you need all detailed reconciliation sheets."
            ),
        )

        detail_rows = st.select_slider(
            "Rows to preview in detailed tables",
            options=[100, 250, 500, 1000],
            value=250,
        )

        run = st.button(
            "RUN MERGE-FIRST RECONCILIATION",
            type="primary",
            use_container_width=True,
        )

    if not run:
        st.markdown(
            """
### Processing order
1. Clean Paymeter.
2. Merge Paymeter with Macron.
3. Merge VPS with Bank credits.
4. Build KYC from VPS virtual accounts using the two merged reports.
5. Attach KYC to both merged reports.
6. Reconcile the KYC-enriched merged reports.
7. Compare VPS against Macron using KYC and **Expected Token = VPS transaction_amount_minor − ₦100**.
8. Generate management KPIs, exception analysis, Customer Summary and hard source tie-outs.

### Streamlit Cloud note
For large reports, leave **Build full detailed Excel workbook** OFF on the first run.
The management report will still be generated and downloadable. The detailed
workbook can be produced on a later run when required.
            """
        )
        return

    if not all([bank_file, vps_file, pay_file, mac_file]):
        st.error("Upload all four reports.")
        return

    bar = st.progress(0)
    stage = st.empty()

    def cb(pct, msg):
        bar.progress(min(max(int(pct), 0), 100))
        stage.info(msg)

    try:
        result = run_reconciliation(
            bank_file,
            vps_file,
            pay_file,
            mac_file,
            date_tol,
            cb,
        )

        stage.info(
            "Reconciliation completed. Building low-memory Management workbook..."
        )
        management_excel = make_excel(
            result,
            management_only=True,
        )
        stage.success("Management report is ready.")

        st.subheader("Management Dashboard")
        mg = result["management_report"]

        def mg_value(parameter, default=0):
            hit = mg[mg["Parameter"] == parameter]
            return hit.iloc[0]["Value"] if len(hit) else default

        c1, c2, c3, c4 = st.columns(4)
        c1.metric(
            "Fully reconciled rate",
            f"{float(mg_value('Fully reconciled rate')):,.2f}%",
        )
        c2.metric(
            "Complete KYC rate",
            f"{float(mg_value('Complete core KYC rate')):,.2f}%",
        )
        c3.metric(
            "VPS token presence rate",
            f"{float(mg_value('Macron token presence rate for VPS payments')):,.2f}%",
        )
        c4.metric(
            "Critical indicators",
            f"{int(float(mg_value('Critical management indicators'))):,}",
        )

        st.markdown("### Management Report")
        st.dataframe(
            result["management_report"],
            use_container_width=True,
            hide_index=True,
        )

        st.download_button(
            "⬇️ DOWNLOAD MANAGEMENT REPORT WORKBOOK",
            data=management_excel,
            file_name="Merge_First_KYC_Management_Report_V12.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            type="primary",
            use_container_width=True,
        )

        st.markdown("### Management Drill-down")
        management_view = st.selectbox(
            "Choose management view",
            [
                "Management Breakdown",
                "Management Exceptions",
                "Customer Risk",
                "Summary Control",
                "Source Coverage",
            ],
        )
        management_map = {
            "Management Breakdown": result["management_breakdown"],
            "Management Exceptions": result["management_exceptions"],
            "Customer Risk": result["management_customer_risk"],
            "Summary Control": result["summary_control"],
            "Source Coverage": result["source_coverage"],
        }
        selected_management = management_map[management_view]
        st.caption(
            f"{len(selected_management):,} row(s) in this view."
        )
        st.dataframe(
            selected_management.head(detail_rows),
            use_container_width=True,
            hide_index=True,
        )

        st.markdown("### Detailed Reconciliation Preview")
        detail_view = st.selectbox(
            "Choose one detailed table to preview",
            [
                "Full Reconciliation",
                "KYC",
                "VPS vs Macron",
                "Customer Summary",
                "Paymeter-Macron Merge",
                "VPS-Bank Merge",
                "Paid No Token",
                "Token No VPS",
                "Missing KYC",
            ],
        )
        detail_map = {
            "Full Reconciliation": result["final_recon"],
            "KYC": result["kyc"],
            "VPS vs Macron": result["vps_macron"],
            "Customer Summary": result["customer_summary"],
            "Paymeter-Macron Merge": result["pm_mac_merge"],
            "VPS-Bank Merge": result["vps_bank_merge"],
            "Paid No Token": result["paid_no_token"],
            "Token No VPS": result["token_no_vps"],
            "Missing KYC": result["missing_kyc"],
        }
        selected_detail = detail_map[detail_view]
        st.caption(
            f"{len(selected_detail):,} row(s) total; showing first "
            f"{min(detail_rows, len(selected_detail)):,}."
        )
        st.dataframe(
            selected_detail.head(detail_rows),
            use_container_width=True,
            hide_index=True,
        )

        if build_full_workbook:
            stage.info(
                "Building full detailed workbook in low-memory streaming mode..."
            )
            full_excel = make_excel(
                result,
                management_only=False,
            )
            stage.success("Full detailed workbook is ready.")
            st.download_button(
                "⬇️ DOWNLOAD FULL DETAILED RECONCILIATION WORKBOOK",
                data=full_excel,
                file_name="Merge_First_KYC_Reconciliation_V12.xlsx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                type="secondary",
                use_container_width=True,
            )

        # Encourage prompt release of temporary objects after Streamlit has
        # serialized the current response.
        gc.collect()

    except MemoryError:
        stage.empty()
        st.error(
            "The app ran out of memory while processing the uploaded reports. "
            "Re-run with 'Build full detailed Excel workbook' OFF and preview "
            "fewer rows. The management report uses substantially less memory."
        )
    except Exception as exc:
        stage.empty()
        st.exception(exc)


if __name__ == "__main__":
    main()
