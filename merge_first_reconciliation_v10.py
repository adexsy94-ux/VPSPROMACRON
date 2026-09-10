# merge_first_reconciliation_v10.py
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


def parse_datetime(series, utc=False):
    if utc:
        d = pd.to_datetime(series, errors="coerce", utc=True)
        try:
            return d.dt.tz_convert("Africa/Lagos").dt.tz_localize(None)
        except Exception:
            return d.dt.tz_localize(None)
    return pd.to_datetime(series, errors="coerce", dayfirst=True)


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

    df["PM_DateTime"] = parse_datetime(df["Created At"])
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
# MANAGEMENT / EXCEPTION REPORTS
# ================================================================

def build_management(final_recon, kyc, pay, control):
    paid_no_token = final_recon[(final_recon["VPS Present?"] == "YES") & (final_recon["Macron Present?"] == "NO")]
    token_no_vps = final_recon[(final_recon["Macron Present?"] == "YES") & (final_recon["VPS Present?"] == "NO")]
    missing_kyc = final_recon[final_recon["FINAL_KYC_ID"].map(clean_id).eq("")]
    metrics = [
        ("Full reconciliation rows", len(final_recon), "count"),
        ("Fully reconciled rows", int((final_recon["FINAL_Reconciliation_Status"] == "FULLY RECONCILED").sum()), "count"),
        ("KYC customers", len(kyc), "people"),
        ("Rows with no KYC", len(missing_kyc), "count"),
        ("VPS payments with no Macron token", len(paid_no_token), "count"),
        ("Customers with VPS payment but no Macron token", paid_no_token["FINAL_KYC_ID"].replace("", np.nan).nunique(), "people"),
        ("Customer amount paid on VPS but no token", pd.to_numeric(paid_no_token["VPS Customer Amount Paid"], errors="coerce").sum(), "amount"),
        ("Expected token value not issued (VPS paid - 100)", pd.to_numeric(paid_no_token["Expected Macron Token (VPS Paid - 100)"], errors="coerce").sum(), "amount"),
        ("Macron tokens with no VPS payment", len(token_no_vps), "count"),
        ("Customers with Macron token but no VPS payment", token_no_vps["FINAL_KYC_ID"].replace("", np.nan).nunique(), "people"),
        ("Macron token value with no VPS payment", pd.to_numeric(token_no_vps["Macron Token Amount"], errors="coerce").sum(), "amount"),
        ("Paymeter rows cleaned", len(pay), "count"),
        ("Paymeter rows with Address spill repaired", int(pay.attrs.get("repaired_rows", 0)), "count"),
        ("Paymeter Address spill cells deleted", int(pay.attrs.get("spill_cells_deleted", 0)), "count"),
        ("Customer Summary tie-out controls passed", int((control["Status"] == "MATCH").sum()), "count"),
        ("Customer Summary tie-out controls failed", int((control["Status"] != "MATCH").sum()), "count"),
    ]
    return pd.DataFrame(metrics, columns=["Metric", "Value", "Unit"])


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
        "11 Reconciliation Summary": "Management exception and value summary.",
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
            rows.append({"Sheet": sheet, "Sheet Purpose": purpose.get(sheet, "Reconciliation output."), "Column Header": col, "Meaning": meaning, "How Calculated / Arrived At": calc})
    return pd.DataFrame(rows)


# ================================================================
# EXCEL OUTPUT
# ================================================================

def make_excel(result):
    output = io.BytesIO()
    sheets = [
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
        ("11 Reconciliation Summary", result["management"]),
        ("12 Paid No Token", result["paid_no_token"]),
        ("13 Token No VPS", result["token_no_vps"]),
        ("14 Missing KYC", result["missing_kyc"]),
        ("15 Source Coverage", result["source_coverage"]),
    ]
    guide = build_column_guide(sheets)
    sheets.insert(0, ("00 Column Guide", guide))

    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="dd-mmm-yyyy hh:mm:ss", engine_kwargs={"options": {"strings_to_urls": False}}) as writer:
        wb = writer.book
        fmt_header = wb.add_format({"bold": True, "font_color": "white", "bg_color": "#1F4E78", "border": 1, "text_wrap": True, "valign": "vcenter"})
        fmt_money = wb.add_format({"num_format": '#,##0.00;[Red](#,##0.00);-'})
        fmt_wrap = wb.add_format({"text_wrap": True, "valign": "top"})
        fmt_good = wb.add_format({"bg_color": "#C6EFCE", "font_color": "#006100"})
        fmt_bad = wb.add_format({"bg_color": "#FFC7CE", "font_color": "#9C0006"})

        for name, df in sheets:
            safe = df.drop(columns=[c for c in df.columns if c.endswith("_Date") or c.endswith("_DateTime")], errors="ignore").copy()
            safe.to_excel(writer, sheet_name=name[:31], index=False)
            ws = writer.sheets[name[:31]]
            ws.freeze_panes(1, 0)
            if len(safe.columns):
                ws.autofilter(0, 0, max(len(safe), 1), len(safe.columns) - 1)
            for j, c in enumerate(safe.columns):
                ws.write(0, j, c, fmt_header)
                lower = c.lower()
                width = 16
                if any(x in lower for x in ["name", "narration", "method", "status", "address", "purpose", "meaning", "calculated"]):
                    width, cell_fmt = 30, fmt_wrap
                elif any(x in lower for x in ["amount", "credit", "charge", "difference", "total"]):
                    width, cell_fmt = 18, fmt_money
                else:
                    width, cell_fmt = min(max(len(c) + 2, 12), 24), None
                ws.set_column(j, j, width, cell_fmt)
            for status_col in ["Status", "FINAL_Reconciliation_Status", "FINAL_Amount_Status", "VPS_vs_Macron_Status", "PMMAC_Match_Status", "VPSBANK_Match_Status"]:
                if status_col in safe.columns and len(safe):
                    j = safe.columns.get_loc(status_col)
                    ws.conditional_format(1, j, len(safe), j, {"type": "text", "criteria": "containing", "value": "MATCH", "format": fmt_good})
                    ws.conditional_format(1, j, len(safe), j, {"type": "text", "criteria": "containing", "value": "NO ", "format": fmt_bad})
                    ws.conditional_format(1, j, len(safe), j, {"type": "text", "criteria": "containing", "value": "DIFFERENCE", "format": fmt_bad})
    output.seek(0)
    return output


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
    management = build_management(final_recon, kyc, pay, summary_control)

    progress_call(progress, 100, "Reconciliation calculations complete.")
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
        "management": management,
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
    st.set_page_config(page_title="Merge-First KYC Reconciliation", page_icon="🧾", layout="wide")
    st.title("🧾 Merge-First KYC Reconciliation")
    st.caption("Paymeter+Macron → VPS+Bank → KYC → attach KYC to both merges → reconcile both → VPS vs Macron")
    st.info(
        "This version follows the merge order exactly. Paymeter is cleaned first. "
        "Expected Macron token is always VPS transaction_amount_minor less ₦100."
    )

    with st.sidebar:
        bank_file = st.file_uploader("Providus Bank Statement", type=["xlsx", "xls"])
        vps_file = st.file_uploader("VPS Report", type=["xlsx", "xls"])
        pay_file = st.file_uploader("Paymeter Report", type=["csv", "xlsx", "xls"])
        mac_file = st.file_uploader("Macron Report", type=["xlsx", "xls"])
        date_tol = st.slider("Fallback date tolerance (days)", 0, 7, DATE_TOLERANCE_DAYS)
        st.metric("Fixed Macron token charge", "₦100.00")
        run = st.button("RUN MERGE-FIRST RECONCILIATION", type="primary", use_container_width=True)

    if not run:
        st.markdown("""
### Exact processing order
1. Clean Paymeter.
2. Merge Paymeter with Macron.
3. Merge VPS with Bank credits.
4. Build KYC starting from VPS virtual account numbers and enrich it from the Paymeter-Macron merge.
5. Attach KYC to the Paymeter-Macron merged report.
6. Attach KYC to the VPS-Bank merged report.
7. Reconcile the two KYC-enriched merged reports.
8. Compare VPS against Macron using KYC and **Expected Token = VPS transaction_amount_minor - ₦100**.
""")
        return

    if not all([bank_file, vps_file, pay_file, mac_file]):
        st.error("Upload all four reports.")
        return

    bar = st.progress(0)
    stage = st.empty()
    def cb(pct, msg):
        bar.progress(int(pct))
        stage.info(msg)

    try:
        result = run_reconciliation(bank_file, vps_file, pay_file, mac_file, date_tol, cb)
        stage.success("Reconciliation calculation completed. Building Excel workbook...")
        excel = make_excel(result)
        stage.success("Excel workbook ready.")

        st.subheader("Reconciliation Summary")
        st.dataframe(result["management"], use_container_width=True, hide_index=True)
        tabs = st.tabs(["Full Reconciliation", "KYC", "VPS vs Macron", "Customer Summary", "Summary Control", "Source Coverage"])
        with tabs[0]: st.dataframe(result["final_recon"].head(1500), use_container_width=True, hide_index=True)
        with tabs[1]: st.dataframe(result["kyc"].head(1500), use_container_width=True, hide_index=True)
        with tabs[2]: st.dataframe(result["vps_macron"].head(1500), use_container_width=True, hide_index=True)
        with tabs[3]: st.dataframe(result["customer_summary"].head(1500), use_container_width=True, hide_index=True)
        with tabs[4]: st.dataframe(result["summary_control"], use_container_width=True, hide_index=True)
        with tabs[5]: st.dataframe(result["source_coverage"], use_container_width=True, hide_index=True)

        st.download_button(
            "⬇️ DOWNLOAD RECONCILIATION WORKBOOK",
            data=excel,
            file_name="Merge_First_KYC_Reconciliation_V10.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
        )
    except Exception as exc:
        stage.empty()
        st.exception(exc)


if __name__ == "__main__":
    main()
