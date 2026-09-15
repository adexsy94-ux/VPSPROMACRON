MERGE-FIRST KYC RECONCILIATION V12 — STREAMLIT CLOUD SAFE

WHY V11 COULD SHOW "OH NO"
==========================
Streamlit Community Cloud uses finite RAM. V11 held all reconciliation
DataFrames in memory and immediately built a large multi-sheet Excel
workbook in memory. For larger source reports this can create a large
temporary RAM spike. Streamlit's generic "Oh no" page can occur when the
app process exceeds resource limits.

V12 CHANGES
===========
1. Full detailed Excel is NOT generated automatically.
2. A smaller Management workbook is always available.
3. The full workbook is optional via sidebar checkbox.
4. Excel files are written to a temporary disk file using XlsxWriter
   constant_memory=True instead of pandas.to_excel(BytesIO).
5. Detailed Streamlit tables are shown ONE AT A TIME through selectboxes
   instead of rendering many large tables in multiple tabs.
6. Preview rows default to only 250.
7. Management Report, Management Breakdown, Exceptions, Customer Risk,
   Summary Control and Source Coverage remain available.
8. Critical financial rule remains:
      Expected Macron Token =
      VPS transaction_amount_minor - ₦100 per VPS transaction.

GITHUB FILES
============
Put these in the same repository folder:
    merge_first_reconciliation_v12.py
    requirements.txt

STREAMLIT CLOUD
===============
Main file:
    merge_first_reconciliation_v12.py

If an older app was originally deployed with another Python version and
still behaves unpredictably, redeploy it using Python 3.12 in Streamlit
Advanced settings.

FIRST LARGE-DATA RUN
====================
Leave:
    Build full detailed Excel workbook = OFF

This will complete the reconciliation and Management Report with the
lowest memory requirement.

Turn the detailed workbook option ON only when you need the full set of
transaction-level Excel sheets.
