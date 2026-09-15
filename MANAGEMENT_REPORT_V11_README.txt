MERGE-FIRST KYC RECONCILIATION V11 — MANAGEMENT REPORT
=======================================================

V11 keeps the same merge-first reconciliation architecture and adds a complete
management-reporting layer.

NEW MANAGEMENT OUTPUTS
======================
11 Management Report
- Reporting periods for Bank, VPS, Paymeter and Macron
- Latest common date / source-period alignment control
- Source transaction counts and financial totals
- VPS customer payments, settlements and charges
- Expected Macron token = VPS transaction_amount_minor - ₦100
- Paymeter input, charges and token/net totals
- Macron transaction/token totals
- Paymeter↔Macron exact-match performance and unmatched counts/values
- VPS↔Bank exact/fallback match performance and unmatched counts/values
- KYC completeness, conflict, no-KYC and one-side-KYC indicators
- End-to-end reconciliation rates and statuses
- Amount-alignment / amount-difference indicators
- VPS↔Macron token presence and value-accuracy rates
- Paid/no-token customer count and financial exposure
- Token/no-VPS customer count and financial exposure
- Absolute and net value variance
- Customer Summary tie-out and source-coverage controls
- Paymeter cleaning metrics
- Missing-key/data-quality metrics for all four reports
- Overall critical/attention indicator count and conclusion

11A Mgmt Breakdown
- Final reconciliation status breakdown
- Amount status breakdown
- KYC status breakdown
- Paymeter-Macron merge status
- VPS-Bank merge status
- VPS-Macron status
- Counts, rates and applicable financial totals

11B Mgmt Exceptions
- All non-fully-reconciled transaction journeys
- Priority (CRITICAL / ATTENTION)
- Customer/KYC/account identifiers
- VPS paid amount
- Expected Macron token
- Actual Macron token
- Financial exposure/variance
- Key transaction references and matching status

11C Customer Risk
- Customer-level exception population
- Paid/no-token counts
- Token/no-VPS counts
- Value differences
- Unlinked identities
- Customer financial totals and KYC status

STREAMLIT DASHBOARD
===================
The app now displays headline KPIs and separate tabs for Management Report,
Management Breakdown, Management Exceptions and Customer Risk.

DEPLOYMENT
==========
Commit both files to GitHub:
    merge_first_reconciliation_v11.py
    requirements.txt

Set the Streamlit app entry point to:
    merge_first_reconciliation_v11.py

Then reboot/redeploy the Streamlit app.
