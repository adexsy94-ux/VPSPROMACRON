MERGE-FIRST KYC RECONCILIATION V9
=================================

This version follows the requested processing order exactly:

0. CLEAN PAYMETER FIRST
   - Keep the first/original Address cell.
   - Delete Address fragments that spilled into extra CSV cells.
   - Shift Transaction Amount, Input Amount, System Charge, RRN, Reference and later fields back to their intended columns.
   - Ignore blank export columns after Status Checked.
   - Only the cleaned Paymeter report is used downstream.

1. MERGE PAYMETER WITH MACRON
   Primary match:
       Paymeter Reference = Macron REFERENCE ID
   Output:
       02 PM Macron Merge

2. MERGE VPS WITH BANK STATEMENT
   Primary match:
       VPS session_id found inside Bank Narration
   Controlled fallback:
       unique VPS settled amount + date
   Bank customer-payment population is Bank CREDIT rows.
   Output:
       03 VPS Bank Merge

3. BUILD KYC FROM THE TWO MERGED REPORTS
   Base KYC population:
       every unique VPS virtual_acct_no
   Primary Paymeter identity link:
       VPS settlement_ref = Paymeter RRN
   Fallback identity:
       VPS virtual account = Paymeter Account Number
   Once a Paymeter account is proven, its Paymeter/Macron transaction history is used to attach customer name, phone, meter, Macron account and Macron meter.
   Output:
       04 KYC

4. ATTACH KYC TO BOTH MERGED REPORTS
   Outputs:
       05 PM Macron + KYC
       06 VPS Bank + KYC

5. RECONCILE BOTH KYC-ENRICHED MERGED REPORTS
   Primary transaction bridge:
       VPS settlement_ref = Paymeter RRN
   Controlled fallback:
       same KYC + matching gross/token amount + nearby date
   Output:
       07 Full Reconciliation

6. VPS AGAINST MACRON
   CRITICAL RULE:
       VPS transaction_amount_minor = total fund actually paid by customer
       Expected Macron Token = transaction_amount_minor - ₦100
   Output:
       08 VPS vs Macron

7. CUSTOMER SUMMARY + HARD TIE-OUT
   Customer Summary includes formal KYC customers AND unlinked Paymeter/Macron/Bank identities so no source figure disappears.
   10 Summary Control must tie Customer Summary back to Bank, VPS, cleaned Paymeter and Macron. The app stops if a tie-out fails.

RUN
===
pip install -r requirements_merge_first_v9.txt
streamlit run merge_first_reconciliation_v9.py

PERFORMANCE TEST
================
The merge-first engine was stress-tested with a doubled transaction population:
- VPS: 5,340 rows
- Paymeter: 5,722 rows
- Macron: 5,506 rows
- Bank credits: 5,666 rows
- Final reconciliation: 6,704 rows

The reconciliation calculation completed in under 6 seconds in the test environment, and source coverage confirmed 0 missing and 0 duplicate source-row uses. Excel generation is a separate stage and can take longer for large workbooks.
