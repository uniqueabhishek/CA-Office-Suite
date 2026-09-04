Tally API Integration Plan
Goal
Enable the CA Firm Web App to pull data (e.g., Trial Balance, Ledger Vouchers, Outstanding reports) from TallyPrime/Tally.ERP 9 and also push data if needed.

User Review Required
Tally Configuration: Tally must be running and "Client/Server configuration" enabled (usually port 9000).
Codebase Location: Still pending. Building as a standalone module tally_integration for now.
Proposed Changes
New Module: tally_integration
[NEW] tally_client.py
Class TallyClient:
__init__(host, port): Connection setup.
send_request(xml_payload): Base wrapper for HTTP POST.
get_companies(): List open companies in Tally.
get_trial_balance(): Fetch the Trial Balance report.
get_ledger_vouchers(ledger_name, data_range): Fetch transactions for a specific ledger.
execute_tdl(tdl_query): Advanced method to run custom TDL queries.
[NEW] response_parser.py
Parses the raw XML response from Tally into Python dictionaries or Pandas DataFrames for easy display in the webapp.
[NEW] tally_templates.py
XML templates for "Export Data" requests.
GET_COMPANIES_XML
GET_TRIAL_BALANCE_XML
Verification Plan
Manual Verification
Test Fetch: Run test_tally_fetch.py to pull the list of companies and a sample report.
Data Check: Compare the printed output or JSON with what is visible in Tally.