"""
XML Templates for Tally Integration
This file contains the raw TDSL (Tally Definition Language) XML structures
needed to communicate with TallyPrime / Tally.ERP 9.
"""

# The standard envelope for all Tally requests
ENVELOPE_TEMPLATE = """<ENVELOPE>
    <HEADER>
        <TALLYREQUEST>Export Data</TALLYREQUEST>
    </HEADER>
    <BODY>
        <EXPORTDATA>
            <REQUESTDESC>
                <REPORTNAME>{report_name}</REPORTNAME>
                <STATICVARIABLES>
                    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
                    {static_variables}
                </STATICVARIABLES>
            </REQUESTDESC>
        </EXPORTDATA>
    </BODY>
</ENVELOPE>"""

# Template to fetch the list of companies open in Tally
# We use a simple function call "List of Companies" if available,
# or we can just ask for "List of Accounts" which usually returns basics.
# A more reliable way is to query the "Company" collection.
GET_COMPANIES_XML = """<ENVELOPE>
    <HEADER>
        <TALLYREQUEST>Export Data</TALLYREQUEST>
    </HEADER>
    <BODY>
        <EXPORTDATA>
            <REQUESTDESC>
                <REPORTNAME>List of Companies</REPORTNAME>
                <STATICVARIABLES>
                    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
                </STATICVARIABLES>
            </REQUESTDESC>
        </EXPORTDATA>
    </BODY>
</ENVELOPE>"""

# A simple TDL to fetch Ledger Names (Trial Balance basics)
# We can use Tally's built-in "Trial Balance" report name.
GET_TRIAL_BALANCE_XML = """<ENVELOPE>
    <HEADER>
        <TALLYREQUEST>Export Data</TALLYREQUEST>
    </HEADER>
    <BODY>
        <EXPORTDATA>
            <REQUESTDESC>
                <REPORTNAME>Trial Balance</REPORTNAME>
                <STATICVARIABLES>
                    <SVEXPORTFORMAT>$$SysName:XML</SVEXPORTFORMAT>
                    <SVCOMPANYNAME>{company_name}</SVCOMPANYNAME>
                </STATICVARIABLES>
            </REQUESTDESC>
        </EXPORTDATA>
    </BODY>
</ENVELOPE>"""
