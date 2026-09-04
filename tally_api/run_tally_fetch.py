"""
Manual smoke script: connects to a running Tally instance and prints a sample
of what it returns. Not a pytest test - it needs Tally listening on the host.
"""

from tally_api.tally_client import TallyClient


def main():
    # Allow user to pass host/port via args, else default
    host = "localhost"
    port = 9000

    client = TallyClient(host, port)

    if not client.check_connection():
        print("Failed to connect to Tally. Please ensure Tally is running and configured on port 9000.")
        return

    # Fetch Companies
    print("\n--- Companies ---")
    companies_xml = client.get_companies()
    if companies_xml:
        print(companies_xml[:500] + "...")  # Print first 500 chars to avoid flooding

    # Example: Fetch TB for "Demo Company"
    # In a real scenario, we'd parse the companies_xml to get a real name
    print("\n--- Trial Balance (Demo Company) ---")
    tb_xml = client.get_trial_balance("Demo Company")
    if tb_xml:
        print(tb_xml[:500] + "...")


if __name__ == "__main__":
    main()
