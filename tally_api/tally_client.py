"""
HTTP client for TallyPrime / Tally.ERP 9.
"""

import requests

from tally_api import tally_templates


class TallyClient:
    """
    Client for interacting with TallyPrime/ERP 9 via XML over HTTP.
    """
    def __init__(self, host="localhost", port=9000):
        self.host = host
        self.port = port
        self.base_url = f"http://{host}:{port}"
        self.headers = {'Content-Type': 'text/xml; charset=utf-8'}

    def _send_request(self, xml_payload):
        """
        Sends the XML payload to Tally and returns the response text.
        """
        try:
            response = requests.post(self.base_url, data=xml_payload, headers=self.headers)
            response.raise_for_status()
            return response.text
        except requests.exceptions.RequestException as e:
            print(f"Error connecting to Tally: {e}")
            return None

    def check_connection(self):
        """
        Checks if Tally is reachable.
        """
        print(f"Checking connection to {self.base_url}...")
        # Sending a minimal request or just checking if the server accepts connections
        try:
            # Tally usually responds to a simple GET or an invalid POST with some data
            # But creating a valid 'List of Companies' request is safer
            response = self._send_request(tally_templates.GET_COMPANIES_XML)
            if response:
                print("Connection successful!")
                return True
            return False
        except Exception:
            return False

    def get_companies(self):
        """
        Fetches the list of companies currently open in Tally.
        """
        print("Fetching companies...")
        response_xml = self._send_request(tally_templates.GET_COMPANIES_XML)
        # TODO: Parse this XML to return a list of names
        return response_xml

    def get_trial_balance(self, company_name):
        """
        Fetches the Trial Balance for a specific company.
        """
        print(f"Fetching Trial Balance for {company_name}...")
        # Format the template with the company name
        xml_payload = tally_templates.GET_TRIAL_BALANCE_XML.format(company_name=company_name)
        response_xml = self._send_request(xml_payload)
        return response_xml
