"""
HTTP client for TallyPrime / Tally.ERP 9.
"""

import logging

import requests

from tally_api import tally_templates

logger = logging.getLogger(__name__)

# Tally answers from the same machine, so a request that has not come back in
# this long is not slow, it is stuck. Without a bound, a Tally that accepts the
# connection and then never replies - busy, or with no company loaded - blocks
# the caller forever, which is worst in check_connection, whose whole job is to
# answer quickly.
DEFAULT_TIMEOUT_SECONDS = 10


class TallyClient:
    """
    Client for interacting with TallyPrime/ERP 9 via XML over HTTP.
    """

    def __init__(self, host="localhost", port=9000, timeout=DEFAULT_TIMEOUT_SECONDS):
        self.host = host
        self.port = port
        self.timeout = timeout
        self.base_url = f"http://{host}:{port}"
        self.headers = {"Content-Type": "text/xml; charset=utf-8"}

    def _send_request(self, xml_payload):
        """
        Sends the XML payload to Tally and returns the response text.
        """
        try:
            response = requests.post(
                self.base_url,
                data=xml_payload,
                headers=self.headers,
                timeout=self.timeout,
            )
            response.raise_for_status()
            return response.text
        except requests.exceptions.RequestException:
            logger.warning("Error connecting to Tally at %s", self.base_url, exc_info=True)
            return None

    def check_connection(self):
        """
        Checks if Tally is reachable.

        Every connection failure already arrives here as None from
        _send_request, so there is nothing left to catch: a second, broader
        guard would only hide faults in this module.
        """
        logger.info("Checking connection to %s...", self.base_url)
        # Tally usually responds to a simple GET or an invalid POST with some data
        # But creating a valid 'List of Companies' request is safer
        return self._send_request(tally_templates.GET_COMPANIES_XML) is not None

    def get_companies(self):
        """
        Fetches the list of companies currently open in Tally.
        """
        logger.info("Fetching companies...")
        response_xml = self._send_request(tally_templates.GET_COMPANIES_XML)
        # TODO: Parse this XML to return a list of names
        return response_xml

    def get_trial_balance(self, company_name):
        """
        Fetches the Trial Balance for a specific company.
        """
        logger.info("Fetching Trial Balance for %s...", company_name)
        # Format the template with the company name
        xml_payload = tally_templates.GET_TRIAL_BALANCE_XML.format(company_name=company_name)
        response_xml = self._send_request(xml_payload)
        return response_xml
