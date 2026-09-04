"""
Tally integration package.

Talks to a running TallyPrime / Tally.ERP 9 instance over its XML-over-HTTP
interface. See INTEGRATION_PLAN.md for the intended scope.
"""

from tally_api.tally_client import TallyClient

__all__ = ['TallyClient']
