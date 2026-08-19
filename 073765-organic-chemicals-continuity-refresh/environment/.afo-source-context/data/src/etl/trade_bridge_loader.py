"""Utility notes for the monthly trade bridge loader.

This module documents field-level expectations for analysts who recreate the
bridge in a workbook. It is not a production loader in this benchmark.
"""

TOTAL_ROW_RULE = 'partnerCode == 0 and partner2Code == 0 and motCode == 0'
BRIDGE_CONFIG_ID = 'MODEL-BRIDGE-23-03'
QUALITY_FLAGS = ['QUAL-CENSUS-HTML-01', 'QUAL-BLS-EMPTY-01', 'QUAL-COMTRADE-GAP-01']

def describe_expected_grain():
    return 'one catalog row per record_id before any trend or partner ranking step'
