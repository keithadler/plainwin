#!/usr/bin/env python3
"""Write a flat-XML spreadsheet of N rows to stdout, for LibreOffice to convert to .xlsx for timing runs."""
import sys

rows = int(sys.argv[1]) if len(sys.argv) > 1 else 20000
print('<?xml version="1.0" encoding="UTF-8"?>')
print('<office:document xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" '
      'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" '
      'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" office:version="1.2" '
      'office:mimetype="application/vnd.oasis.opendocument.spreadsheet">')
print('<office:body><office:spreadsheet><table:table table:name="Rows">'
      '<table:table-column table:number-columns-repeated="5"/>')
headers = ["Id", "Name", "Cost", "Qty", "Total"]
print("<table:table-row>" + "".join(
    f'<table:table-cell office:value-type="string"><text:p>{h}</text:p></table:table-cell>' for h in headers)
    + "</table:table-row>")
for i in range(2, rows + 2):
    cost, qty = 10 + (i % 90), 1 + (i % 17)
    print(f'<table:table-row>'
          f'<table:table-cell office:value-type="float" office:value="{i}"><text:p>{i}</text:p></table:table-cell>'
          f'<table:table-cell office:value-type="string"><text:p>Item {i}</text:p></table:table-cell>'
          f'<table:table-cell office:value-type="float" office:value="{cost}"><text:p>{cost}</text:p></table:table-cell>'
          f'<table:table-cell office:value-type="float" office:value="{qty}"><text:p>{qty}</text:p></table:table-cell>'
          f'<table:table-cell table:formula="of:=[.C{i}]*[.D{i}]" office:value-type="float" '
          f'office:value="{cost * qty}"><text:p>{cost * qty}</text:p></table:table-cell>'
          f'</table:table-row>')
print('</table:table></office:spreadsheet></office:body></office:document>')
