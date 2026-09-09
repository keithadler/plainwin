#!/usr/bin/env python3
"""
Writes the flat-XML sources for the pictures in the README, and converts them with LibreOffice.

    python3 scripts/make-demo.py            # into docs/demo/

The pictures should show the app doing the work people actually do, so the sheet has real headings, real
formatting and real formulas, and the document has headings, a table and a list. The content is invented:
Pine Street Holdings, Sam Rivera, Woodland Ave.
"""
import subprocess, os

OUT = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "docs", "demo")

NS = '''xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
 xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0"
 xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
 xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"
 xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0"
 xmlns:number="urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0"
 xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0"
 xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"
 xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0"
 xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"
 xmlns:dc="http://purl.org/dc/elements/1.1/"
 xmlns:calcext="urn:org:documentfoundation:names:experimental:calc:xmlns:calcext:1.0"
 office:version="1.2"'''


def esc(s):
    """Text content. Attributes go through attr(), which also has to deal with the quotes in a COUNTIF."""
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def attr(s):
    return esc(s).replace('"', "&quot;")


# ---------------- spreadsheet ----------------

SHEET_STYLES = '''
 <office:automatic-styles>
  <number:number-style style:name="N0"><number:number number:decimal-places="0" number:min-integer-digits="1" number:grouping="true"/></number:number-style>
  <number:percentage-style style:name="NP"><number:number number:decimal-places="1" number:min-integer-digits="1"/><number:text>%</number:text></number:percentage-style>
  <style:style style:name="co1" style:family="table-column"><style:table-column-properties style:column-width="2.0in"/></style:style>
  <style:style style:name="co2" style:family="table-column"><style:table-column-properties style:column-width="1.05in"/></style:style>
  <style:style style:name="ttl" style:family="table-cell"><style:text-properties fo:font-weight="bold" fo:font-size="13pt"/></style:style>
  <style:style style:name="hdr" style:family="table-cell"><style:text-properties fo:font-weight="bold"/><style:table-cell-properties fo:border-bottom="0.5pt solid #808080"/></style:style>
  <style:style style:name="num" style:family="table-cell" style:data-style-name="N0"/>
  <style:style style:name="pct" style:family="table-cell" style:data-style-name="NP"/>
  <style:style style:name="tot" style:family="table-cell" style:data-style-name="N0"><style:text-properties fo:font-weight="bold"/><style:table-cell-properties fo:border-top="0.5pt solid #808080"/></style:style>
  <style:style style:name="totl" style:family="table-cell"><style:text-properties fo:font-weight="bold"/><style:table-cell-properties fo:border-top="0.5pt solid #808080"/></style:style>
 </office:automatic-styles>'''


def cell(v=None, style=None, formula=None, pct=False):
    """
    A cell carries its formula and its answer, the way Excel writes one. LibreOffice does not recalculate on a
    headless convert, so a formula with no cached value would show as nought in every picture.
    """
    a = ''
    if style:
        a += ' table:style-name="%s"' % attr(style)
    if formula:
        a += ' table:formula="of:%s"' % attr(formula)
    if v is None:
        return '<table:table-cell%s/>' % a
    if isinstance(v, str):
        return '<table:table-cell%s office:value-type="string"><text:p>%s</text:p></table:table-cell>' % (a, esc(v))
    if pct:
        return ('<table:table-cell%s office:value-type="percentage" office:value="%s">'
                '<text:p>%s%%</text:p></table:table-cell>' % (a, v / 100.0, round(v, 1)))
    shown = ("{:,}".format(v) if abs(v) >= 1000 else str(v))
    return ('<table:table-cell%s office:value-type="float" office:value="%s">'
            '<text:p>%s</text:p></table:table-cell>' % (a, v, shown))


def row(cells):
    return "<table:table-row>" + "".join(cells) + "</table:table-row>\n"


def sheet(name, rows, cols=5):
    cols_xml = ('<table:table-column table:style-name="co1"/>'
                '<table:table-column table:style-name="co2" table:number-columns-repeated="%d"/>' % (cols - 1))
    return '<table:table table:name="%s">\n%s\n%s</table:table>\n' % (name, cols_xml, "".join(rows))


def spreadsheet():
    lines = [("Licences", 24000, 26400),
             ("Professional services", 18500, 21750),
             ("Hardware", 6200, 5100),
             ("Support contracts", 9400, 9400),
             ("Training", 3100, 4600),
             ("Hosting", 7250, 7900)]
    q1_total = sum(l[1] for l in lines)
    q2_total = sum(l[2] for l in lines)
    changes = [l[2] - l[1] for l in lines]

    rows = [row([cell("Pine Street Holdings", "ttl")]),
            row([cell("Revenue by line, quarter ending 30 June")]),
            row([]),
            row([cell("Item", "hdr"), cell("Q1", "hdr"), cell("Q2", "hdr"), cell("Change", "hdr"), cell("Share", "hdr")])]
    first = 5
    last = first + len(lines) - 1
    total_row = last + 1
    for i, (name, q1, q2) in enumerate(lines):
        r = first + i
        rows.append(row([cell(name),
                         cell(q1, "num"), cell(q2, "num"),
                         cell(q2 - q1, "num", "=[.C%d]-[.B%d]" % (r, r)),
                         cell(round(q2 / q2_total * 100, 1), "pct", "=[.C%d]/[.C$%d]*100" % (r, total_row), pct=True)]))
    rows.append(row([cell("Total", "totl"),
                     cell(q1_total, "tot", "=SUM([.B%d:.B%d])" % (first, last)),
                     cell(q2_total, "tot", "=SUM([.C%d:.C%d])" % (first, last)),
                     cell(q2_total - q1_total, "tot", "=SUM([.D%d:.D%d])" % (first, last)),
                     cell()]))
    rows.append(row([]))
    riser = lines[changes.index(max(changes))][0]
    rows.append(row([cell("Biggest riser", "hdr"),
                     cell(riser, "hdr",
                          '=INDEX([.A%d:.A%d];MATCH(MAX([.D%d:.D%d]);[.D%d:.D%d];0))'
                          % (first, last, first, last, first, last))]))
    rows.append(row([cell("Lines above 9,000"),
                     cell(sum(1 for l in lines if l[2] > 9000),
                          None, '=COUNTIF([.C%d:.C%d];">9000")' % (first, last))]))
    rows.append(row([cell("Growth from the lines that grew"),
                     cell(sum(c for c in changes if c > 0), "num",
                          '=SUMIF([.D%d:.D%d];">0")' % (first, last))]))

    data = [("PS-1041", "Woodland Ave Partners", "12 Apr", 8400),
            ("PS-1042", "Harbour Lane Clinic", "18 Apr", 3150),
            ("PS-1043", "Sam Rivera Consulting", "02 May", 12600),
            ("PS-1044", "Woodland Ave Partners", "14 May", 5900),
            ("PS-1045", "Kestrel Freight", "27 May", 4750),
            ("PS-1046", "Harbour Lane Clinic", "09 Jun", 6300),
            ("PS-1047", "Kestrel Freight", "21 Jun", 9050)]
    detail = [row([cell("Invoice", "hdr"), cell("Client", "hdr"), cell("Raised", "hdr"), cell("Amount", "hdr")])]
    for inv, client, raised, amt in data:
        detail.append(row([cell(inv), cell(client), cell(raised), cell(amt, "num")]))
    detail.append(row([cell("Total", "totl"), cell(), cell(),
                       cell(sum(d[3] for d in data), "tot", "=SUM([.D2:.D%d])" % (1 + len(data)))]))

    notes = [row([cell("Written up by Sam Rivera on 3 July.")]),
             row([cell("Hardware is down because the Woodland Ave order slipped into Q3.")]),
             row([cell("Hosting moves to the annual plan in Q4.")])]

    body = sheet("Summary", rows) + sheet("Detail", detail, cols=4) + sheet("Notes", notes, cols=2)
    return ('<?xml version="1.0" encoding="UTF-8"?>\n<office:document %s '
            'office:mimetype="application/vnd.oasis.opendocument.spreadsheet">\n%s\n'
            '<office:body><office:spreadsheet>\n%s</office:spreadsheet></office:body></office:document>\n'
            % (NS, SHEET_STYLES, body))


# ---------------- document ----------------

def document():
    def para(s):
        return '<text:p text:style-name="Standard">%s</text:p>\n' % esc(s)

    def head(s, lvl):
        return '<text:h text:style-name="Heading_20_%d" text:outline-level="%d">%s</text:h>\n' % (lvl, lvl, esc(s))

    body = head("Woodland Ave: annual review", 1)
    body += para("Prepared by Sam Rivera for Pine Street Holdings, 3 July.")
    body += head("What changed this year", 2)
    body += para("The building was full for eleven of the twelve months, and the one vacancy was let inside three "
                 "weeks. Repairs came in under the amount set aside, mostly because the roof work was deferred "
                 "rather than avoided.")
    body += '<text:list text:style-name="L1">\n'
    for item in ["Occupancy 96 per cent, against 91 per cent last year.",
                 "Repairs 14,200, against 17,000 set aside.",
                 "One tenant on a month to month, everyone else on a term.",
                 "Roof survey booked for September."]:
        body += '<text:list-item><text:p text:style-name="Standard">%s</text:p></text:list-item>\n' % esc(item)
    body += '</text:list>\n'
    body += head("The numbers", 2)
    rows = [("Line", "This year", "Last year"),
            ("Rent collected", "182,400", "171,900"),
            ("Repairs", "14,200", "17,000"),
            ("Insurance", "9,800", "9,100"),
            ("Management", "12,600", "12,600")]
    body += '<table:table table:name="Numbers"><table:table-column table:number-columns-repeated="3"/>\n'
    for i, r in enumerate(rows):
        style = "Table_20_Heading" if i == 0 else "Table_20_Contents"
        body += "<table:table-row>" + "".join(
            '<table:table-cell office:value-type="string"><text:p text:style-name="%s">%s</text:p></table:table-cell>'
            % (style, esc(c)) for c in r) + "</table:table-row>\n"
    body += '</table:table>\n'
    body += head("What to do next", 2)
    body += para("Hold the rent where it is for the tenant on a month to month until the roof is done, then move "
                 "everyone onto the same renewal date. The saving on repairs should stay set aside rather than be "
                 "taken out.")
    body += para("Nothing in this note is agreed until the board has seen it.")
    return ('<?xml version="1.0" encoding="UTF-8"?>\n<office:document %s '
            'office:mimetype="application/vnd.oasis.opendocument.text">\n'
            '<office:body><office:text>\n%s</office:text></office:body></office:document>\n' % (NS, body))


# ---------------- presentation ----------------

def presentation():
    slides = [("Woodland Ave", ["Annual review", "Pine Street Holdings", "3 July"]),
              ("The year in one line", ["Full for eleven months of twelve",
                                        "Repairs under the amount set aside",
                                        "One vacancy, let in three weeks"]),
              ("Rent collected", ["182,400 this year", "171,900 last year", "Up 6 per cent"]),
              ("What it cost", ["Repairs 14,200", "Insurance 9,800", "Management 12,600"]),
              ("The roof", ["Survey booked for September", "Work deferred, not avoided",
                            "The money stays set aside"]),
              ("What we are asking", ["Hold the rent until the roof is done",
                                      "Move everyone to one renewal date",
                                      "The board sees this before anything is agreed"])]
    body = ""
    for i, (title, bullets) in enumerate(slides):
        body += '<draw:page draw:name="Slide %d" draw:master-page-name="Default">\n' % (i + 1)
        body += ('<draw:frame presentation:class="title" svg:width="9in" svg:height="1.2in" svg:x="0.5in" '
                 'svg:y="0.5in"><draw:text-box><text:p>%s</text:p></draw:text-box></draw:frame>\n' % esc(title))
        body += ('<draw:frame presentation:class="outline" svg:width="9in" svg:height="4in" svg:x="0.5in" '
                 'svg:y="2in"><draw:text-box>')
        for b in bullets:
            body += '<text:p>%s</text:p>' % esc(b)
        body += '</draw:text-box></draw:frame>\n</draw:page>\n'
    return ('<?xml version="1.0" encoding="UTF-8"?>\n<office:document %s '
            'office:mimetype="application/vnd.oasis.opendocument.presentation">\n'
            '<office:body><office:presentation>\n%s</office:presentation></office:body></office:document>\n'
            % (NS, body))


def convert(flat, name, target):
    os.makedirs(OUT, exist_ok=True)
    src = os.path.join(OUT, name)
    with open(src, "w") as fh:
        fh.write(flat)
    subprocess.run(["soffice", "--headless", "--convert-to", target, "--outdir", OUT, src], check=True,
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    os.remove(src)


if __name__ == "__main__":
    convert(spreadsheet(), "quarter.fods", "xlsx")
    convert(document(), "review.fodt", "docx")
    convert(presentation(), "woodland.fodp", "pptx")
    for f in sorted(os.listdir(OUT)):
        print(os.path.join("docs/demo", f))
