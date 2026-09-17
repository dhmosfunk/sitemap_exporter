# -*- coding: utf-8 -*-
#
# Sitemap Excel Exporter  -  Burp Suite extension (Jython / Python 2.7)
# -------------------------------------------------------------------------
# Exports the Burp sitemap to a styled .xlsx (same layout & colours as the
# reference template) for a domain you choose, with duplicate endpoints
# removed.
#
# Load via: Extensions -> Installed -> Add -> Extension type: Python
# (Requires the Jython standalone jar configured in
#  Extensions -> Extensions settings -> Python environment.)
#
# No external libraries required - the .xlsx writer below is pure standard
# library, so it runs inside Jython without openpyxl or any jar.
#
# UTF-8 note: request bodies are decoded as UTF-8 (with an ISO-8859-1
# fallback for binary payloads) so Greek and other multi-byte characters
# are exported correctly instead of being mangled by Burp's Latin-1
# bytesToString().
# -------------------------------------------------------------------------

import zipfile
import threading

from burp import IBurpExtender, ITab, IContextMenuFactory

from javax.swing import (JPanel, JLabel, JTextField, JCheckBox, JButton,
                         JScrollPane, JTextArea, JFileChooser, BorderFactory,
                         SwingUtilities, JMenuItem, Box)
from javax.swing.filechooser import FileNameExtensionFilter
from java.awt import (BorderLayout, GridBagLayout, GridBagConstraints, Insets,
                      Dimension, Font, Color)
from java.io import File
from java.util import ArrayList


# =========================================================================
#  Pure-stdlib .xlsx writer (works under Jython 2.7)
# =========================================================================

COLUMNS = [
    ("Name",            18),
    ("Method",          12),
    ("URL",             62),
    ("Payload",         40),
    ("Unauthenticated", 16),
    ("Escalation",      14),
    ("Observations",    18),
    ("Comments",        26),
]

S_DEFAULT, S_HEADER, S_DATA, S_GET, S_POST, S_DELETE, S_PUT, S_NAME = range(8)


def _clean(text):
    """Drop characters that are illegal in XML 1.0 (NUL and other control
    bytes that show up in raw request bodies and break Excel).

    Greek and other BMP text (code points up to 0xD7FF) is preserved."""
    out = []
    for ch in text:
        cp = ord(ch)
        if cp == 0x9 or cp == 0xA or cp == 0xD \
                or (0x20 <= cp <= 0xD7FF) \
                or (0xE000 <= cp <= 0xFFFD) \
                or cp >= 0x10000:
            out.append(ch)
    return "".join(out)


def _to_unicode(text):
    """Coerce anything to a Python Unicode string.

    Byte strings that contain UTF-8 (e.g. Greek text) are decoded as UTF-8,
    NOT via Python 2's default ASCII codec, which would mangle them or raise
    UnicodeDecodeError."""
    if text is None:
        return u""
    if isinstance(text, bytes):          # str under Jython/py2
        return text.decode("utf-8", "replace")
    try:
        return unicode(text)             # noqa: F821  Jython/py2
    except NameError:
        return str(text)                 # Python 3 fallback


def _esc(text):
    text = _clean(_to_unicode(text))
    return (text.replace("&", "&amp;").replace("<", "&lt;")
                .replace(">", "&gt;").replace('"', "&quot;")
                .replace("'", "&apos;"))


def _col_letter(idx):
    idx += 1
    s = ""
    while idx:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def _method_style(method):
    m = (method or "").upper()
    if m == "GET":
        return S_GET
    if m == "POST":
        return S_POST
    if m == "DELETE":
        return S_DELETE
    if m in ("PUT", "PATCH"):
        return S_PUT
    return S_DATA


CONTENT_TYPES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    '<Default Extension="xml" ContentType="application/xml"/>'
    '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    '<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
    '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
    '</Types>'
)

RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    '</Relationships>'
)

WORKBOOK = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    '<sheets><sheet name="Sitemap" sheetId="1" r:id="rId1"/></sheets>'
    '</workbook>'
)

WORKBOOK_RELS = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    '</Relationships>'
)

# Light / transparent theme.
# Fonts: 0 black  1 black-bold  2 green-bold  3 red-bold  4 amber-bold  5 blue-bold
# Fills: 0 none  1 gray125  2 lt-blue(header)  3 lt-green  4 lt-amber  5 lt-red  6 lt-blue(method)
STYLES = (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    '<fonts count="6">'
      '<font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/></font>'
      '<font><b/><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/></font>'
      '<font><b/><sz val="11"/><color rgb="FF227B3B"/><name val="Calibri"/></font>'
      '<font><b/><sz val="11"/><color rgb="FFB02A2A"/><name val="Calibri"/></font>'
      '<font><b/><sz val="11"/><color rgb="FF9A6B00"/><name val="Calibri"/></font>'
      '<font><b/><sz val="11"/><color rgb="FF2B5CB8"/><name val="Calibri"/></font>'
    '</fonts>'
    '<fills count="7">'
      '<fill><patternFill patternType="none"/></fill>'
      '<fill><patternFill patternType="gray125"/></fill>'
      '<fill><patternFill patternType="solid"><fgColor rgb="FFAED4F5"/><bgColor indexed="64"/></patternFill></fill>'
      '<fill><patternFill patternType="solid"><fgColor rgb="FFCDEFD6"/><bgColor indexed="64"/></patternFill></fill>'
      '<fill><patternFill patternType="solid"><fgColor rgb="FFFDE9B5"/><bgColor indexed="64"/></patternFill></fill>'
      '<fill><patternFill patternType="solid"><fgColor rgb="FFF7CDD1"/><bgColor indexed="64"/></patternFill></fill>'
      '<fill><patternFill patternType="solid"><fgColor rgb="FFCBDDF3"/><bgColor indexed="64"/></patternFill></fill>'
    '</fills>'
    '<borders count="2">'
      '<border><left/><right/><top/><bottom/><diagonal/></border>'
      '<border>'
        '<left style="thin"><color rgb="FFD9D9D9"/></left>'
        '<right style="thin"><color rgb="FFD9D9D9"/></right>'
        '<top style="thin"><color rgb="FFD9D9D9"/></top>'
        '<bottom style="thin"><color rgb="FFD9D9D9"/></bottom>'
        '<diagonal/>'
      '</border>'
    '</borders>'
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    '<cellXfs count="8">'
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
      '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>'
      '<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
      '<xf numFmtId="0" fontId="4" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
      '<xf numFmtId="0" fontId="3" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
      '<xf numFmtId="0" fontId="5" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
      '<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="top" wrapText="1"/></xf>'
    '</cellXfs>'
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    '<dxfs count="3">'
      '<dxf><font><color rgb="FFFF6B6B"/></font><fill><patternFill><bgColor rgb="FF5A1414"/></patternFill></fill></dxf>'
      '<dxf><font><color rgb="FF5FD98A"/></font><fill><patternFill><bgColor rgb="FF14532D"/></patternFill></fill></dxf>'
      '<dxf><font><u/><color rgb="FFFF6B6B"/></font><fill><patternFill><bgColor rgb="FF5A1414"/></patternFill></fill></dxf>'
    '</dxfs>'
    '</styleSheet>'
)


def _cell(ref, style, value):
    if value is None or value == "":
        return '<c r="%s" s="%d"/>' % (ref, style)
    return ('<c r="%s" s="%d" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
            % (ref, style, _esc(value)))


def _sheet_xml(rows):
    n = len(rows)
    last_data_row = n + 1 if n else 2
    ncols = len(COLUMNS)
    last_col = _col_letter(ncols - 1)

    p = []
    p.append('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>')
    p.append('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">')
    p.append('<dimension ref="A1:%s%d"/>' % (last_col, last_data_row))
    p.append('<sheetViews><sheetView workbookViewId="0" tabSelected="1">'
             '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
             '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
             '</sheetView></sheetViews>')
    p.append('<sheetFormatPr defaultRowHeight="15"/>')

    p.append('<cols>')
    for i, col in enumerate(COLUMNS):
        c = i + 1
        p.append('<col min="%d" max="%d" width="%d" customWidth="1"/>' % (c, c, col[1]))
    p.append('</cols>')

    p.append('<sheetData>')
    p.append('<row r="1" ht="20" customHeight="1">')
    for i, col in enumerate(COLUMNS):
        p.append(_cell('%s1' % _col_letter(i), S_HEADER, col[0]))
    p.append('</row>')

    for ridx, row in enumerate(rows):
        r = ridx + 2
        p.append('<row r="%d">' % r)
        p.append(_cell('A%d' % r, S_NAME, row.get("name", "")))
        p.append(_cell('B%d' % r, _method_style(row.get("method")), row.get("method", "")))
        p.append(_cell('C%d' % r, S_DATA, row.get("url", "")))
        p.append(_cell('D%d' % r, S_DATA, row.get("payload", "")))
        p.append(_cell('E%d' % r, S_DATA, row.get("unauth", "")))
        p.append(_cell('F%d' % r, S_DATA, row.get("escalation", "")))
        p.append(_cell('G%d' % r, S_DATA, row.get("observations", "")))
        p.append(_cell('H%d' % r, S_DATA, row.get("comments", "")))
        p.append('</row>')
    p.append('</sheetData>')

    p.append('<autoFilter ref="A1:%s1"/>' % last_col)

    cf_range = "E2:F%d" % max(last_data_row, 2)
    p.append('<conditionalFormatting sqref="%s">' % cf_range)
    p.append('<cfRule type="containsText" dxfId="0" priority="1" operator="containsText" stopIfTrue="1" text="Not OK">'
             '<formula>NOT(ISERROR(SEARCH("Not OK",E2)))</formula></cfRule>')
    p.append('<cfRule type="containsText" dxfId="1" priority="2" operator="containsText" text="OK">'
             '<formula>NOT(ISERROR(SEARCH("OK",E2)))</formula></cfRule>')
    p.append('</conditionalFormatting>')

    # Observations (G): any non-empty cell shown as a red "finding"
    obs_range = "G2:G%d" % max(last_data_row, 2)
    p.append('<conditionalFormatting sqref="%s">' % obs_range)
    p.append('<cfRule type="notContainsBlanks" dxfId="2" priority="3">'
             '<formula>LEN(TRIM(G2))&gt;0</formula></cfRule>')
    p.append('</conditionalFormatting>')

    p.append('<dataValidations count="1">')
    p.append('<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="%s">'
             '<formula1>"No - OK,Yes - Not OK"</formula1></dataValidation>' % cf_range)
    p.append('</dataValidations>')

    p.append('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
    p.append('</worksheet>')
    return u"".join(_to_unicode(part) for part in p)


def write_xlsx(path, rows):
    sheet = _sheet_xml(rows)
    zf = zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED)
    try:
        def w(name, data):
            # Always emit UTF-8 encoded bytes.  Unicode -> encode; existing
            # byte strings (the ASCII-only static parts) pass through.
            if isinstance(data, bytes):
                pass
            else:
                data = data.encode("utf-8")
            zf.writestr(name, data)
        w("[Content_Types].xml", CONTENT_TYPES)
        w("_rels/.rels", RELS)
        w("xl/workbook.xml", WORKBOOK)
        w("xl/_rels/workbook.xml.rels", WORKBOOK_RELS)
        w("xl/styles.xml", STYLES)
        w("xl/worksheets/sheet1.xml", sheet)
    finally:
        zf.close()


# =========================================================================
#  Sitemap helpers
# =========================================================================

STATIC_EXTS = (".js", ".css", ".png", ".jpg", ".jpeg", ".gif", ".svg",
               ".ico", ".woff", ".woff2", ".ttf", ".eot", ".map",
               ".webp", ".bmp", ".mp4", ".webm", ".pdf")

# order the sheet groups by method (GET first, DELETE last, others after)
METHOD_ORDER = {"GET": 0, "POST": 1, "PUT": 2, "PATCH": 3, "DELETE": 4,
                "HEAD": 5, "OPTIONS": 6}

NAME_BY_METHOD = {
    "GET": "GET request",
    "POST": "POST request",
    "DELETE": "DELETE request",
    "PUT": "PUT / PATCH request",
    "PATCH": "PUT / PATCH request",
}

MAX_PAYLOAD = 200000   # characters kept in the Payload cell


def _decode_body(byte_array, offset, length):
    """Decode a Java byte[] slice into a Python Unicode string.

    Tries STRICT UTF-8 first so Greek / multi-byte characters are preserved
    exactly.  If the bytes are not valid UTF-8 (e.g. a genuinely binary
    body), falls back to ISO-8859-1, which never fails and keeps every byte
    intact.  This replaces Burp's bytesToString(), which always uses Latin-1
    and therefore turns UTF-8 Greek into mojibake."""
    from java.lang import String
    try:
        from java.nio import ByteBuffer
        from java.nio.charset import Charset, CodingErrorAction
        decoder = Charset.forName("UTF-8").newDecoder()
        decoder.onMalformedInput(CodingErrorAction.REPORT)
        decoder.onUnmappableCharacter(CodingErrorAction.REPORT)
        buf = ByteBuffer.wrap(byte_array, offset, length)
        return unicode(decoder.decode(buf).toString())
    except Exception:
        # Not valid UTF-8 -> keep raw bytes losslessly via Latin-1.
        return unicode(String(byte_array, offset, length, "ISO-8859-1"))


def _host_matches(host, wanted, include_subdomains):
    host = (host or "").lower()
    wanted = (wanted or "").lower().strip()
    if wanted in ("", "*"):
        return True
    # strip scheme / path the user may have pasted
    if "://" in wanted:
        wanted = wanted.split("://", 1)[1]
    wanted = wanted.split("/", 1)[0].split(":", 1)[0]
    if host == wanted:
        return True
    if include_subdomains and host.endswith("." + wanted):
        return True
    return False


# =========================================================================
#  Burp extension
# =========================================================================

class BurpExtender(IBurpExtender, ITab, IContextMenuFactory):

    def registerExtenderCallbacks(self, callbacks):
        self._callbacks = callbacks
        self._helpers = callbacks.getHelpers()
        callbacks.setExtensionName("Sitemap Excel Exporter")

        self._build_ui()
        callbacks.addSuiteTab(self)
        callbacks.registerContextMenuFactory(self)
        self._log("Sitemap Excel Exporter loaded. Enter a domain and click Export.")

    # ---- ITab ----------------------------------------------------------
    def getTabCaption(self):
        return "Sitemap -> Excel"

    def getUiComponent(self):
        return self._panel

    # ---- context menu (right-click in sitemap) -------------------------
    def createMenuItems(self, invocation):
        menu = ArrayList()
        item = JMenuItem("Export sitemap to Excel...",
                         actionPerformed=lambda e: self._on_export(None))
        menu.add(item)
        return menu

    # ---- UI ------------------------------------------------------------
    def _build_ui(self):
        self._panel = JPanel(BorderLayout())
        self._panel.setBorder(BorderFactory.createEmptyBorder(12, 12, 12, 12))

        form = JPanel(GridBagLayout())
        g = GridBagConstraints()
        g.insets = Insets(4, 4, 4, 4)
        g.anchor = GridBagConstraints.WEST

        title = JLabel("Export Burp sitemap to a styled Excel (.xlsx) file")
        title.setFont(Font("SansSerif", Font.BOLD, 14))
        g.gridx = 0; g.gridy = 0; g.gridwidth = 3
        form.add(title, g)
        g.gridwidth = 1

        g.gridx = 0; g.gridy = 1
        form.add(JLabel("Domain:"), g)
        self._domain_field = JTextField(28)
        self._domain_field.setToolTipText(
            "e.g. example.com   (leave blank or '*' for every host in the sitemap)")
        g.gridx = 1; g.gridy = 1; g.gridwidth = 2
        g.fill = GridBagConstraints.HORIZONTAL
        form.add(self._domain_field, g)
        g.gridwidth = 1; g.fill = GridBagConstraints.NONE

        self._cb_subdomains = JCheckBox("Include subdomains", True)
        self._cb_dedup_query = JCheckBox("Ignore query string when de-duplicating", True)
        self._cb_static = JCheckBox("Exclude static resources (js/css/images/fonts)", True)
        self._cb_inscope = JCheckBox("Only in-scope items", False)
        self._cb_hasresp = JCheckBox("Only requests that received a response", False)

        y = 2
        for cb in (self._cb_subdomains, self._cb_dedup_query, self._cb_static,
                   self._cb_inscope, self._cb_hasresp):
            g.gridx = 0; g.gridy = y; g.gridwidth = 3
            form.add(cb, g)
            y += 1
        g.gridwidth = 1

        self._export_btn = JButton("Export to Excel...",
                                    actionPerformed=self._on_export)
        g.gridx = 0; g.gridy = y; g.gridwidth = 2
        form.add(self._export_btn, g)
        g.gridwidth = 1

        self._panel.add(form, BorderLayout.NORTH)

        self._log_area = JTextArea(14, 80)
        self._log_area.setEditable(False)
        self._log_area.setFont(Font("Monospaced", Font.PLAIN, 12))
        self._panel.add(JScrollPane(self._log_area), BorderLayout.CENTER)

    def _log(self, msg):
        def append():
            self._log_area.append(msg + "\n")
            self._log_area.setCaretPosition(self._log_area.getDocument().getLength())
        SwingUtilities.invokeLater(append)

    # ---- export --------------------------------------------------------
    def _on_export(self, event):
        domain = self._domain_field.getText().strip()

        chooser = JFileChooser()
        chooser.setDialogTitle("Save sitemap Excel")
        chooser.setFileFilter(FileNameExtensionFilter("Excel workbook (*.xlsx)", ["xlsx"]))
        suggested = (domain if domain and domain != "*" else "sitemap")
        suggested = suggested.replace("://", "_").replace("/", "_").replace("*", "all")
        chooser.setSelectedFile(File(suggested + "_sitemap.xlsx"))

        if chooser.showSaveDialog(self._panel) != JFileChooser.APPROVE_OPTION:
            return
        path = chooser.getSelectedFile().getAbsolutePath()
        if not path.lower().endswith(".xlsx"):
            path += ".xlsx"

        self._export_btn.setEnabled(False)
        opts = {
            "domain": domain,
            "subdomains": self._cb_subdomains.isSelected(),
            "dedup_query": self._cb_dedup_query.isSelected(),
            "static": self._cb_static.isSelected(),
            "inscope": self._cb_inscope.isSelected(),
            "hasresp": self._cb_hasresp.isSelected(),
        }
        t = threading.Thread(target=self._run_export, args=(path, opts))
        t.setDaemon(True)
        t.start()

    def _run_export(self, path, opts):
        try:
            self._log("-" * 60)
            self._log("Collecting sitemap for domain: '%s'" % (opts["domain"] or "*"))

            items = self._callbacks.getSiteMap(None)
            total = len(items) if items else 0
            self._log("Sitemap entries scanned: %d" % total)

            rows = []
            seen = set()
            hosts = set()

            for item in items:
                try:
                    info = self._helpers.analyzeRequest(item)
                except:
                    continue
                url = info.getUrl()
                if url is None:
                    continue

                host = url.getHost()
                if not _host_matches(host, opts["domain"], opts["subdomains"]):
                    continue
                if opts["inscope"] and not self._callbacks.isInScope(url):
                    continue
                if opts["hasresp"] and item.getResponse() is None:
                    continue

                path_part = url.getPath() or "/"
                if opts["static"] and self._is_static(path_part):
                    continue

                method = (info.getMethod() or "").upper()
                query = url.getQuery()

                if opts["dedup_query"]:
                    key = (method, host.lower(), path_part)
                else:
                    key = (method, host.lower(), path_part, query or "")
                if key in seen:
                    continue
                seen.add(key)
                hosts.add(host.lower())

                full_url = self._build_url(url)
                payload = self._extract_body(item, info, method)
                name = NAME_BY_METHOD.get(method, "%s request" % method)

                rows.append({
                    "name": name,
                    "method": method,
                    "url": full_url,
                    "payload": payload,
                    "unauth": "",
                    "escalation": "",
                    "observations": "",
                    "comments": "",
                })

            # group by method (all GETs together, all POSTs together, ...),
            # then order by URL within each method
            rows.sort(key=lambda r: (METHOD_ORDER.get(r["method"], 99),
                                     r["method"], r["url"].lower()))

            if not rows:
                self._log("No matching endpoints found. "
                          "Check the domain spelling / that it exists in the sitemap.")
                return

            write_xlsx(path, rows)
            self._log("Unique endpoints written: %d" % len(rows))
            if len(hosts) > 1:
                self._log("Hosts included: %s" % ", ".join(sorted(hosts)))
            self._log("Saved: %s" % path)
        except Exception as ex:
            self._log("ERROR: %s" % ex)
            import traceback
            self._log(traceback.format_exc())
        finally:
            SwingUtilities.invokeLater(lambda: self._export_btn.setEnabled(True))

    # ---- small helpers -------------------------------------------------
    def _is_static(self, path_part):
        low = path_part.lower()
        for ext in STATIC_EXTS:
            if low.endswith(ext):
                return True
        return False

    def _build_url(self, url):
        # reconstruct a clean absolute URL, hiding default ports
        scheme = url.getProtocol()
        host = url.getHost()
        port = url.getPort()
        path_part = url.getPath() or "/"
        query = url.getQuery()
        netloc = host
        if port != -1 and not (
                (scheme == "http" and port == 80) or
                (scheme == "https" and port == 443)):
            netloc = "%s:%d" % (host, port)
        full = "%s://%s%s" % (scheme, netloc, path_part)
        if query:
            full += "?" + query
        return full

    def _extract_body(self, item, info, method):
        if method in ("GET", "HEAD", "OPTIONS"):
            return ""
        try:
            request = item.getRequest()
            if request is None:
                return ""
            offset = info.getBodyOffset()
            length = len(request) - offset
            if length <= 0:
                return ""
            # Decode the raw body bytes as UTF-8 ourselves so Greek (and any
            # other multi-byte) characters survive.  Burp's bytesToString()
            # uses ISO-8859-1, which would turn e.g. "ασφάλεια" into mojibake.
            body = _decode_body(request, offset, length)
            body = body.strip()
            if len(body) > MAX_PAYLOAD:
                body = body[:MAX_PAYLOAD] + u" ...[truncated]"
            return body
        except:
            return ""