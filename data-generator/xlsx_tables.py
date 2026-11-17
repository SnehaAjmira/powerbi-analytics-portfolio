"""Helpers for writing large Excel files fast.

xlsxwriter's constant_memory mode streams rows to disk, which keeps memory low
for files with hundreds of thousands of rows, but it cannot create Excel
"Tables" (ListObjects). Power Query often reads named tables, so this module
streams the rows first and then injects the table definitions into the .xlsx
package afterwards.
"""
import datetime as dt
import os
import re
import shutil
import tempfile
import zipfile
from xml.sax.saxutils import escape

import xlsxwriter


def _col_letter(n: int) -> str:
    s = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


class StreamingWorkbook:
    """Write sheets row by row; optionally register each sheet as an Excel Table."""

    def __init__(self, path):
        self.path = path
        self.wb = xlsxwriter.Workbook(path, {"constant_memory": True, "strings_to_numbers": False})
        self.date_fmt = self.wb.add_format({"num_format": "yyyy-mm-dd"})
        self.dt_fmt = self.wb.add_format({"num_format": "yyyy-mm-dd hh:mm:ss"})
        self.hdr_fmt = self.wb.add_format({"bold": True})
        self.tables = []  # (sheet_index, table_name, headers, nrows)
        self._sheet_count = 0

    def write_sheet(self, sheet_name, headers, rows, table_name=None):
        ws = self.wb.add_worksheet(sheet_name)
        self._sheet_count += 1
        ws.write_row(0, 0, headers, self.hdr_fmt)
        r = 0
        for r, row in enumerate(rows, start=1):
            for c, v in enumerate(row):
                if v is None or v == "":
                    continue
                if isinstance(v, dt.datetime):
                    if v.hour or v.minute or v.second:
                        ws.write_datetime(r, c, v, self.dt_fmt)
                    else:
                        ws.write_datetime(r, c, v, self.date_fmt)
                elif isinstance(v, dt.date):
                    ws.write_datetime(r, c, dt.datetime(v.year, v.month, v.day), self.date_fmt)
                elif isinstance(v, (int, float)) and not isinstance(v, bool):
                    ws.write_number(r, c, v)
                else:
                    ws.write_string(r, c, str(v))
        if table_name:
            self.tables.append((self._sheet_count, table_name, headers, max(r, 1)))
        return r

    def close(self):
        self.wb.close()
        if self.tables:
            _inject_tables(self.path, self.tables)


def _inject_tables(path, tables):
    tmp_dir = tempfile.mkdtemp()
    tmp_path = os.path.join(tmp_dir, "out.xlsx")
    with zipfile.ZipFile(path) as zin, zipfile.ZipFile(tmp_path, "w", zipfile.ZIP_DEFLATED) as zout:
        names = zin.namelist()
        ct = zin.read("[Content_Types].xml").decode()
        sheet_xml = {}
        for i, (sheet_idx, tname, headers, nrows) in enumerate(tables, start=1):
            ref = f"A1:{_col_letter(len(headers) - 1)}{nrows + 1}"
            cols = "".join(
                f'<tableColumn id="{j}" name="{escape(h, {chr(34): "&quot;"})}"/>'
                for j, h in enumerate(headers, start=1)
            )
            table_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="{i}" '
                f'name="{tname}" displayName="{tname}" ref="{ref}" totalsRowShown="0">'
                f'<autoFilter ref="{ref}"/><tableColumns count="{len(headers)}">{cols}</tableColumns>'
                '<tableStyleInfo name="TableStyleLight9" showFirstColumn="0" showLastColumn="0" '
                'showRowStripes="1" showColumnStripes="0"/></table>'
            )
            zout.writestr(f"xl/tables/table{i}.xml", table_xml)
            rels = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rIdTbl1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" '
                f'Target="../tables/table{i}.xml"/></Relationships>'
            )
            zout.writestr(f"xl/worksheets/_rels/sheet{sheet_idx}.xml.rels", rels)
            ct = ct.replace(
                "</Types>",
                f'<Override PartName="/xl/tables/table{i}.xml" '
                'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>',
            )
            sheet_xml[f"xl/worksheets/sheet{sheet_idx}.xml"] = True

        for n in names:
            if n == "[Content_Types].xml":
                zout.writestr(n, ct)
                continue
            if n in sheet_xml:
                data = zin.read(n).decode()
                # tableParts must be the last child of <worksheet>
                if 'xmlns:r=' not in data[:500]:
                    data = data.replace(
                        "<worksheet ",
                        '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ',
                        1,
                    )
                data = re.sub(r"</worksheet>\s*$", '<tableParts count="1"><tablePart r:id="rIdTbl1"/></tableParts></worksheet>', data)
                zout.writestr(n, data)
                continue
            zout.writestr(n, zin.read(n))
    shutil.move(tmp_path, path)
    shutil.rmtree(tmp_dir, ignore_errors=True)
