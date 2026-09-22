"""Write the Temperature / Humidity Excel workbook in the warehouse format."""

from __future__ import annotations

from array import array
from datetime import date, datetime
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo
from xml.sax.saxutils import escape

# Excel's date serial origin (with the 1900 leap-year bug, correct after 1900-03-01).
_EXCEL_ORD = date(1899, 12, 30).toordinal()

_NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
_NS_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types"

_CONTENT_TYPES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="{_NS_CT}">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
</Types>
"""

_ROOT_RELS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{_NS_PKG}">
  <Relationship Id="rId1" Type="{_NS_DOC}/officeDocument" Target="xl/workbook.xml"/>
</Relationships>
"""

_WB_RELS = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="{_NS_PKG}">
  <Relationship Id="rId1" Type="{_NS_DOC}/worksheet" Target="worksheets/sheet1.xml"/>
  <Relationship Id="rId2" Type="{_NS_DOC}/worksheet" Target="worksheets/sheet2.xml"/>
  <Relationship Id="rId3" Type="{_NS_DOC}/styles" Target="styles.xml"/>
</Relationships>
"""

_WORKBOOK = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="{_NS_MAIN}" xmlns:r="{_NS_DOC}">
  <sheets>
    <sheet name="Temperature" sheetId="1" r:id="rId1"/>
    <sheet name="Humidity" sheetId="2" r:id="rId2"/>
  </sheets>
</workbook>
"""

_STYLES = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="{_NS_MAIN}">
  <numFmts count="3">
    <numFmt numFmtId="164" formatCode="DD/MM/YYYY"/>
    <numFmt numFmtId="165" formatCode="HH:MM"/>
    <numFmt numFmtId="166" formatCode="0.0"/>
  </numFmts>
  <fonts count="2">
    <font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font>
    <font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/><family val="2"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF305496"/><bgColor rgb="FF305496"/></patternFill></fill>
  </fills>
  <borders count="1">
    <border><left/><right/><top/><bottom/><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="5">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>
    <xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
    <xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>
"""


def write_workbook(
    output_path: Path,
    loggers: list[str],
    timestamps: list[datetime],
    temps: dict[str, array],
    hums: dict[str, array],
    on_progress=None,
) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)

    n_rows = len(timestamps)
    n_cols = 2 + len(loggers)
    letters = [_col_name(i) for i in range(n_cols)]
    excel_dates = [ts.toordinal() - _EXCEL_ORD for ts in timestamps]
    excel_times = [(ts.hour * 60 + ts.minute) / 1440 for ts in timestamps]
    last_ref = f"{letters[-1]}{n_rows + 1}" if n_cols else "A1"

    temp_cols = [temps[name] for name in loggers]
    hum_cols = [hums[name] for name in loggers]

    with ZipFile(output_path, "w", compression=ZIP_DEFLATED, compresslevel=1) as zf:
        _zip_text(zf, "[Content_Types].xml", _CONTENT_TYPES)
        _zip_text(zf, "_rels/.rels", _ROOT_RELS)
        _zip_text(zf, "xl/workbook.xml", _WORKBOOK)
        _zip_text(zf, "xl/_rels/workbook.xml.rels", _WB_RELS)
        _zip_text(zf, "xl/styles.xml", _STYLES)
        _stream_sheet(
            zf,
            "xl/worksheets/sheet1.xml",
            loggers,
            letters,
            last_ref,
            excel_dates,
            excel_times,
            temp_cols,
            tab_color="548235",
            on_progress=on_progress,
            label="Temperature",
        )
        _stream_sheet(
            zf,
            "xl/worksheets/sheet2.xml",
            loggers,
            letters,
            last_ref,
            excel_dates,
            excel_times,
            hum_cols,
            tab_color="2F75B5",
            on_progress=on_progress,
            label="Humidity",
        )
    return output_path


def _zip_text(zf: ZipFile, name: str, text: str) -> None:
    info = ZipInfo(name)
    info.compress_type = ZIP_DEFLATED
    zf.writestr(info, text.encode("utf-8"))


def _stream_sheet(
    zf: ZipFile,
    zip_name: str,
    loggers: list[str],
    letters: list[str],
    last_ref: str,
    excel_dates: list[int],
    excel_times: list[float],
    columns: list[array],
    tab_color: str,
    on_progress,
    label: str,
) -> None:
    info = ZipInfo(zip_name)
    info.compress_type = ZIP_DEFLATED
    n_rows = len(excel_dates)
    last_col = 2 + len(loggers)

    with zf.open(info, "w") as raw:
        write = raw.write
        write(
            (
                f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"""
                f"""<worksheet xmlns="{_NS_MAIN}">"""
                f"""<sheetPr><tabColor rgb="FF{tab_color}"/></sheetPr>"""
                f"""<dimension ref="A1:{last_ref}"/>"""
                """<sheetViews><sheetView workbookViewId="0">"""
                """<pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/>"""
                """<selection pane="bottomRight" activeCell="C2" sqref="C2"/>"""
                """</sheetView></sheetViews>"""
                """<sheetFormatPr defaultRowHeight="15"/>"""
                f"""<cols><col min="1" max="1" width="14" customWidth="1"/>"""
                f"""<col min="2" max="2" width="10" customWidth="1"/>"""
                f"""<col min="3" max="{max(last_col, 3)}" width="9" customWidth="1"/></cols>"""
                """<sheetData>"""
            ).encode("utf-8")
        )

        header_cells = [
            f'<c r="A1" t="inlineStr" s="2"><is><t>DATE</t></is></c>',
            f'<c r="B1" t="inlineStr" s="2"><is><t>Time</t></is></c>',
        ]
        for i, name in enumerate(loggers, start=2):
            header_cells.append(
                f'<c r="{letters[i]}1" t="inlineStr" s="2"><is><t>{escape(name)}</t></is></c>'
            )
        write(f'<row r="1" ht="20" customHeight="1">{"".join(header_cells)}</row>'.encode("utf-8"))

        buf: list[str] = []
        for row_i, (excel_date, excel_time) in enumerate(zip(excel_dates, excel_times)):
            r = row_i + 2
            parts = [
                f'<row r="{r}">',
                f'<c r="A{r}" s="1"><v>{excel_date}</v></c>',
                f'<c r="B{r}" s="3"><v>{excel_time:.10g}</v></c>',
            ]
            for col_i, column in enumerate(columns):
                value = column[row_i]
                if value == value:
                    parts.append(f'<c r="{letters[col_i + 2]}{r}" s="4"><v>{value:.1f}</v></c>')
            parts.append("</row>")
            buf.append("".join(parts))
            if len(buf) >= 200:
                write("".join(buf).encode("utf-8"))
                buf.clear()
                if on_progress and row_i % 1000 == 0:
                    on_progress(f"Writing {label} sheet {row_i + 1}/{n_rows}...")

        if buf:
            write("".join(buf).encode("utf-8"))

        write(
            (
                f"""</sheetData>"""
                f"""<autoFilter ref="A1:{last_ref}"/>"""
                """</worksheet>"""
            ).encode("utf-8")
        )
        if on_progress:
            on_progress(f"Finished {label} sheet.")


def _col_name(idx: int) -> str:
    name = []
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        name.append(chr(65 + rem))
    return "".join(reversed(name))
