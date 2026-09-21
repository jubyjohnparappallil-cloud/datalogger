"""Write the Temperature / Humidity Excel workbook in the warehouse format."""

from __future__ import annotations

import math
from array import array
from datetime import datetime, time as dtime
from pathlib import Path

import xlsxwriter


def write_workbook(
    output_path: Path,
    loggers: list[str],
    timestamps: list[datetime],
    temps: dict[str, array],
    hums: dict[str, array],
) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook = xlsxwriter.Workbook(str(output_path), {"constant_memory": True})

    header_fmt = workbook.add_format(
        {
            "bold": True,
            "font_name": "Calibri",
            "font_size": 11,
            "font_color": "white",
            "bg_color": "#305496",
            "align": "center",
            "valign": "vcenter",
            "border": 1,
        }
    )
    date_fmt = workbook.add_format(
        {
            "num_format": "DD/MM/YYYY",
            "font_name": "Calibri",
            "font_size": 11,
            "font_color": "#C00000",
            "align": "center",
            "border": 1,
        }
    )
    time_fmt = workbook.add_format(
        {
            "num_format": "HH:MM",
            "font_name": "Calibri",
            "font_size": 11,
            "font_color": "#C00000",
            "align": "center",
            "border": 1,
        }
    )
    num_fmt = workbook.add_format(
        {
            "num_format": "0.0",
            "font_name": "Calibri",
            "font_size": 11,
            "font_color": "#C00000",
            "align": "center",
            "border": 1,
        }
    )

    _write_sheet(workbook, "Temperature", loggers, timestamps, temps, header_fmt, date_fmt, time_fmt, num_fmt)
    _write_sheet(workbook, "Humidity", loggers, timestamps, hums, header_fmt, date_fmt, time_fmt, num_fmt)
    workbook.close()
    return output_path


def _write_sheet(workbook, name, loggers, timestamps, values, header_fmt, date_fmt, time_fmt, num_fmt):
    sheet = workbook.add_worksheet(name)
    sheet.freeze_panes(1, 2)
    sheet.set_row(0, 20)
    sheet.set_column(0, 0, 14)
    sheet.set_column(1, 1, 10)
    sheet.set_column(2, 2 + max(len(loggers) - 1, 0), 9)

    headers = ["DATE", "Time", *loggers]
    for col, header in enumerate(headers):
        sheet.write(0, col, header, header_fmt)

    columns = [values[logger] for logger in loggers]
    for row_i, ts in enumerate(timestamps, start=1):
        sheet.write_datetime(row_i, 0, datetime.combine(ts.date(), dtime.min), date_fmt)
        sheet.write_datetime(row_i, 1, dtime(ts.hour, ts.minute, 0), time_fmt)
        index = row_i - 1
        for col_i, column in enumerate(columns, start=2):
            value = column[index]
            if math.isnan(value):
                sheet.write_blank(row_i, col_i, None, num_fmt)
            else:
                sheet.write_number(row_i, col_i, round(value, 1), num_fmt)

    last_col = 1 + len(loggers)
    last_row = len(timestamps)
    if last_row:
        sheet.autofilter(0, 0, last_row, last_col)
    sheet.set_tab_color("#548235" if name == "Temperature" else "#2F75B5")
