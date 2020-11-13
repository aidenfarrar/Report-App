import pandas as pd
import numpy as np
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.cell.cell import Cell
from openpyxl.styles import Font
import datetime


def format_line(line, t, ws):
    size_dict = {'header': 16, 'title': 14, 'table': 12}
    size = size_dict[t]
    for c in line:
        c = Cell(ws, value=c)
        c.font = Font(bold=True, size=size)
        yield c


def remove_excess(d):
    d = d[(d.T != 0).any()]
    d = d.loc[:, (d != 0).any(axis=0)]
    return d


def generate_morning_report(date):
    if type(date) == datetime.date:
        date = date.strftime('%Y-%m-%d')

    locations = []
    try:
        cvg_file = pd.read_excel(
            r'G:\Line Reports\Reports\CVG Daily Events\CVG {} Line Maintenance Report.xlsx'.format(date),
            sheet_name=None)
        locations.append([cvg_file, 'CVG'])
    except FileNotFoundError:
        print('No CVG')
    try:
        iln_file = pd.read_excel(
            r'G:\Line Reports\Reports\ILN Daily Events\ILN {} Line Maintenance Report.xlsx'.format(date),
            sheet_name=None)
        locations.append([iln_file, 'ILN'])
    except FileNotFoundError:
        print('No ILN')
    try:
        mia_file = pd.read_excel(
            r'G:\Line Reports\Reports\MIA Daily Events\MIA {} Line Maintenance Report.xlsx'.format(date),
            sheet_name=None)
        locations.append([mia_file, 'MIA'])
    except FileNotFoundError:
        print('No MIA')

    if not locations:
        return
    order = ['EE Status', 'Delays Cancellations', 'Aircraft in Work', 'Training', 'Events']
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Morning Report'
    date_cell = Cell(ws, value='Morning Report: ' + date)
    date_cell.font = Font(bold=True, size=14)
    ws.append([date_cell])
    ws.append([])
    cols = ['A', 'B', 'C', 'D', 'E']
    widths = [21, 12, 20, 24, 24]
    for df_dict, loc in locations:
        ws.append(format_line([loc + ' Line Maintenance Operations'], 'header', ws))
        for sheet in order:
            df = df_dict[sheet]
            ws.append(format_line([sheet], 'title', ws))
            if sheet == 'Events':
                df = df.set_index('Customer')
                df = remove_excess(df)
                df = df.replace(0, np.nan)
                df = df.reset_index()
            ws.append(format_line(list(df), 'table', ws))
            for row in dataframe_to_rows(df, index=False, header=False):
                ws.append(row)
            ws.append([])
            for col, width in zip(cols, widths):
                ws.column_dimensions[col].width = width
    wb.save(r'G:\Line Reports\Reports\Morning Reports\{} Morning Report.xlsx'.format(date))
