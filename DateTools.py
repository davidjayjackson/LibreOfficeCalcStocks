import datetime

def _col_to_name(idx0):
    """0-based column index -> Calc column letters (0=A, 25=Z, 26=AA, ...)"""
    name = ""
    n = idx0 + 1
    while n > 0:
        n, r = divmod(n - 1, 26)
        name = chr(65 + r) + name
    return name

def _get_python_date_from_cell(cell, calc_epoch):
    """
    Returns a datetime.date from a Calc cell that contains:
    - a numeric Calc date serial (preferred), OR
    - an ISO string YYYY-MM-DD (fallback)
    """
    if cell.Value != 0:
        return calc_epoch + datetime.timedelta(days=int(cell.Value))

    s = (cell.String or "").strip()
    if not s:
        return None

    try:
        return datetime.date.fromisoformat(s)
    except Exception:
        return None

def date_starts_insert_left():
    doc = XSCRIPTCONTEXT.getDocument()
    sheet = doc.CurrentController.ActiveSheet

    # Determine selected column (works for a cell or a range)
    sel = doc.CurrentController.Selection
    try:
        date_col = sel.RangeAddress.StartColumn
    except Exception:
        date_col = sel.CellAddress.Column

    # Insert 4 columns to the LEFT of the date column
    sheet.Columns.insertByIndex(date_col, 4)

    # After insertion, the original date column has shifted RIGHT by 4
    col_year    = date_col
    col_quarter = date_col + 1
    col_month   = date_col + 2
    col_week    = date_col + 3
    col_date    = date_col + 4

    year_L    = _col_to_name(col_year)
    quarter_L = _col_to_name(col_quarter)
    month_L   = _col_to_name(col_month)
    week_L    = _col_to_name(col_week)
    date_L    = _col_to_name(col_date)

    # Headers (row 1)
    sheet.getCellRangeByName(f"{year_L}1").String = "Year Start"
    sheet.getCellRangeByName(f"{quarter_L}1").String = "Quarter Start"
    sheet.getCellRangeByName(f"{month_L}1").String = "Month Start"
    sheet.getCellRangeByName(f"{week_L}1").String = "Week Start"

    calc_epoch = datetime.date(1899, 12, 30)

    # Date format YYYY-MM-DD
    nf = doc.NumberFormats
    locale = doc.CharLocale
    fmt_str = "YYYY-MM-DD"
    fmt_key = nf.queryKey(fmt_str, locale, True)
    if fmt_key == -1:
        fmt_key = nf.addNew(fmt_str, locale)

    # Fill down from row 2 until date column is blank
    row = 2
    while True:
        date_cell = sheet.getCellRangeByName(f"{date_L}{row}")
        d = _get_python_date_from_cell(date_cell, calc_epoch)
        if d is None:
            break

        week_start = d - datetime.timedelta(days=d.weekday())  # Monday-based
        month_start = datetime.date(d.year, d.month, 1)
        q_month = ((d.month - 1) // 3) * 3 + 1
        quarter_start = datetime.date(d.year, q_month, 1)
        year_start = datetime.date(d.year, 1, 1)

        # Write results (Calc serial dates)
        sheet.getCellRangeByName(f"{year_L}{row}").Value    = (year_start - calc_epoch).days
        sheet.getCellRangeByName(f"{quarter_L}{row}").Value = (quarter_start - calc_epoch).days
        sheet.getCellRangeByName(f"{month_L}{row}").Value   = (month_start - calc_epoch).days
        sheet.getCellRangeByName(f"{week_L}{row}").Value    = (week_start - calc_epoch).days

        # Apply date formatting
        sheet.getCellRangeByName(f"{year_L}{row}").NumberFormat = fmt_key
        sheet.getCellRangeByName(f"{quarter_L}{row}").NumberFormat = fmt_key
        sheet.getCellRangeByName(f"{month_L}{row}").NumberFormat = fmt_key
        sheet.getCellRangeByName(f"{week_L}{row}").NumberFormat = fmt_key
        date_cell.NumberFormat = fmt_key

        row += 1

g_exportedScripts = (date_starts_insert_left,)
