def get_shape(wb):
    """
    Returns dictionary indicating row and column ranges for each worksheet.
    """
    sheets = wb.sheetnames
    dims = {}
    for sheet_name in sheets:
        sheet = wb[sheet_name]
        column_list = [cell.column for cell in sheet[1]]
        # WARNING: this fails on MergeCells
        column_letters = [cell.column_letter for cell in sheet[1]]
        first_letter = column_letters[0]
        row_list = [cell.row for cell in sheet[first_letter]]
        dims[sheet_name] = {
            "sheet_name": sheet_name,
            "cols": [column_list[0], column_list[len(column_list) - 1]],
            "col_letters": [column_letters[0], column_letters[len(column_letters) - 1]],
            "rows": [row_list[0], row_list[len(row_list) - 1]],
        }
    return dims
