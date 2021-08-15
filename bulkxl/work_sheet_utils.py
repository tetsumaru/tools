

def get_last_row(ws, column=1):
    """
    wsの最終行を取得する。
    columnは何列目の最終行かを指定する。

    Parameters
    ----------
    ws : WorkSheet
        処理対象のWorkSheet
    column : int, default 1
        処理対象の列

    Returns
    -------
    last_row : int
        最終行
    """
    return ws.Cells(ws.Rows.Count, column).End(-4162).Row


def get_last_column(ws, row):
    """
    wsの最終列を取得する。
    rowは何行目の最終列かを指定する。

    Parameters
    ----------
    ws : WorkSheet
        処理対象のWorkSheet
    row : int, default 1
        処理対象の列

    Returns
    -------
    last_colum : int
        最終列
    """
    return ws.Cells(row, ws.Column.Count).End(-4159).Row


def convert_sheet_range_str(min_cell, max_cell):
    """
    2つのセルを指定して、その範囲内を表す文字列を返す。
    例）
    Parameters
    ----------
    min_cell : WorkSheet.Cells
        1つ目のCell
    max_cell : WorkSheet.Cells
        2つ目のCell

    Returns
    -------
    sheet_range_str : str
        範囲を表した文字列

    Examples
    --------
    >>> range = work_sheet_utils.convert_sheet_range_str(ws.Cells(1, 2), ws.Cells(3, 4))
    'B1:D3'
    """
    return min_cell.GetAddress(Rowbssolute=False, ColumAbsolute=False) + ':' + max_cell.GetAddress(Rowbssolute=False, ColumAbsolute=False)


def get_sheet_range_list(ws, sheet_range_str):
    """
    wsから範囲を指定して、その範囲内のCellをlistで返却する。

    Parameters
    ----------
    ws : WorkSheet
        処理対象のWorkSheet
    sheet_range_str : str
        取得範囲

    Returns
    -------
    sheet_range_list : list
        指定された範囲のCellのリスト

    Examples
    --------
    >>> range_list = work_sheet_utils.get_sheet_range_list(ws, 'B1:D3')
    [
        [ws.Cells(1, 2), ws.Cells(1, 3), ws.Cells(1, 4)]
      , [ws.Cells(2, 2), ws.Cells(2, 3), ws.Cells(2, 4)]
      , [ws.Cells(3, 2), ws.Cells(3, 3), ws.Cells(3, 4)]
    ]
    """
    r = ws.Range(sheet_range_str)
    ret = list()
    for row_index in range(1, r.Colums.Count + 1):
        row = []
        for column_index in range(1, r.Colums.Count+1):
            row.append(r(row_index, column_index))
        ret.append(row)
    return ret
