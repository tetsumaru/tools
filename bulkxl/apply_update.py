
import win32com.client
from pywintypes import com_error
from itertools import groupby
from pathlib import Path
import work_sheet_utils


def execute(sheet_name, work_file, header_row, extra_column_count=3):
    app = win32com.client.Dispatch('Excel.Application')
    app.Visible = False
    app.DisplayAlerts = False
    index_wb = app.Workbooks.Add(str(Path(work_file)))
    index_ws = index_wb.Worksheets(sheet_name)
    record_count = 2
    index_B_last_row = work_sheet_utils.get_last_row(index_ws, 2)
    index_last_column = work_sheet_utils.get_last_column(
        index_ws, header_row)
    target_file_list = [cell for cell in index_ws.Range(
        'B2:B{}'.format(index_B_last_row))]

    for key, record_list in groupby(target_file_list, key=lambda record: record[0].value):
        index_record_count = len(tuple(record_list))
        try:
            target_app = win32com.client.Dispatch('Excel.Application')
            target_app.Visible = False
            target_app.DisplayAlerts = False
            target_wb = target_app.workbooks.Add(str(Path(key)))
            target_ws = target_wb.Worksheets(sheet_name)
            target_ws.Activate()
            index_range_str = work_sheet_utils.convert_range_str_from_int(
                target_ws, record_count, 3, record_count + index_record_count-1, index_last_column + extra_column_count)
            print(index_range_str)
            index_ws.Range(index_range_str).Copy(
                target_ws.Range('A{}'.format(header_row+1)))
            print(key)
            target_wb.SaveAs(str(Path(key)))
            target_app.DisplayAlerts = True
        except com_error as e:
            print(e)
            continue
        finally:
            record_count += index_record_count
            target_wb.Close(False)
