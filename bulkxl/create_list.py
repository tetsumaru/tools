import win32com.client
from pywintypes import com_error
import search_file
from pathlib import Path
import work_sheet_utils


def execute(target_dir, exclusion_file, sheet_name, work_file, header_row, extra_column_count=3):
    """
    作業用ファイルを作成する。
    作業用ファイルにはtarget_dirに指定された、ディレクトリ配下のExcelファイルのレコードの一覧が表示される。

    Parameters
    ----------
    target_dir : str
        処理対象のディレクトリのフルパス
    exclusion_name_list : list
        処理対象外になる文字列のlist
    sheet_name : str
        抽出対象のシート
    work_file : str
        作業用ファイルのフルパス
    header_row : int
        処理対象のファイルのヘッダーの行数
    extra_column_count : int ,default 3
        header_rowより右の列のCellを取得したい場合設定する
    """
    target_list = search_file.execute(target_dir, exclusion_file.split(','))
    app = win32com.client.Dispatch('Excel.Application')
    app.Visible = False
    app.DisplayAlerts = False
    index_wb = app.Workbooks.Add()
    index_ws = index_wb.Worksheets(1)
    index_ws.name = sheet_name
    record_count = 2
    is_configured_header = False
    for excel_file_path in target_list:
        try:
            target_wb = app.workbooks.Add(str(Path(excel_file_path)))
            target_ws = target_wb.Worksheets(sheet_name)
            target_ws.Activate()
        except com_error as e:
            print(e)
            continue
        target_last_column = work_sheet_utils.get_last_column(
            target_ws, header_row)
        target_last_row = work_sheet_utils.get_last_row(target_ws, 1)
        if not is_configured_header:
            # ヘッダーの設定
            header_range_str = work_sheet_utils.convert_range_str_from_int(
                target_ws, header_row, 1, header_row, target_last_column)
            index_ws.cells(1, 1).value = 'No'
            index_ws.cells(1, 2).value = 'FilePath'
            target_ws.Range(header_range_str).Copy(index_ws.Range('C1'))
            is_configured_header = True
        link_range = index_ws.Range('B{}'.format(record_count) + ':' +
                                    'B{}'.format(record_count + target_last_row - header_row))
        link_range.value = excel_file_path
        index_ws.Hyperlinks.Add(Anchor=link_range, Address=excel_file_path, ScreenTip=Path(
            excel_file_path).name, TextToDisplay=Path(excel_file_path).stem)
        range_str = work_sheet_utils.convert_range_str_from_int(
            target_ws, header_row + 1, 1, target_last_row, target_last_column + extra_column_count)
        target_ws.Range(range_str).Copy(
            index_ws.Range('C{}'.format(record_count)))
        record_count += target_last_row - header_row
        target_wb.Close(False)
    index_wb.SaveAs(str(Path(work_file)))
    index_wb.Close(False)
    app.Quit()


execute('C:/Users/tnaka/OneDrive/デスクトップ/hoge/', '対象外,taissyogai',
        'hoge', 'C:/Users/tnaka/OneDrive/デスクトップ/temp2.xlsx', 2)
