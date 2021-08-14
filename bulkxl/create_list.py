
from win32com.client import client
import search_file


def execute(target_dir, exclusion_dir, sheet_name, work_file, header_record):
    target_list = search_file.execute(target_dir, exclusion_dir.split(','))
    app = client.Dispatch('Excel.Application')
    index_wb = app.Workbooks.Add()
    index_ws = index_wb.Worksheets(1)
    index_ws.name = sheet_name
    record_count = 1
    is_configured_header = False
    for excel_file_path in target_list:
        try:
            target_wb = app.add(excel_file_path)
            target_ws = target_wb.Worksheets(sheet_name)
            target_ws.Activate()
        except Exception as e:
            print(e)
            continue
        target_record_list = [
            cell.row for cell in work_sheet["A:A"] if cell.value is not None]
        a_max_row = max(target_record_list, key=lambda record: record)
        if not is_configured_header:
            for header in work_sheet.iter_rows(min_row=header_record, max_row=header_record):
                new_work_sheet.cell(record_count, 1).value = 'No'
                new_work_sheet.cell(record_count, 2).value = 'FilePath'
                copying_record(new_work_sheet,
                               record_count, header)
                record_count += 1
            is_configured_header = True
        for record in work_sheet.iter_rows(min_row=header_record+1, max_row=a_max_row,  max_col=work_sheet.max_column):
            new_work_sheet.cell(record_count, 1).value = record_count-1
            new_work_sheet.cell(record_count, 2).value = excel_file_path
            new_work_sheet.cell(record_count, 2).hyperlink = excel_file_path
            new_work_sheet.row_dimensions[record_count].height = work_sheet.row_dimensions[record[0].row].height
            copying_record(new_work_sheet, record_count, record)
            record_count += 1
    new_work_book.save(work_file)


def copying_record(new_work_sheet, record_count, record):
    for cell in record:
        new_work_sheet.cell(record_count, cell.column +
                            2).alignment = openpyxl.styles.Alignment(wrapText=True)
        new_work_sheet.cell(record_count, cell.column+2).value = cell.value
        new_work_sheet.cell(record_count, cell.column +
                            2).font = cell.font._StyleProxy__target
        new_work_sheet.cell(record_count, cell.column +
                            2).alignment = cell.alignment._StyleProxy__target
