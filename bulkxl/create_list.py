import openpyxl
import search_file


def execute(target_dir, exclusion_dir, sheet_name, work_file, header_record):
    list = search_file.execute(target_dir, exclusion_dir.split(','))
    new_work_book = openpyxl.Workbook()
    new_work_sheet = new_work_book["Sheet"]
    new_work_sheet.title = sheet_name
    record_count = 1
    is_configured_header = False
    for excel_file_path in list:
        work_book = openpyxl.load_workbook(excel_file_path)
        work_sheet = work_book[sheet_name]
        if not is_configured_header:
            for header in work_sheet.iter_rows(min_row=header_record, max_row=header_record):
                for cell in header:
                    new_work_sheet.cell(
                        record_count, cell.column+3).value = cell.value
                record_count += 1
            is_configured_header = True
        for record in work_sheet.iter_rows(min_row=header_record+1, max_col=work_sheet.max_column):
            for cell in record:
                new_work_sheet.cell(
                    record_count, cell.column+3).value = cell.value
            record_count += 1
    new_work_book.save(work_file)


target_dir = "C:/Users/tnaka/OneDrive/デスクトップ/hoge/"
exclusion_dir = '対象外,たいしょうがい'
sheet_name = 'hoge'
work_file = 'C:/Users/tnaka/OneDrive/デスクトップ/temp.xlsx'
header_record = 2
execute(target_dir, exclusion_dir, sheet_name, work_file, header_record)
