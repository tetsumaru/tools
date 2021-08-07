class MenuForm:
    target_dir_text_box = None
    exclusion_dir_text_box = None
    target_sheet_text_box = None
    work_file_text_box = None
    header_record_text_box = None

    def __init__(self, target_dir_text_box, exclusion_dir_text_box, target_sheet_text_box, work_file_text_box, header_record_text_box):
        self.target_dir_text_box = target_dir_text_box
        self.exclusion_dir_text_box = exclusion_dir_text_box
        self.target_sheet_text_box = target_sheet_text_box
        self.work_file_text_box = work_file_text_box
        self.header_record_text_box = header_record_text_box
