import tkinter
import create_list


def createGui():
    root = tkinter.Tk()
    root.title("bulkxl")
    root.geometry("400x300")

    target_dir_rable = tkinter.Label(root, text="対象ディレクトリ(フルパス)")
    target_dir_rable.place(x=5, y=0)
    target_dir_text_box = tkinter.Entry()
    target_dir_text_box.configure(state='normal', width=50)
    target_dir_text_box.place(x=5, y=20)

    exclusion_dir_rable = tkinter.Label(root, text="除外対象ディレクトリ名(「,」区切り)")
    exclusion_dir_rable.place(x=5, y=50)
    exclusion_dir_text_box = tkinter.Entry()
    exclusion_dir_text_box.configure(state='normal', width=50)
    exclusion_dir_text_box.place(x=5, y=70)

    target_sheet_rabel = tkinter.Label(root, text="取得するシート名")
    target_sheet_rabel.place(x=5, y=100)
    target_sheet_text_box = tkinter.Entry()
    target_sheet_text_box.configure(state='normal', width=50)
    target_sheet_text_box.place(x=5, y=120)

    work_file_rabel = tkinter.Label(root, text="作業用ファイル(フルパス)")
    work_file_rabel.place(x=5, y=150)
    work_file_text_box = tkinter.Entry()
    work_file_text_box.configure(state='normal', width=50)
    work_file_text_box.insert(tkinter.END, u'C:\\temp.xlsx')
    work_file_text_box.place(x=5, y=170)

    header_record_rabel = tkinter.Label(root, text="ヘッダー行")
    header_record_rabel.place(x=5, y=200)
    header_record_text_box = tkinter.Entry()
    header_record_text_box.configure(state='normal', width=10)
    header_record_text_box.insert(tkinter.END, u'0')
    header_record_text_box.place(x=5, y=220)

    get_button = tkinter.Button(text='取得', width=10, command=lambda: create_list.execute(
        target_dir_text_box.get(), exclusion_dir_text_box.get(), target_sheet_text_box.get(), work_file_text_box.get(), header_record_text_box.get()))
    get_button.place(x=5, y=250)

    update_button = tkinter.Button(text='更新', width=10)
    update_button.place(x=105, y=250)

    root.mainloop()


createGui()
