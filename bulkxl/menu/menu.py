import tkinter 

def createGui():
    root = tkinter.Tk()
    root.title("bulkxl")
    root.geometry("400x250")

    target_dir_rable = tkinter.Label(root, text="対象ディレクトリ(フルパス)")
    target_dir_rable.place(x=5,y=0) 
    target_dir_text_box = tkinter.Entry()
    target_dir_text_box.configure(state='normal', width=50)
    target_dir_text_box.place(x=5,y=20) 

    exclusion_dir_rable = tkinter.Label(root, text="除外対象ディレクトリ名(「,」区切り)")
    exclusion_dir_rable.place(x=5,y=50) 
    exclusion_dir_text_box = tkinter.Entry()
    exclusion_dir_text_box.configure(state='normal', width=50)
    exclusion_dir_text_box.place(x=5,y=70) 

    target_sheet_rabel = tkinter.Label(root, text="取得するシート名")
    target_sheet_rabel.place(x=5,y=100) 
    target_sheet_text_box = tkinter.Entry()
    target_sheet_text_box.configure(state='normal', width=50)
    target_sheet_text_box.place(x=5,y=120) 

    work_file_rabel = tkinter.Label(root, text="作業用ファイル(フルパス)")
    work_file_rabel.place(x=5,y=150) 
    work_file_text_box = tkinter.Entry()
    work_file_text_box.configure(state='normal', width=50)
    work_file_text_box.insert(tkinter.END, u'C:\\temp.xlsx') 
    work_file_text_box.place(x=5,y=170)

    get_button = tkinter.Button(text='取得', width=10)
    get_button.place(x=5, y=200)

    update_button = tkinter.Button(text='更新', width=10)
    update_button.place(x=105, y=200)

    root.mainloop()

createGui()