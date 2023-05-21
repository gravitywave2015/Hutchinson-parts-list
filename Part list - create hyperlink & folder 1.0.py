from tkinter import *
from tkinter.filedialog import askopenfilename
import openpyxl
import os 

def create_link_folder(fn, path = "H:\99-CFT\D&R"):
    wb = openpyxl.load_workbook(fn)
    for sheet in wb.sheetnames:
        worksheet = wb[sheet]
        if worksheet.sheet_state == 'hidden':
            continue
        for i in range(4,19):
            if type(worksheet[i][4].value) == str:
                folder_path = "{}\{}".format(path,worksheet[i][4].value)
                if worksheet[i][13].hyperlink:
                    worksheet[i][13].hyperlink = None
                worksheet[i][13].value = '=HYPERLINK("{}")'.format(folder_path)
                if not os.path.exists(folder_path):
                    os.makedirs(folder_path)
    fn_new = fn[:-5] + '-update.xlsx'
    wb.save(fn_new)
    wb.close()
    os.startfile(fn_new)
    print('Done')

def main_tk():
    window = Tk()
    window.title('创建文件夹和超链接')
    window.geometry('500x100')

    file_path = ''

    def get_file():
        global file_path
        file_path = askopenfilename()  
        print(file_path)
        create_link_folder(file_path, path = "H:\99-CFT\D&R")

    button1 = Button(window, text='选择文件并打开更新的副本', command=get_file)
    button1.pack(pady = 20)

    window.mainloop()

if __name__ == '__main__':
    main_tk()