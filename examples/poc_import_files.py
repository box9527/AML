#!/usr/bin/env python3.10
# coding: utf-8
# @carl9527


import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import customtkinter
from poc_keyword_oriented import Productivity


class FileSelectorApp:
    def __init__(self, root):
        self.root = root
        self.root.geometry("420x200")
        self.root.minsize(width=420, height=200)
        self.root.resizable(width=False, height=False)

        self.font = customtkinter.CTkFont(family="DFKai-SB", size=16, weight="bold")
        self.root.title("檔案選擇器")

        #photo = tk.PhotoImage(file='poc_tool8.png')
        #self.root.wm_iconphoto(False, photo)

        self.file_path1 = tk.StringVar()
        self.file_path2 = tk.StringVar()

        # 設定標籤和按鈕
        label1 = customtkinter.CTkLabel(root, width=130, text="客戶金流 PDF:", font=self.font)
        label1.grid(row=0, column=0, pady=20)

        self.entry1 = customtkinter.CTkEntry(root, textvariable=self.file_path1, width=205, font=self.font)
        self.entry1.grid(row=0, column=1, pady=20)

        browse_button1 = customtkinter.CTkButton(root, text="瀏覽", width=60, font=self.font, command=lambda: self.browse_file(self.file_path1))
        browse_button1.grid(row=0, column=2, padx=10)

        label2 = customtkinter.CTkLabel(root, width=130, text="工具七參考檔:", font=self.font)
        label2.grid(row=1, column=0, pady=20)

        self.entry2 = customtkinter.CTkEntry(root, textvariable=self.file_path2, width=205, font=self.font)
        self.entry2.grid(row=1, column=1, pady=20)

        browse_button2 = customtkinter.CTkButton(root, text="瀏覽", width=60, font=self.font, command=lambda: self.browse_file(self.file_path2))
        browse_button2.grid(row=1, column=2, padx=10)

        confirm_button = customtkinter.CTkButton(root, text="確定", command=self.confirm_selection, font=self.font)
        confirm_button.grid(row=2, column=0, columnspan=3, pady=20)

    def browse_file(self, var):
        file_path = filedialog.askopenfilename()
        if file_path:
            if self.check_file_extension(file_path):
                var.set(file_path)
            else:
                messagebox.showwarning("警告", "檔案副檔名必須是pdf或xlsm。")
                var.set("")

    def check_file_extension(self, file_path):
        valid_extensions = ['.pdf', '.xlsm']
        return any(file_path.lower().endswith(ext) for ext in valid_extensions)

    def confirm_selection(self):
        source = self.file_path1.get()
        tool7 = self.file_path2.get()

        if (len(source) <= 0) or (len(tool7) <= 0):
            messagebox.showwarning("警告", "客戶金流資料與工具七檔案必須同時存在。")
            self.file_path1.set("")
            self.file_path2.set("")
            return
        # 處理選擇檔案的邏輯，這裡只是簡單地打印檔案路徑
        print("檔案1:", source)
        print("檔案2:", tool7)
        self.root.destroy()

        product = Productivity()
        bsuccess = product.output(source, tool7)
        if bsuccess == False:
            logger.critical('Keywords procedure is failed.')


if __name__ == "__main__":
    #root = tk.Tk()
    customtkinter.set_appearance_mode("dark")
    root = customtkinter.CTk()

    app = FileSelectorApp(root)
    root.mainloop()
