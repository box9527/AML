import tkinter as tk
from tkinter import filedialog, messagebox
from extract_text import PDFProcessor
import zipfile
import os
import tabula
from txtrank_v2 import TextRankSummarization
class GUIApp:
    def __init__(self, root): 
        self.root = root
        self.root.title("pdf檔選擇器")
        self.root.geometry("400x200")
        self.file_path = tk.StringVar()
        self.process_zip = False 
        
        # 創建 PDFProcessor 實例
        self.pdf_processor = PDFProcessor()
        
        # 選擇壓縮檔或資料夾按鈕
        self.button_zip = tk.Button(self.root, text="選擇zip檔", command=lambda: self.browse("file", True))
        self.button_zip.pack()
        self.button_folder = tk.Button(self.root, text="選擇資料夾", command=lambda: self.browse("folder", False))
        self.button_folder.pack()
        
        # 檔案相關的 Label
        self.file_label = tk.Label(self.root, text="檔案尚未選擇")
        self.file_label.pack()
        
        # 狀態相關的 Label
        self.status_label = tk.Label(self.root, text="等待選擇檔案")
        self.status_label.pack()
        
        # 開始按鈕
        self.confirm_button = tk.Button(self.root, text="開始", command=self.confirm_selection)
        self.confirm_button.pack()


    def browse(self, file_type, process_zip):
        if file_type == "file":
            file_path = filedialog.askopenfilename(
                initialdir="/",
                title="Select a file",
                filetypes=(("Zip files", "*.zip"), ("All files", "*.*"))
            )
        elif file_type == "folder":
            file_path = filedialog.askdirectory()

        if file_path:
            self.file_path.set(file_path)
            self.source_folder = file_path
            self.process_zip = process_zip  # 新增一個屬性用來記錄是否處理壓縮檔
            selection_type = "壓縮檔" if process_zip else "資料夾"
            self.file_label.config(text=f"選擇的{selection_type}路徑: {self.source_folder}")
            self.status_label.config(text="選擇完成後按下『開始』")


    def run_gui(self):
        self.root.mainloop()

    def confirm_selection(self):
        print("確定按鈕被按下了!")
        if not hasattr(self, 'source_folder') or not self.source_folder:
            messagebox.showwarning("警告", "請先選擇檔案。")
            return
        
        if self.process_zip:  # 如果是壓縮檔，呼叫 extract_file 方法
            extract_to_path = './pdfs'
            self.extract_file(extract_to_path)
            self.source_folder = extract_to_path
        else: # 如果是一般檔案，直接使用原始的 source_folder
             self.source_folder = self.source_folder
        
        trs = TextRankSummarization()
        trs.run_processing(self.source_folder)
        
        self.root.quit()
        self.root.destroy()  # 關閉主視窗

    def extract_file(self, extract_to_path):
        file_extension = os.path.splitext(self.source_folder)[1].lower()

        if file_extension == '.zip':
            with zipfile.ZipFile(self.source_folder, 'r') as zip_ref:
                for file_info in zip_ref.infolist():
                    try:
                        filename = file_info.filename.encode('cp437').decode('utf-8')
                        #print("檔案名稱：", filename)

                    except UnicodeDecodeError:
                        filename = file_info.filename.encode('cp437').decode('cp437')

                    if filename.lower().endswith('.pdf'):
                        zip_ref.extract(file_info, path=extract_to_path)
                        original_path = os.path.join(extract_to_path, file_info.filename)
                        new_path = os.path.join(extract_to_path, filename)
                        os.rename(original_path, new_path)
                        print("重新命名檔案：", new_path)
