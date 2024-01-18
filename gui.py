import tkinter as tk
from tkinter import filedialog, messagebox

class GUIApp:
    def __init__(self, root):
        self.root = root
        self.root.title("選擇pdf檔案")
        self.root.geometry("400x200")
        self.file_path = tk.StringVar()

        self.button = tk.Button(self.root, text="選擇檔案", command=self.browse_file)
        self.button.pack(pady=20)  # 上、下方添加 20 像素的內部填充

        self.file_label = tk.Label(self.root, text="檔案尚未選擇")
        self.file_label.pack()

        self.status_label = tk.Label(self.root, text="等待選擇檔案")
        self.status_label.pack()

        self.confirm_button = tk.Button(self.root, text="開始", command=self.confirm_selection)
        self.confirm_button.pack(pady=20)

    def browse_file(self):
        file_path = filedialog.askopenfilename()
        if file_path:
            self.file_path.set(file_path)
            self.source_file = file_path
            self.file_label.config(text=f"選擇的檔案路徑: {self.source_file}")
            self.status_label.config(text="選擇完成後按下『開始』")

    def run_gui(self):
        self.root.mainloop()
    
    def confirm_selection(self):
        print("確定按鈕被按下了！")
        if not self.source_file:
            messagebox.showwarning("警告", "請先選擇檔案。")
            return
        self.root.destroy()
