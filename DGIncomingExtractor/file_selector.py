from tkinter import ttk
from tkinter import filedialog
import openpyxl


class FileSelector:

    def __init__(self, master):
        self.main = ttk.LabelFrame(master, text="提取 Excel 文件")

        self.path = ""
        self.btn_selectfile = ttk.Button(
            self.main, text="選擇文件", command=self.select_file)
        self.btn_selectfile.grid(row=0, column=0)
        self.lbl_selectfile = ttk.Label(self.main)
        self.lbl_selectfile.grid(row=0, column=1)

        lbl = ttk.Label(self.main, text="選擇工作頁:")
        lbl.grid(row=1, column=0)
        self.combo_sheet = ttk.Combobox(self.main, state="readonly", font=(18))
        self.combo_sheet.grid(row=1, column=1)

        lbl = ttk.Label(self.main, text="選擇開始行數:")
        lbl.grid(row=2, column=0)
        self.entry_row = ttk.Entry(self.main, font=(18))
        self.entry_row.insert(0, 9)
        self.entry_row.grid(row=2, column=1)

    def select_file(self):
        path = filedialog.askopenfilename(
            title="Open Excel File",
            filetypes=(("Excel Files", "*.xlsx"),)
        )
        if not path in [None, ""]:
            self.path = path
            self.lbl_selectfile.config(text=path)
            wb = openpyxl.open(path, read_only=True)
            shs = wb.sheetnames
            self.combo_sheet.config(values=shs)
            self.combo_sheet.current(0)

    def get_config(self) -> dict:
        config = {}
        config["path"] = self.path
        config["sheet"] = self.combo_sheet.get()
        config["startrow"] = self.entry_row.get()
        return config
