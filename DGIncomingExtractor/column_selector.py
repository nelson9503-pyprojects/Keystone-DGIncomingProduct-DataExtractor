from tkinter import ttk


class ColumnSelector:

    def __init__(self, master):
        self.main = ttk.LabelFrame(master, text="選擇提取欄位(Column)")
        self.setup_label()
        self.setup_combos()

    def setup_label(self):
        textlist = [
            "生產通知單",
            "工作單",
            "規格",
            "顏色",
            "物料名稱",
            "數量",
            "包裝",
            "總數量",
            "單位",
            "來貨單編號",
            "到港日期",
            "價值(HKD)"
        ]
        row = 0
        for text in textlist:
            lbl = ttk.Label(self.main, text=text)
            lbl.grid(row=row, column=0)
            row += 1

    def setup_combos(self):
        self.combos = []
        defaultCols = [
            "B", "D", "F", "I", "R",
            "J", "L", "N", "O", "Y",
            "Z", "U"
        ]
        vals = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L",
                "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]
        row = 0
        for col in defaultCols:
            combo = ttk.Combobox(self.main, values=vals, font=(18))
            combo.set(col)
            combo.grid(row=row, column=1)
            self.combos.append(combo)
            row += 1

    def get_config(self):
        items = [
            "memo",
            "workOrder",
            "spec",
            "color",
            "material",
            "amount",
            "packing",
            "totalAmount",
            "unit",
            "incoming",
            "date",
            "price"
        ]
        config = {}
        for i in range(len(items)):
            config[items[i]] = self.combos[i].get()
        return config
