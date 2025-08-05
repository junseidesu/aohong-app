import os
import customtkinter as ctk
import openpyxl as opx
from tkinter import filedialog, messagebox

PRICE_FILE_PATH="フロンガス単価表2025.xlsx"
STOCK_FILE_PATH="2025在庫.xlsx"
SALES_FILE_PATH="R706 得意先別売上分析表.xlsx"

def update_stock_list():
    price_list=opx.load_workbook(PRICE_FILE_PATH)
    stock_list=opx.load_workbook(STOCK_FILE_PATH)
    try:
        price_sheet=price_list["一般総平均"]
    except:
        messagebox.showerror("エラー", "一般総平均シートが見つかりません")
        return
    
    
class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.title("在庫単価転記アプリ")
        self.geometry("400x300")
        
        # フレームを自分自身（self）にアタッチ
        self.frame = ctk.CTkFrame(master=self)
        self.frame.pack(pady=20, padx=20, fill="both", expand=True)
        
        # ボタンをフレームに配置
        self.button_1 = ctk.CTkButton(master=self.frame, text="在庫単価を在庫表に転記", command=update_stock_list)
        self.button_1.pack(pady=10, padx=10)

# アプリのインスタンスを作成して実行
app = App()
app.mainloop()