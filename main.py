import os
import customtkinter as ctk
import openpyxl
from tkinter import filedialog, messagebox

def update_stock_list():
    try:
        # 1. 単価表ファイルを選択
        price_file = filedialog.askopenfilename(
            title="単価表ファイルを選択してください",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not price_file:
            return
        
        # 2. 在庫表ファイルを選択
        stock_file = filedialog.askopenfilename(
            title="在庫表ファイルを選択してください", 
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not stock_file:
            return
        
        # 3. 単価表から情報を読み取り
        price_wb = openpyxl.load_workbook(price_file)
        price_ws = price_wb.active
        
        # 単価情報を辞書に格納（商品コード: 単価）
        price_dict = {}
        for row in price_ws.iter_rows(min_row=2, values_only=True):  # 2行目から開始（1行目はヘッダー）
            if row[0] and row[1]:  # 商品コード（A列）と単価（B列）が存在する場合
                product_code = str(row[0])
                unit_price = row[1]
                price_dict[product_code] = unit_price
        
        # 4. 在庫表を開いて更新
        stock_wb = openpyxl.load_workbook(stock_file)
        stock_ws = stock_wb.active
        
        updated_count = 0
        for row_num, row in enumerate(stock_ws.iter_rows(min_row=2), start=2):
            product_code = str(row[0].value) if row[0].value else ""
            
            # 単価表に該当する商品コードがあれば更新
            if product_code in price_dict:
                # C列に単価を更新（列番号は適宜調整）
                stock_ws.cell(row=row_num, column=3).value = price_dict[product_code]
                updated_count += 1
        
        # 5. 在庫表を保存
        stock_wb.save(stock_file)
        
        # 6. 完了メッセージ
        messagebox.showinfo("完了", f"{updated_count}件の商品の単価を更新しました。")
        
    except Exception as e:
        messagebox.showerror("エラー", f"処理中にエラーが発生しました: {str(e)}")

ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")

app=ctk.CTk()
app.geometry("400x300")
app.title("CustomTkinter Example")

frame=ctk.CTkFrame(master=app)
frame.pack(pady=20, padx=20, fill="both", expand=True)

button_1=ctk.CTkButton(master=frame, text="在庫単価を在庫表に転記", command=update_stock_list)
button_1.pack(pady=10, padx=10)

app.mainloop()