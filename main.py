import os
import customtkinter as ctk
import openpyxl as opx
from tkinter import filedialog, messagebox
import datetime

class App(ctk.CTkToplevel):
    def __init__(self):
        super().__init__()
        self.title("設定")
        self.geometry("400x300")

        

        self.price_file_frame=ctk.CTkFrame(master=self)
        self.price_file_frame.pack(pady=10)
        self.price_file_button=ctk.CTkButton(
            master=self.price_file_frame,
            text="単価表ファイルを選択",
            
        )

    def select_file(self, file_type):
        file_path = filedialog.askopenfilename(filetypes=[("Excelファイル", "*.xlsx")])
        if file_path:
            if file_type == "price":
                self.PRICE_FILE_PATH = file_path
            elif file_type == "stock":
                self.STOCK_FILE_PATH = file_path
            elif file_type == "sales":
                self.SALES_FILE_PATH = file_path


class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        self.PRICE_FILE_PATH="フロンガス単価表2025.xlsx"
        self.STOCK_FILE_PATH="2025在庫.xlsx"
        self.SALES_FILE_PATH="R706 得意先別売上分析表.xlsx"

        self.file_frame=ctk.CTkFrame(master=self)
        self.file_frame.pack(padx=10, fill="x")

        self.price_file_button=ctk.CTkButton(
            master=self.file_frame,
            text="単価表"
        )

        self.id_row_in_price=3
        self.price_row_in_price=25
        self.id_column_in_stock=3
        self.price_column_in_stock=10
        self.data_start_row_in_stock=7
        self.id_column_in_sales=4
        self.profit_column_in_sales=9
        self.profit_rate_column_in_sales=10
        self.sales_column_in_sales=6
        self.sales_num_column_in_sales=8

        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")
        self.title("在庫単価転記アプリ")
        self.geometry("400x300")

        self.iconbitmap("reclaim_icon.ico")

        self.frame=ctk.CTkFrame(master=self)
        self.frame.pack(pady=20, padx=20, fill="both", expand=True)

        self.year_var = ctk.StringVar(value=f"{str(datetime.datetime.now().year)}年")
        self.month_var = ctk.StringVar(value=f"{str(datetime.datetime.now().month)}月")

        self.year_month_menu_frame=ctk.CTkFrame(master=self.frame)
        self.year_month_menu_frame.pack(pady=10, fill="x")

        current_year = datetime.datetime.now().year
        year_options=[f"{i}年" for i in range(current_year-1, current_year+2)]
        self.year_menu=ctk.CTkOptionMenu(master=self.year_month_menu_frame, variable=self.year_var, values=year_options)
        self.year_menu.grid(row=0, column=0, padx=10, pady=10)
        self.year_menu.set(f"{current_year}年")

        current_month = datetime.datetime.now().month
        month_options = [f"{i}月" for i in range(1, 13)]
        self.month_menu = ctk.CTkOptionMenu(master=self.year_month_menu_frame, variable=self.month_var, values=month_options)
        self.month_menu.grid(row=0, column=1, padx=10, pady=10)
        self.month_menu.set(f"{current_month}月")

        self.year_month_menu_frame.grid_columnconfigure(0, weight=1)
        self.year_month_menu_frame.grid_columnconfigure(1, weight=1)

        self.button_1=ctk.CTkButton(master=self.frame, text="在庫単価を在庫表に転記", command=self.update_stock_list)
        self.button_1.pack(pady=10, padx=10)

        self.button_2=ctk.CTkButton(master=self.frame, text="在庫単価を売上表に転記", command=self.update_sales_list)
        self.button_2.pack(pady=10, padx=10)

    def update_stock_list(self):
        #まずは月を取得→在庫単価のファイルからIDと在庫単価の辞書を作成
        price_list=opx.load_workbook(self.PRICE_FILE_PATH, data_only=True)
        stock_list=opx.load_workbook(self.STOCK_FILE_PATH, data_only=True)
        try:
            price_sheet=price_list["一般総平均"]
        except:
            messagebox.showerror("エラー", "一般総平均シートが見つかりません")
            return
        
        selected_year, selected_month= self.get_selected_year_month()

        selected_year_int=int(selected_year[:-1])
        selected_month_int=int(selected_month[:-1])

        converted_year_month=f"{selected_year_int}{selected_month_int:02}"
        month_cell=self.search_string_in_row(price_sheet, 1, converted_year_month)
        next_month_cell=self.search_string_in_row(price_sheet, 1, self.calculate_calendar(selected_year_int, selected_month_int, 1))
        
        if not month_cell:
            messagebox.showerror("エラー", f"{converted_year_month}のセルが見つかりません")
            return
        
        id_price_dict={}
        next_month_col=None
        if next_month_cell:
            next_month_col=next_month_cell.column
        
        else:
            for i in range(1,price_sheet.max_column+2):
                num_of_none=0
                for j in range(1, 26):
                    if price_sheet.cell(row=j, column=i).value is None:
                        num_of_none+=1
                if num_of_none==25:
                    next_month_col=i
                    break
        
        
        for i in range(month_cell.column, next_month_col):
                id_cell=price_sheet.cell(row=self.id_row_in_price, column=i)
                price_cell=price_sheet.cell(row=self.price_row_in_price, column=i)
                if id_cell.value and price_cell.value:
                    id_price_dict[id_cell.value]=price_cell.value
                
        #ここから在庫表に転記
        sheetnames=stock_list.sheetnames
        stock_sheet=None
        for sheetname in sheetnames:
            if converted_year_month in sheetname:
                stock_sheet=stock_list[sheetname]
                break
        
        if not stock_sheet:
            messagebox.showerror("エラー", f"{converted_year_month}の在庫シートが見つかりません")
            return

        fix_point_count=0
        for row_num in range(1, stock_sheet.max_row+1):
            id=stock_sheet.cell(row=row_num, column=self.id_column_in_stock).value
            price=id_price_dict.get(id, None)
            if price:
                stock_sheet.cell(row=row_num, column=self.price_column_in_stock, value=price)
                fix_point_count+=1
                print(f"ID: {id}, Price: {price} updated")



        try:
            stock_list.save(self.STOCK_FILE_PATH)
            messagebox.showinfo("成功", f"{fix_point_count}件の価格を更新しました")
        except PermissionError:
            messagebox.showerror("エラー", "在庫ファイルが開かれています。閉じてから再度実行してください。")
        except Exception as e:
            messagebox.showerror("エラー", f"予期しないエラーが発生しました: {e}")

    def update_sales_list(self):
        stock_list=opx.load_workbook(self.STOCK_FILE_PATH, data_only=True)
        sales_list=opx.load_workbook(self.SALES_FILE_PATH, data_only=True)

        selected_year, selected_month= self.get_selected_year_month()

        selected_year_int=int(selected_year[:-1])
        selected_month_int=int(selected_month[:-1])

        converted_year_month=f"{selected_year_int}{selected_month_int:02}"
        sheetnames=stock_list.sheetnames
        stock_sheet=None
        for sheetname in sheetnames:
            if converted_year_month in sheetname:
                stock_sheet=stock_list[sheetname]
                break
        
        if not stock_sheet:
            messagebox.showerror("エラー", f"{converted_year_month}の在庫シートが見つかりません")
            return

        sales_sheet=sales_list.active

        id_price_dict={}
        for row in range(self.data_start_row_in_stock, stock_sheet.max_row + 1):
            id=stock_sheet.cell(row=row, column=self.id_column_in_stock).value
            price=stock_sheet.cell(row=row, column=self.price_column_in_stock).value
            if id and price:
                id_price_dict[id]=price

        fix_point_count=0
        for row in range(1, sales_sheet.max_row+1):
            id=sales_sheet.cell(row=row, column=self.id_column_in_sales).value
            price=id_price_dict.get(id, None)
            if price:
                sales=float(sales_sheet.cell(row=row, column=self.sales_column_in_sales).value)
                sales_num=float(sales_sheet.cell(row=row, column=self.sales_num_column_in_sales).value)
                if sales and sales_num:
                    profit=sales-sales_num * price
                    profit_rate=sales/profit
                    sales_sheet.cell(row=row, column=self.profit_column_in_sales, value=profit)
                    sales_sheet.cell(row=row, column=self.profit_rate_column_in_sales, value=profit_rate)
                    fix_point_count+=1
                    print(f"ID: {id}, Price: {price} updated")

        try:
            sales_list.save(self.SALES_FILE_PATH)
            messagebox.showinfo("成功", f"{fix_point_count}件の売上データを更新しました")
        except PermissionError:
            messagebox.showerror("エラー", "売上ファイルが開かれています。閉じてから再度実行してください。")
        except Exception as e:
            messagebox.showerror("エラー", f"予期しないエラーが発生しました: {e}")
        


    def search_string_in_row(self, sheet, row, search_string):
        """
        指定されたExcelシートの特定の行から、指定された文字列を含むセルを検索する関数
        
        Args:
            sheet (openpyxl.worksheet): 検索対象のExcelワークシート
            row (int): 検索対象の行番号（1から始まる）
            search_string (str): 検索したい文字列
        
        Returns:
            openpyxl.cell.Cell or None: 
                - 検索文字列が見つかった場合：該当するセルオブジェクト
                - 見つからなかった場合：None
        
        処理の流れ:
            1. 指定された行のすべてのセルを順番にチェック
            2. 各セルに値が存在するかを確認
            3. セルの値に検索文字列が含まれているかを判定
            4. 最初に見つかったセルを返す
            5. 見つからない場合はNoneを返す
        
        使用例:
            # 1行目から"202501"という文字列を含むセルを検索
            cell = self.search_string_in_row(worksheet, 1, "202501")
            if cell:
                print(f"見つかりました: 列{cell.column}, 値:{cell.value}")
            else:
                print("見つかりませんでした")
        """
        target_cell=None
        for row in  sheet.iter_rows(min_row=row, max_row=row):
            for cell in row:
                if cell.value:
                    if search_string in cell.value:
                        target_cell=cell
                        break
            if target_cell:
                break
            
        return target_cell
    
    def get_selected_year_month(self):
        selected_year=self.year_var.get()
        selected_month=self.month_var.get()

        return selected_year, selected_month

    def calculate_calendar(self, year, month, gap):
        if month+gap<13 and month+gap>0:
            return f"{year}{month+gap:02}"
        elif month+gap>12:
            return f"{year+1}{month+gap-12:02}"
        else:
            return f"{year-1}{month+gap+12:02}"

if __name__ == "__main__":
    app = App()
    app.mainloop()