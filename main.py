import os
import json
import customtkinter as ctk
import openpyxl as opx
from tkinter import filedialog, messagebox
import datetime

class Settings:
    # デフォルト設定
    PRICE_FILE_PATH = "フロンガス単価表2025.xlsx"
    STOCK_FILE_PATH = "2025在庫.xlsx"
    SALES_FILE_PATH = "R706 得意先別売上分析表.xlsx"
    
    ID_ROW_IN_PRICE = 3
    PRICE_ROW_IN_PRICE = 25
    ID_COLUMN_IN_STOCK = 3
    PRICE_COLUMN_IN_STOCK = 10
    DATA_START_ROW_IN_STOCK = 7
    ID_COLUMN_IN_SALES = 4
    PROFIT_COLUMN_IN_SALES = 9
    PROFIT_RATE_COLUMN_IN_SALES = 10
    SALES_COLUMN_IN_SALES = 6
    SALES_NUM_COLUMN_IN_SALES = 8
    
    APP_NAME = "在庫単価転記アプリ"
    WINDOW_WIDTH = 400
    WINDOW_HEIGHT = 300
    THEME = "dark"
    COLOR_THEME = "blue"
    ICON_FILE = "reclaim_icon.ico"
    SETTINGS_FILE = "settings.json"

    _DEFAULT_VALUES = {
        "PRICE_FILE_PATH": "フロンガス単価表2025.xlsx",
        "STOCK_FILE_PATH": "2025在庫.xlsx",
        "SALES_FILE_PATH": "R706 得意先別売上分析表.xlsx",
        "ID_ROW_IN_PRICE": 3,
        "PRICE_ROW_IN_PRICE": 25,
        "ID_COLUMN_IN_STOCK": 3,
        "PRICE_COLUMN_IN_STOCK": 10,
        "DATA_START_ROW_IN_STOCK": 7,
        "ID_COLUMN_IN_SALES": 4,
        "PROFIT_COLUMN_IN_SALES": 9,
        "PROFIT_RATE_COLUMN_IN_SALES": 10,
        "SALES_COLUMN_IN_SALES": 6,
        "SALES_NUM_COLUMN_IN_SALES": 8,
        "APP_NAME": "在庫単価転記アプリ",
        "WINDOW_WIDTH": 400,
        "WINDOW_HEIGHT": 300,
        "THEME": "dark",
        "COLOR_THEME": "blue",
        "ICON_FILE": "reclaim_icon.ico"
    }
    
    @classmethod
    def load_settings(cls):
        """settings.jsonから設定を読み込み"""
        try:
            if os.path.exists(cls.SETTINGS_FILE):
                with open(cls.SETTINGS_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                
                # ファイルパス設定
                files = data.get("files", {})
                cls.PRICE_FILE_PATH = files.get("price_file_path", cls.PRICE_FILE_PATH)
                cls.STOCK_FILE_PATH = files.get("stock_file_path", cls.STOCK_FILE_PATH)
                cls.SALES_FILE_PATH = files.get("sales_file_path", cls.SALES_FILE_PATH)
                
                # 行・列設定
                positions = data.get("positions", {})
                cls.ID_ROW_IN_PRICE = positions.get("id_row_in_price", cls.ID_ROW_IN_PRICE)
                cls.PRICE_ROW_IN_PRICE = positions.get("price_row_in_price", cls.PRICE_ROW_IN_PRICE)
                cls.ID_COLUMN_IN_STOCK = positions.get("id_column_in_stock", cls.ID_COLUMN_IN_STOCK)
                cls.PRICE_COLUMN_IN_STOCK = positions.get("price_column_in_stock", cls.PRICE_COLUMN_IN_STOCK)
                cls.DATA_START_ROW_IN_STOCK = positions.get("data_start_row_in_stock", cls.DATA_START_ROW_IN_STOCK)
                cls.ID_COLUMN_IN_SALES = positions.get("id_column_in_sales", cls.ID_COLUMN_IN_SALES)
                cls.PROFIT_COLUMN_IN_SALES = positions.get("profit_column_in_sales", cls.PROFIT_COLUMN_IN_SALES)
                cls.PROFIT_RATE_COLUMN_IN_SALES = positions.get("profit_rate_column_in_sales", cls.PROFIT_RATE_COLUMN_IN_SALES)
                cls.SALES_COLUMN_IN_SALES = positions.get("sales_column_in_sales", cls.SALES_COLUMN_IN_SALES)
                cls.SALES_NUM_COLUMN_IN_SALES = positions.get("sales_num_column_in_sales", cls.SALES_NUM_COLUMN_IN_SALES)
                
                print("設定を読み込みました")
        except Exception as e:
            print(f"設定の読み込みエラー: {e}")
    
    @classmethod
    def save_settings(cls):
        """現在の設定をsettings.jsonに保存"""
        try:
            settings_data = {
                "files": {
                    "price_file_path": cls.PRICE_FILE_PATH,
                    "stock_file_path": cls.STOCK_FILE_PATH,
                    "sales_file_path": cls.SALES_FILE_PATH
                },
                "positions": {
                    "id_row_in_price": cls.ID_ROW_IN_PRICE,
                    "price_row_in_price": cls.PRICE_ROW_IN_PRICE,
                    "id_column_in_stock": cls.ID_COLUMN_IN_STOCK,
                    "price_column_in_stock": cls.PRICE_COLUMN_IN_STOCK,
                    "data_start_row_in_stock": cls.DATA_START_ROW_IN_STOCK,
                    "id_column_in_sales": cls.ID_COLUMN_IN_SALES,
                    "profit_column_in_sales": cls.PROFIT_COLUMN_IN_SALES,
                    "profit_rate_column_in_sales": cls.PROFIT_RATE_COLUMN_IN_SALES,
                    "sales_column_in_sales": cls.SALES_COLUMN_IN_SALES,
                    "sales_num_column_in_sales": cls.SALES_NUM_COLUMN_IN_SALES
                }
            }
            
            with open(cls.SETTINGS_FILE, "w", encoding="utf-8") as f:
                json.dump(settings_data, f, ensure_ascii=False, indent=2)
            
            print("設定を保存しました")
            return True
        except Exception as e:
            print(f"設定の保存エラー: {e}")
            return False
        
    @classmethod
    def reset_to_defaults(cls):
        """設定をデフォルト値にリセット"""
        try:
            # デフォルト値をクラス変数に設定
            for key, value in cls._DEFAULT_VALUES.items():
                setattr(cls, key, value)
            
            print("設定をデフォルト値にリセットしました")
            return True
        except Exception as e:
            print(f"設定リセットエラー: {e}")
            return False
    
    @classmethod
    def get_default_value(cls, key):
        """指定されたキーのデフォルト値を取得"""
        return cls._DEFAULT_VALUES.get(key, None)
    
    @classmethod
    def is_default_value(cls, key):
        """現在の値がデフォルト値かどうかを確認"""
        current_value = getattr(cls, key, None)
        default_value = cls.get_default_value(key)
        return current_value == default_value

# アプリ起動時に設定を読み込み
Settings.load_settings()

class SettingsWindow(ctk.CTkToplevel):
    def __init__(self):
        super().__init__()
        self.title("設定")
        self.geometry("500x400")
        
        try:
            self.iconbitmap(Settings.ICON_FILE)
        except:
            pass

        # ファイル選択フレーム
        self.file_frame = ctk.CTkFrame(master=self)
        self.file_frame.pack(pady=10, padx=10, fill="x")

        ctk.CTkLabel(self.file_frame, text="ファイル設定", font=("Arial", 16, "bold")).pack(pady=10)

        # 価格ファイル選択
        self.price_file_button = ctk.CTkButton(
            master=self.file_frame,
            text="単価表ファイルを選択",
            command=lambda: self.select_file("price")
        )
        self.price_file_button.pack(pady=5)

        # 在庫ファイル選択
        self.stock_file_button = ctk.CTkButton(
            master=self.file_frame,
            text="在庫ファイルを選択",
            command=lambda: self.select_file("stock")
        )
        self.stock_file_button.pack(pady=5)

        # 売上ファイル選択
        self.sales_file_button = ctk.CTkButton(
            master=self.file_frame,
            text="売上ファイルを選択",
            command=lambda: self.select_file("sales")
        )
        self.sales_file_button.pack(pady=5)

        # 現在の設定表示
        self.info_frame = ctk.CTkFrame(master=self)
        self.info_frame.pack(pady=10, padx=10, fill="both", expand=True)

        ctk.CTkLabel(self.info_frame, text="現在の設定", font=("Arial", 14, "bold")).pack(pady=5)

        self.price_label = ctk.CTkLabel(self.info_frame, text=f"単価ファイル: {os.path.basename(Settings.PRICE_FILE_PATH)}")
        self.price_label.pack(pady=2)

        self.stock_label = ctk.CTkLabel(self.info_frame, text=f"在庫ファイル: {os.path.basename(Settings.STOCK_FILE_PATH)}")
        self.stock_label.pack(pady=2)

        self.sales_label = ctk.CTkLabel(self.info_frame, text=f"売上ファイル: {os.path.basename(Settings.SALES_FILE_PATH)}")
        self.sales_label.pack(pady=2)

        # 保存ボタン
        self.save_button = ctk.CTkButton(
            master=self,
            text="設定を保存",
            command=self.save_settings
        )
        self.save_button.pack(pady=10)

        self.reset_button = ctk.CTkButton(
            master=self,
            text="設定をリセット",
            command=self.reset_settings
        )
        self.reset_button.pack(pady=10)

    def select_file(self, file_type):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excelファイル", "*.xlsx"), ("すべてのファイル", "*.*")]
        )
        if file_path:
            if file_type == "price":
                Settings.PRICE_FILE_PATH = file_path  # クラス変数を直接変更
                self.price_label.configure(text=f"単価表: {os.path.basename(file_path)}")
                messagebox.showinfo("ファイル選択", f"単価表ファイルを選択しました:\n{os.path.basename(file_path)}")
            elif file_type == "stock":
                Settings.STOCK_FILE_PATH = file_path  # クラス変数を直接変更
                self.stock_label.configure(text=f"在庫表: {os.path.basename(file_path)}")
                messagebox.showinfo("ファイル選択", f"在庫ファイルを選択しました:\n{os.path.basename(file_path)}")
            elif file_type == "sales":
                Settings.SALES_FILE_PATH = file_path  # クラス変数を直接変更
                self.sales_label.configure(text=f"売上表: {os.path.basename(file_path)}")
                messagebox.showinfo("ファイル選択", f"売上ファイルを選択しました:\n{os.path.basename(file_path)}")

    def save_settings(self):
        if Settings.save_settings():
            messagebox.showinfo("設定保存", "設定を保存しました")
    
    def reset_settings(self):
        Settings.reset_to_defaults()
        Settings.save_settings()
        messagebox.showinfo("設定リセット", "設定をリセットしました")

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # UI設定を Settings から直接取得
        ctk.set_appearance_mode(Settings.THEME)
        ctk.set_default_color_theme(Settings.COLOR_THEME)
        self.title(Settings.APP_NAME)
        self.geometry(f"{Settings.WINDOW_WIDTH}x{Settings.WINDOW_HEIGHT}")

        try:
            self.iconbitmap(Settings.ICON_FILE)
        except:
            pass

        # 設定ボタンフレーム
        self.file_frame = ctk.CTkFrame(master=self)
        self.file_frame.pack(padx=10, pady=5, fill="x")

        self.settings_button = ctk.CTkButton(
            master=self.file_frame,
            text="設定",
            command=self.open_settings
        )
        self.settings_button.pack(pady=5)

        # メインフレーム
        self.frame = ctk.CTkFrame(master=self)
        self.frame.pack(pady=20, padx=20, fill="both", expand=True)

        # 年月選択
        self.year_var = ctk.StringVar(value=f"{str(datetime.datetime.now().year)}年")
        self.month_var = ctk.StringVar(value=f"{str(datetime.datetime.now().month)}月")

        self.year_month_menu_frame = ctk.CTkFrame(master=self.frame)
        self.year_month_menu_frame.pack(pady=10, fill="x")

        current_year = datetime.datetime.now().year
        year_options = [f"{i}年" for i in range(current_year-1, current_year+2)]
        self.year_menu = ctk.CTkOptionMenu(master=self.year_month_menu_frame, variable=self.year_var, values=year_options)
        self.year_menu.grid(row=0, column=0, padx=10, pady=10)
        self.year_menu.set(f"{current_year}年")

        current_month = datetime.datetime.now().month
        month_options = [f"{i}月" for i in range(1, 13)]
        self.month_menu = ctk.CTkOptionMenu(master=self.year_month_menu_frame, variable=self.month_var, values=month_options)
        self.month_menu.grid(row=0, column=1, padx=10, pady=10)
        self.month_menu.set(f"{current_month}月")

        self.year_month_menu_frame.grid_columnconfigure(0, weight=1)
        self.year_month_menu_frame.grid_columnconfigure(1, weight=1)

        self.button_1 = ctk.CTkButton(master=self.frame, text="在庫単価を在庫表に転記", command=self.update_stock_list)
        self.button_1.pack(pady=10, padx=10)

        self.button_2 = ctk.CTkButton(master=self.frame, text="在庫単価を売上表に転記", command=self.update_sales_list)
        self.button_2.pack(pady=10, padx=10)

    def open_settings(self):
        """設定ウィンドウを開く"""
        SettingsWindow()

    def update_stock_list(self):
        # Settings から直接ファイルパスを取得
        try:
            price_list = opx.load_workbook(Settings.PRICE_FILE_PATH, data_only=True)
            stock_list = opx.load_workbook(Settings.STOCK_FILE_PATH, data_only=True)
        except FileNotFoundError as e:
            messagebox.showerror("エラー", f"ファイルが見つかりません: {e}")
            return
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの読み込みに失敗しました: {e}")
            return

        try:
            price_sheet = price_list["一般総平均"]
        except KeyError:
            messagebox.showerror("エラー", "一般総平均シートが見つかりません")
            return
        
        selected_year, selected_month = self.get_selected_year_month()
        selected_year_int = int(selected_year[:-1])
        selected_month_int = int(selected_month[:-1])
        converted_year_month = f"{selected_year_int}{selected_month_int:02}"
        
        month_cell = self.search_string_in_row(price_sheet, 1, converted_year_month)
        next_month_cell = self.search_string_in_row(price_sheet, 1, self.calculate_calendar(selected_year_int, selected_month_int, 1))
        
        if not month_cell:
            messagebox.showerror("エラー", f"{converted_year_month}のセルが見つかりません")
            return
        
        id_price_dict = {}
        next_month_col = None
        if next_month_cell:
            next_month_col = next_month_cell.column
        else:
            for i in range(1, price_sheet.max_column+2):
                num_of_none = 0
                for j in range(1, 26):
                    if price_sheet.cell(row=j, column=i).value is None:
                        num_of_none += 1
                if num_of_none == 25:
                    next_month_col = i
                    break
        
        # Settings から行・列番号を直接取得
        for i in range(month_cell.column, next_month_col):
            id_cell = price_sheet.cell(row=Settings.ID_ROW_IN_PRICE, column=i)
            price_cell = price_sheet.cell(row=Settings.PRICE_ROW_IN_PRICE, column=i)
            if id_cell.value and price_cell.value:
                id_price_dict[id_cell.value] = price_cell.value
                
        # 在庫表に転記
        sheetnames = stock_list.sheetnames
        stock_sheet = None
        for sheetname in sheetnames:
            if converted_year_month in sheetname:
                stock_sheet = stock_list[sheetname]
                break
        
        if not stock_sheet:
            messagebox.showerror("エラー", f"{converted_year_month}の在庫シートが見つかりません")
            return

        fix_point_count = 0
        for row_num in range(1, stock_sheet.max_row+1):
            id = stock_sheet.cell(row=row_num, column=Settings.ID_COLUMN_IN_STOCK).value
            price = id_price_dict.get(id, None)
            if price:
                stock_sheet.cell(row=row_num, column=Settings.PRICE_COLUMN_IN_STOCK, value=price)
                fix_point_count += 1
                print(f"ID: {id}, Price: {price} updated")

        try:
            stock_list.save(Settings.STOCK_FILE_PATH)
            messagebox.showinfo("成功", f"{fix_point_count}件の価格を更新しました")
        except PermissionError:
            messagebox.showerror("エラー", "在庫ファイルが開かれています。閉じてから再度実行してください。")
        except Exception as e:
            messagebox.showerror("エラー", f"予期しないエラーが発生しました: {e}")

    def update_sales_list(self):
        try:
            stock_list = opx.load_workbook(Settings.STOCK_FILE_PATH, data_only=True)
            sales_list = opx.load_workbook(Settings.SALES_FILE_PATH, data_only=True)
        except FileNotFoundError as e:
            messagebox.showerror("エラー", f"ファイルが見つかりません: {e}")
            return
        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの読み込みに失敗しました: {e}")
            return

        selected_year, selected_month = self.get_selected_year_month()
        selected_year_int = int(selected_year[:-1])
        selected_month_int = int(selected_month[:-1])
        converted_year_month = f"{selected_year_int}{selected_month_int:02}"
        
        sheetnames = stock_list.sheetnames
        stock_sheet = None
        for sheetname in sheetnames:
            if converted_year_month in sheetname:
                stock_sheet = stock_list[sheetname]
                break
        
        if not stock_sheet:
            messagebox.showerror("エラー", f"{converted_year_month}の在庫シートが見つかりません")
            return

        sales_sheet = sales_list.active

        id_price_dict = {}
        for row in range(Settings.DATA_START_ROW_IN_STOCK, stock_sheet.max_row + 1):
            id = stock_sheet.cell(row=row, column=Settings.ID_COLUMN_IN_STOCK).value
            price = stock_sheet.cell(row=row, column=Settings.PRICE_COLUMN_IN_STOCK).value
            if id and price:
                id_price_dict[id] = price

        fix_point_count = 0
        for row in range(1, sales_sheet.max_row+1):
            id = sales_sheet.cell(row=row, column=Settings.ID_COLUMN_IN_SALES).value
            price = id_price_dict.get(id, None)
            if price:
                try:
                    sales_value = sales_sheet.cell(row=row, column=Settings.SALES_COLUMN_IN_SALES).value
                    sales_num_value = sales_sheet.cell(row=row, column=Settings.SALES_NUM_COLUMN_IN_SALES).value
                    
                    if sales_value is not None and sales_num_value is not None:
                        sales = float(sales_value)
                        sales_num = float(sales_num_value)
                        
                        if sales and sales_num and sales != 0:
                            profit = sales - sales_num * price
                            profit_rate = profit / sales  # 利益率の計算を修正
                            sales_sheet.cell(row=row, column=Settings.PROFIT_COLUMN_IN_SALES, value=profit)
                            sales_sheet.cell(row=row, column=Settings.PROFIT_RATE_COLUMN_IN_SALES, value=profit_rate)
                            fix_point_count += 1
                            print(f"ID: {id}, Price: {price} updated")
                except (ValueError, TypeError, ZeroDivisionError) as e:
                    print(f"行 {row} でデータ変換エラー: {e}")
                    continue

        try:
            sales_list.save(Settings.SALES_FILE_PATH)
            messagebox.showinfo("成功", f"{fix_point_count}件の売上データを更新しました")
        except PermissionError:
            messagebox.showerror("エラー", "売上ファイルが開かれています。閉じてから再度実行してください。")
        except Exception as e:
            messagebox.showerror("エラー", f"予期しないエラーが発生しました: {e}")

    def search_string_in_row(self, sheet, row, search_string):
        """指定されたExcelシートの特定の行から、指定された文字列を含むセルを検索する関数"""
        target_cell = None
        for row_data in sheet.iter_rows(min_row=row, max_row=row):
            for cell in row_data:
                if cell.value is not None:
                    try:
                        if search_string in str(cell.value):
                            target_cell = cell
                            break
                    except Exception:
                        continue
            if target_cell:
                break
        return target_cell
    
    def get_selected_year_month(self):
        selected_year = self.year_var.get()
        selected_month = self.month_var.get()
        return selected_year, selected_month

    def calculate_calendar(self, year, month, gap):
        if month + gap < 13 and month + gap > 0:
            return f"{year}{month+gap:02}"
        elif month + gap > 12:
            return f"{year+1}{month+gap-12:02}"
        else:
            return f"{year-1}{month+gap+12:02}"

if __name__ == "__main__":
    app = App()
    app.mainloop()