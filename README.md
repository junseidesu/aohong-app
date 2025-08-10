# 在庫単価転記アプリ

Excel の既存帳票（単価表 / 在庫表 / 売上分析表）間で単価・利益情報を転記・計算するための Python + CustomTkinter GUI アプリケーションです。

## 特長
- 単価表 ("一般総平均" シート) から対象年月の ID/単価を抽出し在庫表へ一括転記
- 在庫表に反映済み単価を用いて売上表に利益・利益率を計算し書き込み
- 設定ウィンドウから Excel ファイルパスおよび行・列番号を GUI で変更
- 設定保存後、メインウィンドウ UI (タイトル / サイズ / テーマ / アイコン) が即時反映
- 設定 JSON (`settings.json`) による永続化。欠損キーは自動でデフォルト補完
- 設定ウィンドウは開いたとき最前面に出て、既存インスタンスがあれば再利用

## 動作要件
- Python 3.10 以降 (3.13 でも動作確認)
- Windows (他 OS でも基本的には動く想定ですが未検証)
- 必要ライブラリ:
  - `customtkinter`
  - `openpyxl`

### インストール
PowerShell:
```powershell
pip install customtkinter openpyxl
```

## ファイル構成 (抜粋)
```
main.py                # アプリ本体
settings.json          # 保存された設定 (初回は無い場合あり)
2025在庫.xlsx
フロンガス単価表2025.xlsx
R706 得意先別売上分析表.xlsx
reclaim_icon.ico       # アイコン
```

## 実行
PowerShell でリポジトリディレクトリに移動して:
```powershell
python .\main.py
```

## 使い方
1. アプリ起動後、対象の「年」「月」をドロップダウンで選択
2. 「在庫単価を在庫表に転記」: 単価表→在庫表へ単価一括反映
3. 「在庫単価を売上表に転記」: 在庫表の単価を使って売上表へ利益・利益率計算反映
4. 「設定」ボタンで設定ウィンドウを開き、各 Excel ファイルを選択後「設定を保存」
   - 保存した時点で UI テーマやタイトルが即反映されます
   - 途中で閉じた場合 (保存押下前) の変更は破棄されます
5. 必要に応じて「設定をリセット」でデフォルトへ戻し自動保存

## 設定 (Settings クラス)
`Settings._DEFAULT_VALUES` が唯一のデフォルトソースで、起動時に `settings.json` (あれば) で上書きされます。

### `settings.json` 例
```json
{
  "files": {
    "price_file_path": "フロンガス単価表2025.xlsx",
    "stock_file_path": "2025在庫.xlsx",
    "sales_file_path": "R706 得意先別売上分析表.xlsx"
  },
  "positions": {
    "id_row_in_price": 3,
    "price_row_in_price": 25,
    "id_column_in_stock": 3,
    "price_column_in_stock": 10,
    "data_start_row_in_stock": 7,
    "id_column_in_sales": 4,
    "profit_column_in_sales": 9,
    "profit_rate_column_in_sales": 10,
    "sales_column_in_sales": 6,
    "sales_num_column_in_sales": 8
  }
}
```

### 値の適用フロー
1. アプリ起動 -> `Settings.load_settings()` がデフォルトで初期化後 JSON を反映
2. メイン UI が Settings のクラス変数値を参照
3. 設定ウィンドウでファイル選択 -> クラス変数を即更新 (未保存)
4. 「設定を保存」 -> JSON 書き込み -> `App.refresh_settings_ui()` で UI を更新
5. 次回起動時に保存値が再現

## 処理ロジック概要
### 在庫単価転記
1. 年月 (YYYYMM) を組み立て
2. 単価表 "一般総平均" 1 行目に対象文字列セルを検索
3. 次月開始列との間で ID_ROW / PRICE_ROW を走査し ID→単価 dict を構築
4. 在庫ファイルシート名に YYYYMM を含むシートへ単価を書き込み

### 売上表利益計算
1. 在庫表から ID→単価 dict 作成 (データ開始行以降)
2. 売上表を 1 行目から走査し ID マッチ行で: `profit = sales - sales_num * price`, `profit_rate = profit / sales`
3. 指定列に書き込み

## エラーハンドリング
- ファイル未発見 / PermissionError を messagebox で通知
- データ変換エラー (ValueError など) はコンソールにログ
- 設定読み込み失敗は標準出力ログのみ (今後 GUI 通知を追加しても良い)

## 最前面表示の仕組み
設定ウィンドウ生成時に一時的に `-topmost` 属性を True にし `lift()` / `focus_force()` 後、300ms で False に戻すことで自然な前面化を実現。

## よくある質問 (FAQ)
Q. 設定を変えたのにメイン画面が変わらない。
A. 「設定を保存」ボタンを押したか確認してください。保存後は即反映されます。

Q. テーマ (THEME / COLOR_THEME) を JSON からも変更したい。
A. `save_settings()` の書き込み dict にキーを追加し、`load_settings()` で読み戻す処理を追加してください。

## 今後の改善候補
- 設定ウィンドウに数値 (行/列) 編集フィールドを追加
- 未保存変更がある状態で閉じる際の確認ダイアログ
- ログ出力 (logging) 導入とログファイル保存
- requirements.txt 追加とバージョンピン
- 単体テスト (pytest) 整備
- 進捗/件数をプログレスバーで表示

## ライセンス
(必要に応じてここにライセンス種別を記載)

## 開発者向けメモ
- デフォルト値は `_DEFAULT_VALUES` のみを編集
- UI へ新たな設定を反映したい場合は `App.refresh_settings_ui()` を拡張

---
不明点や機能追加要望があれば Issue / 連絡でお知らせください。
