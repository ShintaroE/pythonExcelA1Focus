import tkinter as tk
from tkinter import filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
from excel_macro import add_macro_to_excel

def on_drop(event):
    file_path = event.data  # ドラッグ＆ドロップされたファイルパスを取得
    if file_path.endswith('.xlsx') or file_path.endswith('.xls'):
        label.config(text=f"選択されたファイル: {file_path}")
        global excel_file
        excel_file = file_path  # グローバル変数にファイルパスを保持
    else:
        label.config(text="Excelファイルを選択してください")

def convert_file():
    if excel_file:
        # Excelファイルの読み込みと変換処理（例: CSVに変換）
        output_path = add_macro_to_excel(excel_file)
        label.config(text=f"変換が完了しました: {output_path}")
    else:
        label.config(text="ファイルが選択されていません")

def open_file():
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
    )
    if file_path:
        label.config(text=f"選択されたファイル: {file_path}")
        global excel_file
        excel_file = file_path

def exit_app():
    if messagebox.askokcancel("終了", "アプリケーションを終了しますか？"):
        root.quit()

# アプリケーションのウィンドウ設定
root = TkinterDnD.Tk()  # TkinterDnD2を使用
root.title("ExcelファイルA1セルフォーカス")
root.geometry("400x250")

excel_file = None

# メニューバーの作成
menubar = tk.Menu(root)
root.config(menu=menubar)

# "ファイル"メニューの追加
file_menu = tk.Menu(menubar, tearoff=0)
menubar.add_cascade(label="ファイル", menu=file_menu)
file_menu.add_command(label="ファイルを開く", command=open_file)
file_menu.add_separator()
file_menu.add_command(label="終了", command=exit_app)

# ドラッグ＆ドロップエリア
frame = tk.Frame(root, width=300, height=100, bg="lightgray")
frame.pack(pady=20)

# ドラッグ＆ドロップイベントの設定
frame.drop_target_register(DND_FILES)
frame.dnd_bind('<<Drop>>', on_drop)

# ラベル
label = tk.Label(root, text="ここにExcelファイルをドラッグ＆ドロップしてください")
label.pack(pady=10)

# 変換ボタン
convert_button = tk.Button(root, text="実行", width=20, height=7 , command=convert_file)
convert_button.pack()

root.mainloop()