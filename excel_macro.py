import xlwings as xw
import os

def add_macro_to_excel(file_path):
    app = None  # アプリケーション変数の初期化
    try:
        # Excelアプリケーションを非表示で起動
        app = xw.App(visible=False)

        # ファイルが .xlsm でない場合、最初に .xlsm ファイルに変換してコピーを作成
        if not file_path.endswith('.xlsm'):
            new_file_path = os.path.splitext(file_path)[0] + '_with_macro.xlsm'
            # Excelファイルを開いて .xlsm として保存
            wb = app.books.open(file_path)
            wb.save(new_file_path)
            wb.close()
            print(f'ファイルを {new_file_path} に変換しました。')
            file_path = new_file_path  # 新しいファイルパスに切り替え
        else:
            new_file_path = file_path

        # 変換後、または .xlsm のファイルを開く
        wb = app.books.open(new_file_path)

        # VBAモジュールを取得（標準モジュールを追加）
        if not any(module.name == 'Module1' for module in wb.api.VBProject.VBComponents):
            vba_module = wb.api.VBProject.VBComponents.Add(1)  # 標準モジュールを追加
        else:
            vba_module = wb.api.VBProject.VBComponents('Module1')

        # マクロコードを挿入
        macro_code = '''
        Sub A1move()
            'シートを定義
            Dim ws As Worksheet
    
            '全てのシートで以下をループ
            For Each ws In Worksheets
                'シートを選択
                ws.Select
                'A1セルを選択
                ws.Range("A1").Select
            Next
            '最初のシートへ移動
            Sheets(1).Select
        End Sub
        '''
        vba_module.CodeModule.AddFromString(macro_code)

        # マクロ付きのファイルを保存
        wb.save()

    except Exception as e:
        # エラーが発生した場合の処理
        print(f"エラーが発生しました: {e}")

    finally:
        # ワークブックとアプリケーションのクリーンアップ
        if wb:
            wb.close()
        if app:
            app.quit()

    return new_file_path  # 新しいファイルのパスを返す

# テスト用に実行
#add_macro_to_excel(r'/path/to/your/file.xlsx')
