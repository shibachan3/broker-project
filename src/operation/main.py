# -*- coding:utf-8 -*-
import tkinter
import tkinter.filedialog

from src.operation.to_write import to_write


class Application(tkinter.Frame):
    def __init__(self, root=None):
        super().__init__(root, width=380, height=280,
                         borderwidth=1, relief='groove')
        self.root = root
        self.pack()
        self.pack_propagate(0)
        self.create_widgets()

    def create_widgets(self):
        # 閉じるボタン
        quit_btn = tkinter.Button(
            self,
            text="閉じる",
            command=self.root.destroy)
        quit_btn.pack(side='bottom')

        # 読み込みボタンの作成と配置
        self.read_button = tkinter.Button(
            self,
            text='ファイル読み込み',
            command=self.read_button_func
        )
        self.read_button.pack()

        # ラジオボタンの設定
        self.selected_radio = tkinter.StringVar()
        self.radio_1 = tkinter.Radiobutton(
            self, text="新規作成",value="new", command=self.radio_click, variable=self.selected_radio
            )
        self.radio_2 = tkinter.Radiobutton(
            self, text="更新",value="update", command=self.radio_click, variable=self.selected_radio
            )
        self.radio_1.pack()
        self.radio_2.pack()

    def radio_click(self):
        # ラジオボタンの値を取得
        self.radio_value = self.selected_radio.get()

    def read_button_func(self):
        '読み込みボタンが押された時の処理'
        if self.radio_value == "new":
            input_file_path = tkinter.filedialog.askopenfilename(
                title="csvファイル選択", filetypes=[("Csv File",".csv")]
                )
            output_file_path = tkinter.filedialog.asksaveasfilename(
                title = "Save",
                filetypes = [("Excel files", ".xlsx .xls")]
                )
            # 自動書き込み
            to_write(sedori_path=input_file_path, output_path=output_file_path)

        if self.radio_value == "update":
            input_csv_path = tkinter.filedialog.askopenfilename(
                title="csvファイル選択", filetypes=[("Csv File",".csv")]
                )
            input_excel_path = tkinter.filedialog.askopenfilename(
                title="excelファイル選択", filetypes=[("Excel File",".xlsx")]
                )

            output_file_path = tkinter.filedialog.asksaveasfilename(
                title = "Save",
                filetypes = [("Excel files", ".xlsx .xls")]
                )
            # 自動書き込み
            to_write(input_csv_path, input_excel_path, output_file_path)



# GUIアプリ生成
root = tkinter.Tk()
root.title('セドライト')
root.geometry('400x300')
app = Application(root=root)
app.mainloop()