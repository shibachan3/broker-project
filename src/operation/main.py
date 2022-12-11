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
        radio_1 = tkinter.Radiobutton(self, text="新規作成",value="nwe", variable=self.selected_radio)
        radio_2 = tkinter.Radiobutton(self, text="更新",value="update", variable=self.selected_radio)
        radio_1.pack()
        radio_2.pack()

    def read_button_func(self):
        '読み込みボタンが押された時の処理'
        input_file_path = tkinter.filedialog.askopenfilename()
        output_file_path = tkinter.filedialog.asksaveasfilename(
            title = "Save",
            filetypes = [("Excel files", ".xlsx .xls")]
            )

        # 自動書き込み
        to_write(input_file_path, output_file_path)

# GUIアプリ生成
root = tkinter.Tk()
root.title('セドライト')
root.geometry('400x300')
app = Application(root=root)
app.mainloop()