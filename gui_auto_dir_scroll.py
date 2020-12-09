# TODO: folder selected error E D drive 
# TODO: folder selected error E D drive to gui_auto_dir2.py
import copy
import glob
import functools
import os
import PyPDF2
import sys
import time

from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from tkinter import messagebox

from mkdirs import *

class Application(ttk.Notebook):

    def __init__(self, master=None):
        super().__init__(master)

        self.clock(master)

        # master.title:164行
        master.geometry('372x450')
        master.maxsize(372, 450)
        master.configure(background='white')

        style = ttk.Style()
        style.configure('new.TFrame',foreground='black', background='white')

        tab1 = ttk.Frame(self.master, style='new.TFrame')
        tab2 = ttk.Frame(self.master, style='new.TFrame')

        self.add(tab1, text='自動フォルダ作成')
        self.add(tab2, text='PDF一括結合')

        Tab1(master=tab1)
        Tab2(master=tab2)

        self.pack()

    def clock(self, master):
        t  = time.strftime('%Y/%m/%d %H:%M:%S')
        master.title('自動化ツール {}'.format(t))
        master.after(1000, lambda :self.clock(master))

    # ユーザーがディレクトリを指定する関数
    def dirdialog_clicked(self, entry):
        iDir = os.path.abspath(os.path.dirname(sys.argv[0]))
        iDir_path = filedialog.askdirectory(initialdir=iDir)
        entry.set(iDir_path)

class Tab1(Frame, Application):

    def __init__(self, master=None):

        self.list_items = ['A会社', 'B会社', 'C会社', 'D会社', 'E会社', 'F支店', 'G支店', 'H支店', 'I支店', 'J支店']
        num_list = len(self.list_items)
        self.mkdir_list = []

    # ========================================================================================================
        hs = ttk.Style()
        hs.configure('header.TFrame', background='white')

        # Tab1のFrame1の作成
        frame1 = ttk.Frame(master, padding=10, relief='solid', style='header.TFrame')
        frame1.grid(row=0, column=1, sticky=W+E)

        # 「フォルダ参照」ラベルの作成
        dir_label = ttk.Label(frame1, text='フォルダ参照>>', padding=(0, 5, 5, 0), background='white')
        dir_label.pack(side=LEFT)

        # 「フォルダ参照」エントリーの作成
        self.entry = StringVar()
        dir_entry = ttk.Entry(frame1, textvariable=self.entry, width=30)
        dir_entry.pack(side=LEFT)

        bs = ttk.Style()
        bs.configure('Length.TButton', width=8)

        # 「フォルダ参照」ボタンの作成
        dir_button = ttk.Button(frame1, text='参照', style='Length.TButton', command=lambda :super(Tab1, self).dirdialog_clicked(self.entry))
        dir_button.pack(side=LEFT, padx=3)
    # ========================================================================================================
        frame2 = ttk.Frame(master, padding=10, relief='solid', style='header.TFrame')
        frame2.grid(row=1, column=1, sticky=W)
        self.canvas = Canvas(frame2, width=330, height=200, bg='white')
        self.canvas.bind_all('<MouseWheel>', self.mouse_y_scroll)
        self.canvas.grid(row=1, rowspan=num_list, column=1)

        vbar = ttk.Scrollbar(frame2, orient='vertical')
        vbar.grid(row=1,rowspan=num_list,column=2, sticky='ns')

        sc_hgt = int(150/6*(num_list+1))-15
        self.canvas.config(scrollregion=(0,0,500,sc_hgt))

        vbar.config(command=self.canvas.yview)
        self.canvas.config(yscrollcommand=vbar.set)

    # ========================================================================================================
        self.frame3 = Frame(self.canvas, bg='white')
        self.canvas.create_window((0,0),window=self.frame3,anchor=NW,width=self.canvas.cget('width')) 
        
        num_label=ttk.Label(self.frame3, text='No.', font=('', 9,'bold','roman','normal','normal'), background='white')
        num_label.grid(row=1,column=0)

        sel_label=ttk.Label(self.frame3,width=5,text='Select', font=('',8,'bold','roman','normal','normal'), background='white')
        sel_label.grid(row=1,column=1) 

        item_label=ttk.Label(self.frame3,width=40,text='item', font=('',9,'bold','roman','normal','normal'), background='white')
        item_label.grid(row=1,column=2)

        self.ch_items_list = [False for fal in range(len(self.list_items))]
        
        row = 0
        end_row=len(self.list_items)

        for row in range(end_row):
            #色の設定
            if row%2==0:
                color='#D3D3D3'  
            else:
                color='white'

            bln=BooleanVar()
            bln.set(False)
            f = functools.partial(self.after_cb, bln, row)
            c = Checkbutton(self.frame3, width=5,text='', background='white', variable = bln, command=f)
            c.grid(row=row+2,column=1,padx=0,pady=0,ipadx=0,ipady=0) 
            b1=ttk.Label(self.frame3,width=40,text=self.list_items[row],background=color)
            b1.grid(row=row+2,column=2,padx=0,pady=0,ipadx=0,ipady=0)

            num = ttk.Label(self.frame3, text='{}'.format(row+1), background='white')
            num.grid(row=row+2, column=0,padx=0,pady=0,ipadx=0,ipady=0)

    # ========================================================================================================
        f4_style = ttk.Style()
        f4_style.configure('framecl.TFrame', background='white')
        frame4 = ttk.Frame(master, width=450, style='framecl.TFrame')
        frame4.grid(row=2, column=1, sticky=W)

        f = functools.partial(self.check_all_checkeboxes, row, end_row)
        all_check_button = Button(frame4, text='全選択', overrelief='groove', command=f)
        all_check_button.grid(row=0, column=1, padx=5)

        f = functools.partial(self.clear_all_checkeboxes, row, end_row)
        all_check_button = Button(frame4, text='全選択', overrelief='groove', command=f)
        all_check_button.grid(row=0, column=2, padx=5)

        label_frame = LabelFrame(frame4, text='オプション', bd=3, width=200, background='white')
        label_frame.grid(row=0, column=0, padx=20, pady=10)

        self.v1 = BooleanVar(label_frame)
        self.v1.set(True)
        self.v2 = BooleanVar(label_frame)
        self.v2.set(False)

        radio_num = Checkbutton(label_frame, text='番号', font=('',9,'bold','roman','normal','normal'),variable=self.v1, width=10, bg='white')
        radio_num.pack(side=LEFT, pady=2)
        radio_date = Checkbutton(label_frame, text='日付', font=('',9,'bold','roman','normal','normal'), variable=self.v2, width=10, bg='white')
        radio_date.pack(side=LEFT, pady=2)
        self.radios = [radio_num, radio_date]
    # ========================================================================================================
        frame5 = Frame(master, background='white')
        frame5.grid(row=3,  column=1, sticky='news')
        attention_text='※「番号」を選択した場合は、選択したフォルダ名の前に番号振りがされ、\r   「日付」を選択した場合、現在の年月日を選択した\rフォルダ名の後に書き込まれます。'
        num_label=ttk.Label(frame5,text=attention_text, font=('',9,'bold','roman','normal','normal'), background='white')
        num_label.grid(row=0,column=0,padx=0,pady=0,ipadx=0,ipady=0)
    # ========================================================================================================
        style = ttk.Style()
        style.configure('footer.TFrame',foreground='black', background='white', bd=5)
        
        # Tab2のFrame3の作成
        frame3 = ttk.Frame(master, padding=(0, 15, 0, 25), style='footer.TFrame')
        frame3.grid(row=4,  column=1, sticky='we')

        # 作成ボタンの設置
        create_dir_button = ttk.Button(frame3, text='作成', command=lambda: self.clicked())
        create_dir_button.pack(fill='x', padx=50, side='left')

        # キャンセルボタンの設置
        cancel_button = ttk.Button(frame3, text=('閉じる'), command=master.quit)
        cancel_button.pack(fill='x', padx=35, side='left')
    # ========================================================================================================

    # チェックボックスを押したとき
    def after_cb(self, bln, irow):
        if bln.get() == True:
            self.ch_items_list[irow] = True
        else:
            self.ch_items_list[irow] = False
        self.mkdir_list = [self.list_items[i] for i in range(len(self.list_items)) if self.ch_items_list[i] == True]


    # スクロール可能な範囲を指定
    def mouse_y_scroll(self, event):
        if event.delta > 0:
            self.canvas.yview_scroll(-1, 'units')
        elif event.delta < 0:
            self.canvas.yview_scroll(1, 'units')

    # フォルダ作成関数
    def clicked(self):
        dir_path = self.entry.get()
        checked_num = self.v1.get()
        checked_date = self.v2.get()

        if not dir_path:
            messagebox.showerror('エラー', 'パスの指定がありません')
            return
        elif not dir_path[:3] == 'C:/':
            messagebox.showerror('エラー', 'パスの入力が違います!')
            return
        else:
            pass

        try:
            msg = messagebox.askyesno(
                '確認', 
                '作成先のパスはあっていますか?\n' +
                '-------------------------------------------------------------------------------\n' +
                '{}\n'.format(dir_path) + 
                '-------------------------------------------------------------------------------\n' +
                '作成しますがよろしいですか?')
            if msg == True:
                if not self.mkdir_list:
                    messagebox.showerror('エラー', '作成するフォルダ名が選択されていません')
                    return
                elif checked_num == False and checked_date == False and self.mkdir_list: # オプションなし(フォルダ名)
                    mkdir_items_no_slc(dir_path, self.mkdir_list)
                    messagebox.showinfo('フォルダ作成情報', 'フォルダが作成されました')
                    return
                elif checked_num == False and checked_date == True and self.mkdir_list: # オプション有り(フォルダ名 + 日付)
                    mkdir_items_date(dir_path, self.mkdir_list)
                    messagebox.showinfo('フォルダ作成情報', 'フォルダが作成されました')
                    return
                elif checked_num == True and checked_date == False and self.mkdir_list: # デフォルト(番号 + フォルダ名)
                    mkdir_items_num(dir_path, self.mkdir_list)
                    messagebox.showinfo('フォルダ作成情報', 'フォルダが作成されました')
                    return
                elif checked_num == True and checked_date == True and self.mkdir_list: # オプションを有り(番号 + フォルダ名 + 日付)
                    mkdir_items_all_slc(dir_path, self.mkdir_list)
                    messagebox.showinfo('フォルダ作成情報', 'フォルダが作成されました')
                    return
                else:
                    return
        except FileExistsError:
            messagebox.showwarning('警告', '同じ名前のフォルダが存在しています')
            return
        except Exception as e:
            print('例外が発生しました', e)
            messagebox.showwarning('警告', '予期しないエラーが発生しました')
            return


    # 「全選択」ボタンを押下したとき
    def check_all_checkeboxes(self, row, end_row):
        self.mkdir_list = copy.copy(self.list_items)
        self.ch_items_list.clear()

        for _ in range(len(self.list_items)):
            self.ch_items_list.append(True)

        for row in range(end_row):
            #色の設定
            if row%2==0:
                color='#D3D3D3'  
            else:
                color='white'

            bln=BooleanVar()
            bln.set(True)
            f = functools.partial(self.after_cb, bln, row)
            c = Checkbutton(self.frame3, width=5,text='', background='white', variable = bln, command=f)
            c.grid(row=row+2,column=1,padx=0,pady=0,ipadx=0,ipady=0) 
            b1=ttk.Label(self.frame3,width=40,text=self.list_items[row],background=color)
            b1.grid(row=row+2,column=2,padx=0,pady=0,ipadx=0,ipady=0)

            num = ttk.Label(self.frame3, text='{}'.format(row+1), background='white')
            num.grid(row=row+2, column=0,padx=0,pady=0,ipadx=0,ipady=0)


    # 「全解除」ボタンを押下したとき
    def clear_all_checkeboxes(self, row, end_row):
        self.mkdir_list.clear()
        self.ch_items_list = [False for fal in range(len(self.list_items))]

        for row in range(end_row):
            #色の設定
            if row%2==0:
                color='#D3D3D3'  
            else:
                color='white'

            bln=BooleanVar()
            bln.set(False)
            f = functools.partial(self.after_cb, bln, row)
            c = Checkbutton(self.frame3, width=5,text='', background='white', variable = bln, command=f)
            c.grid(row=row+2,column=1,padx=0,pady=0,ipadx=0,ipady=0) 
            b1=ttk.Label(self.frame3,width=40,text=self.list_items[row],background=color)
            b1.grid(row=row+2,column=2,padx=0,pady=0,ipadx=0,ipady=0)

            num = ttk.Label(self.frame3, text='{}'.format(row+1), background='white')
            num.grid(row=row+2, column=0,padx=0,pady=0,ipadx=0,ipady=0)

class Tab2(Frame, Application):

    def __init__(self, master=None):
        super().__init__(master)
    # ========================================================================================================
        hs = ttk.Style()
        hs.configure('header.TFrame', background='white')

        # Tab1のFrame1の作成
        frame1 = ttk.Frame(master, padding=(10, 10, 30, 10), relief='solid', style='header.TFrame')
        frame1.grid(row=0, column=1, sticky=W+E)

        # 「フォルダ参照」ラベルの作成
        dir_label = ttk.Label(frame1, text='フォルダ参照>>', padding=(0, 5, 5, 0), background='white')
        dir_label.pack(side=LEFT)

        # 「フォルダ参照」エントリーの作成
        self.entry = StringVar()
        dir_entry = ttk.Entry(frame1, textvariable=self.entry, width=30)
        dir_entry.pack(side=LEFT)

        bs = ttk.Style()
        bs.configure('Length.TButton', width=8)

        # 「フォルダ参照」ボタンの作成
        dir_button = ttk.Button(frame1, text='参照', style='Length.TButton', command=lambda :self.dirdialog_clicked(self.entry))
        dir_button.pack(side=LEFT, padx=3)
    # ========================================================================================================
        ms = ttk.Style()
        ms.configure('main.TFrame',foreground='black', background='white', bd=5)
        frame2 = ttk.Frame(master, padding=(10, 15, 0, 25), style='main.TFrame')
        frame2.grid(row=2,  column=1, sticky='we')
        label1 = Label(frame2, text="複数のPDFファイルを一括結合して1つのPDFファイルを生成します", background='white')
        label1.grid(row=0, column=0)
        label2 = Label(frame2, text="「参照ボタン」を押して複数入ったPDFファイルのフォルダを選択してください", background='white')
        label2.grid(row=1, column=0)
        label3 = Label(frame2, text="「実行ボタン」を押してどこにPDFファイルを保存するかを選択してください", background='white')
        label3.grid(row=2, column=0)
    # ========================================================================================================
        fs = ttk.Style()
        fs.configure('footer.TFrame',foreground='black', background='white', bd=5)
        
        # Tab2のFrame3の作成
        frame3 = ttk.Frame(master, padding=(0, 15, 0, 25), style='footer.TFrame')
        frame3.place(x=10, y=350)
        # frame3.grid(row=6,  column=1, sticky='we')

        # 実行ボタンの設置
        create_dir_button = ttk.Button(frame3, text='実行', command=lambda :self.clicked(master))
        create_dir_button.pack(fill='x', padx=50, side='left')

        # キャンセルボタンの設置
        cancel_button = ttk.Button(frame3, text=('閉じる'), command=master.quit)
        cancel_button.pack(fill='x', padx=35, side='left')
    # ========================================================================================================

    def clicked(self, master):
        dir_path = self.entry.get()     
        if not dir_path:
            messagebox.showerror('エラー', 'パスの指定がありません')
            return

        absPath = os.path.abspath(os.path.dirname(__file__))
        messagebox.showinfo('PDF一括結合プログラム','PDFの保存先を選択してください!')
        dirPath = filedialog.askdirectory(initialdir = absPath)
        print(dirPath)


        file_collection = dir_path + '/' + '*.pdf'
        filepath = sorted(glob.glob(file_collection))

        # フォルダ内のPDFファイルのパスの一覧
        fileslist = [dir_path + '/'+ os.path.basename(f) for f in filepath]
        marge_fileslist = [os.path.basename(f) for f in filepath]
        total_pages = 0

        try:
            # １つのPDFファイルにまとめる
            pdf_writer = PyPDF2.PdfFileWriter()
            for pdf_file in fileslist:
                pdf_reader = PyPDF2.PdfFileReader(str(pdf_file))
                num_pages = pdf_reader.getNumPages()
                total_pages += num_pages
                for i in range(pdf_reader.getNumPages()):
                    pdf_writer.addPage(pdf_reader.getPage(i))
        except Exception:
            return

        print('総ページ数 :', total_pages)
        print('総ファイル数 :', len(marge_fileslist))
        # print('以下のパスにPDFファイルが生成されました')
        try:
            # 保存ファイル名（先頭と末尾のファイル名で作成）
            merged_file = marge_fileslist[0][:-4] + "-" + marge_fileslist[-1][:-4] + '.pdf'
            merged_filePath = dirPath + '/'+ merged_file
        except IndexError:
            messagebox.showwarning('エラー', 'pdfファイルがありません\rPDFファイルがあるフォルダを選択してください')
            return

        # # 保存
        # with open(merged_filePath, "wb") as f:
        #     pdf_writer.write(f)
        # print('保存できました!')

        ms = ttk.Style()
        ms.configure('main.TFrame',foreground='black', background='white', bd=5)
        frame2 = ttk.Frame(master, padding=(10, 15, 0, 25), style='main.TFrame')
        frame2.grid(row=2,  column=1, sticky='we')
        label1 = Label(frame2, text="複数のPDFファイルを一括結合して1つのPDFファイルを生成します", background='white')
        label1.grid(row=0, column=0)
        label2 = Label(frame2, text="「参照ボタン」を押して複数入ったPDFファイルのフォルダを選択してください", background='white')
        label2.grid(row=1, column=0)
        label3 = Label(frame2, text="「実行ボタン」を押してどこにPDFファイルを保存するかを選択してください", background='white')
        label3.grid(row=2, column=0)
        label4 = Label(frame2, text="***************************▼出力結果▼*******************************", background='white')
        label4.grid(row=3, column=0)
        label5 = Label(frame2, text="===========================▼出力結果▼===============================", background='white')
        label5.grid(row=4, column=0)
        
        return

if __name__ == "__main__":
    win = Tk()
    app = Application(master=win)
    app.mainloop()