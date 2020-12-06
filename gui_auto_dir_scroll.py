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
        self.add(tab2, text='1つのPDFファイルにまとめる')

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
        print(iDir_path)
        entry.set(iDir_path)

class Tab1(Frame, Application):

    def __init__(self, master=None):

        self.list_items = ['会社A', '会社B', '会社C', '会社D', '会社E', '支店F', '支店G', '支店H', '支店I', '支店J']
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
        create_dir_button = ttk.Button(frame3, text='作成', command=lambda: self.clicked_mkdir2())
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
    def clicked_mkdir2(self):
        dir_path = self.entry.get()
        checked_num = self.v1.get()
        checked_date = self.v2.get()

        if not dir_path:
            messagebox.showerror('エラー', 'パスの指定がありません')
            return
        try:
            msg = messagebox.askyesno(
                '確認', 
                '作成先のパスはあっていますか?\n' +
                '--------------------------------\n' +
                '{}\n'.format(dir_path) + 
                '--------------------------------\n' +
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
        dir_button = ttk.Button(frame1, text='参照', style='Length.TButton', command=lambda :self.dirdialog_clicked(self.entry))
        dir_button.pack(side=LEFT, padx=3)
    # ========================================================================================================

    # ========================================================================================================
        style = ttk.Style()
        style.configure('footer.TFrame',foreground='black', background='white', bd=5)
        
        # Tab2のFrame3の作成
        frame3 = ttk.Frame(master, padding=(0, 15, 0, 25), style='footer.TFrame')
        frame3.grid(row=4,  column=1, sticky='we')

        # 実行ボタンの設置
        create_dir_button = ttk.Button(frame3, text='作成', command=lambda :super(Tab2, self).dirdialog_clicked(self.entry))
        create_dir_button.pack(fill='x', padx=50, side='left')

        # キャンセルボタンの設置
        cancel_button = ttk.Button(frame3, text=('閉じる'), command=master.quit)
        cancel_button.pack(fill='x', padx=35, side='left')
    # ========================================================================================================

    def clicked_mkdir2(self, master):
        frame2 = ttk.Frame(master, padding=(0, 15, 0, 25), style='footer.TFrame')
        frame2.grid(row=2,  column=1, sticky='we')
        

if __name__ == "__main__":
    win = Tk()
    app = Application(master=win)
    app.mainloop()