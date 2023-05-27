# -*- coding: utf-8 -*-
"""
HOTAP (Houwen Online Tab Access Program) tab opener for CSV files.

Created on Mon May 15 14:34:00 2023

@author: Willem Roefs

pyinstaller --onefile --hidden-import tkinter -w --splash splash.png 
--icon Logo-Houwen-Online.ico --add-data "Logo-Houwen_Online.ico;." hotap.py
"""


import random
import webbrowser
# import tkinter as tk
from tkinter import IntVar, StringVar, BooleanVar, Button, Entry, Label
from tkinter import filedialog, N, S, E, W
from tkinter import ttk, messagebox
from keyboard import press_and_release
import pandas as pd
import sys
from os import path

try:
    import pyi_splash
except: 
    print('run at dev')
VERSION = '0.0.1'


class TabOpener(ttk.Frame):
    '''Tabopener window class'''

    def __init__(self, name='tab', master=None):
        try:
            pyi_splash.close()
        except:
            print('Devinit')
        self.now = 0
        self.count_tabs = 0
        ttk.Frame.__init__(self, master, name=name)
        # print(self.resource_path('Logo-Houwen-Online.ico'))
        self.master.iconbitmap(str(
            self.resource_path('Logo-Houwen-Online.ico')))
        self.master.title('HOTAP ' + VERSION)

        self.master.protocol("WM_DELETE_WINDOW", self.stop_execution)
        self.master.lift()
        self.master.attributes('-topmost', True)
        top = self.master.winfo_toplevel()
        top.rowconfigure(0, weight=1)
        top.columnconfigure(0, weight=1)
        self.max_tabs_open = IntVar()
        self.max_tabs_open.set(10)
        self.x_pad = 5
        self.time_between_opening = IntVar()
        self.time_between_opening.set(5000)  # 5000  # in milliseconds
        self.start_opening = BooleanVar()
        self.url_path_text = StringVar()
        self.url_path = None
        self.start_opening.set(False)
        self.rowconfigure(2, weight=1)
        self.columnconfigure(2, weight=1)
        self.grid(sticky=N+S+E+W)
        self.main_frame = ttk.Frame(self)
        self.main_frame.grid(sticky=N+S+E+W)
        self.top_logo = ttk.Label(self.main_frame,
                                  text='Houwen Online Tab Access Program',
                                  font=('Helvetica bold', 15))
        self.top_logo.grid(row=0, column=0, sticky='new', columnspan=2,
                           padx=self.x_pad)
        self.open_file_btn = Button(self.main_frame, text='Open CSV file',
                                    command=self.open_file)
        self.open_file_btn.grid(row=1, column=0, sticky='new', columnspan=2,
                                ipady=15, padx=self.x_pad)
        self.file_to_open = Entry(self.main_frame,
                                  textvariable=self.url_path_text)
        self.file_to_open.grid(row=2, column=0, sticky='new', columnspan=2,
                               pady=5, padx=self.x_pad)
        self.tabs_open_lbl = Label(self.main_frame, text='Max tabs open: ')
        self.tabs_open_lbl.grid(row=3, column=0, sticky='new', columnspan=1,
                                pady=5)
        self.tabs_open_entry = Entry(self.main_frame,
                                     textvariable=self.max_tabs_open)
        self.tabs_open_entry.grid(row=3, column=1, sticky='new', columnspan=1,
                                  pady=5, padx=self.x_pad)
        self.time_open_lbl = Label(self.main_frame,
                                   text='Time between tabs opening [ms]: ')
        self.time_open_lbl.grid(row=4, column=0, sticky='new', columnspan=1,
                                pady=5, padx=self.x_pad)
        self.time_open_entry = Entry(self.main_frame,
                                     textvariable=self.time_between_opening)
        self.time_open_entry.grid(row=4, column=1, sticky='new', columnspan=1,
                                  pady=5, padx=self.x_pad)
        self.start_stop_button = ttk.Button(self.main_frame,
                                            text='Start',
                                            command=self.start_stop)
        self.start_stop_button.grid(row=5, column=0,
                                    sticky='new', columnspan=2,
                                    ipady=15, padx=self.x_pad)

        self.exit_button = ttk.Button(self.main_frame, text="Exit",
                                      command=self.stop_execution)
        self.exit_button.grid(row=6, column=0, sticky='new', columnspan=2,
                              ipady=15, padx=self.x_pad)

        # try:
        #     # read URLs from excel file
        #     self.url_path = filedialog.askopenfilename(
        #         parent=self, title="Select Links csv file",
        #         filetypes=[(".csv files", "*.csv")])
        #     df_urls = pd.read_csv(self.url_path)

        #     # shuffle URLs randomly
        #     self.urls = df_urls['URL'].tolist()
        #     random.shuffle(self.urls)
        #     # self.openurl()
        # except FileNotFoundError:
        #     print('FileNotFound')
        #     self.stop_execution()
        #     df_urls = None

    def resource_path(self, relative_path):
        """ Get absolute path to resource"""
        try:
            # PyInstaller creates a temp folder and stores path in _MEIPASS
            base_path = sys._MEIPASS
        except Exception:
            base_path = path.abspath(".")

        return path.join(base_path, relative_path)

    def open_file(self):
        try:
            # read URLs from excel file
            self.url_path = filedialog.askopenfilename(
                parent=self, title="Select Links csv file",
                filetypes=[(".csv files", "*.csv")])
            self.url_path_text.set(self.url_path)
            df_urls = pd.read_csv(self.url_path)

            # shuffle URLs randomly
            self.urls = df_urls['URL'].tolist()
            random.shuffle(self.urls)
            # self.openurl()
        except FileNotFoundError:
            print('FileNotFound')
            messagebox.showerror("Error", "File Not Found")
            df_urls = None

    def stop_execution(self):
        '''Closing the window of if an error occurs destoys the frame'''
        self.master.destroy()

    def start_stop(self):
        '''Start stop button function'''
        if self.url_path is not None:
            if self.start_opening.get():
                self.start_opening.set(False)
                self.start_stop_button.config(text='Start')
                self.master.update_idletasks()
            else:
                self.start_opening.set(True)
                self.start_stop_button.config(text='Stop')
                self.master.update_idletasks()
                self.openurl()
        else:
            messagebox.showerror("Error", "No file entered!")

    def openurl(self):
        '''Opens the urls at a specified interval'''
        if self.count_tabs >= 10:
            self.closeoldtab()
        # for url in self.urls:
        url = self.urls[self.now]
        webbrowser.open(url)
        self.count_tabs += 1  # len(self.urls)
        self.now = random.randint(0, len(self.urls)-1)
        # print(self.now)
        if self.count_tabs >= self.max_tabs_open.get():
            press_and_release('ctrl+1')  # focus on first tab
            self.after(50, self.closeoldtab)

        # print(self.count_tabs)
        if self.start_opening.get():
            self.after(self.time_between_opening.get(), self.openurl)

    def closeoldtab(self):
        '''Closes the left most tab every x seconds'''
        # for amount in self.urls:
        print('deleting ')  # + amount)
        press_and_release('ctrl+w')  # closes tab
        self.count_tabs -= 1
        # self.after(2000, self.closeoldtab)


if __name__ == '__main__':
    TabOpener().mainloop()
