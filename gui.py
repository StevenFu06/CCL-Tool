from tkinter import ttk
import tkinter as tk
from tkinter import filedialog, messagebox
from PIL import Image, ImageTk
from enovia import Enovia
import threading
from selenium.common.exceptions import UnexpectedAlertPresentException
from multiprocessing import freeze_support
import time
import progressbar
import sys
import datetime as dt
from ccl import CCL
import pandas as pd
from StyleFrame import Styler, StyleFrame
from package import Parser
import re
import os


class Root(tk.Tk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.iconbitmap('icons\\sciex.ico')
        self.title('CCL Tool')
        self._set_style()
        self.config(background='white')

        self._set_ccl_var()
        self._set_notebook()

        self.frames = {}
        for F in (BomCompare, UpdateCCL, DocumentCollector, Illustrations, ROHSCompare):
            frame = F(self, self.main_notebook)
            self.frames[F] = frame

        for F in (ROHSCompare, FilterCompare, InsertDelIllustration):
            frame = F(self, self.extra_notebook)
            self.frames[F] = frame

        self.main_options()
        self.extra_frame()
        self.settings_frame()
        self.update()
        self.minsize(self.winfo_width(), self.winfo_height())

    def _set_style(self):
        self.style = ttk.Style()
        self.style.theme_create('MyStyle', parent='vista', settings={
            'TNotebook': {'configure': {'tabmargins': [2, 5, 2, 0],
                                        'background': 'white'}},
            'TNotebook.Tab': {'configure': {'padding': [50, 2],
                                            'font': ('Segoe', '12'),
                                            'background': 'white'}},
            'TCheckbutton': {'configure': {'background': 'white'}},
            'TLabel': {'configure': {'background': 'white'}},
            'TEntry': {'configure': {'background': 'white'}},
        })
        self.style.layout('Tab',
                          [('Notebook.tab', {
                              'sticky': 'nswe',
                              'children': [(
                                  'Notebook.padding', {'side': 'top', 'sticky': 'nswe', 'children': [(
                                      'Notebook.label', {'side': 'top', 'sticky': ''})]})]})]
                          )
        self.style.theme_use('MyStyle')

    def _set_notebook(self):
        self.notebook = ttk.Notebook(self)

        self.main_notebook = tk.Frame(self.notebook, bg='white')
        self.extra_notebook = tk.Frame(self.notebook, bg='white')
        self.settings_notebook = tk.Frame(self.notebook, bg='white')

        self.notebook.add(self.main_notebook, text='CCL Tools')
        self.notebook.add(self.extra_notebook, text='Extra Tools')
        self.notebook.add(self.settings_notebook, text='Settings')
        self.notebook.pack(expand=True, fill='both')

    def _set_ccl_var(self):
        self.compare_select = tk.BooleanVar()  # For Checkbox bom comparison
        self.update_select = tk.BooleanVar()  # For Checkbox ccl update
        self.docs_select = tk.BooleanVar()  # for checkbox document collection
        self.ills_select = tk.BooleanVar()  # For checkbox illustration collection

        self.cache_dir = '/'  # Cached default open directory
        self.shared = tk.BooleanVar(value=True)  # For shared variables in settings
        self.headless = tk.BooleanVar(value=True)  # For troubleshooting, toggles headless
        self.incomplete_run = tk.BooleanVar()  # For troubleshooting, allows run without all fields
        self.processes = 1  # Number of cores to use when multiprocessing

        self.bom_compare_old = None  # Bom compare old avl bom
        self.bom_compare_new = None  # Bom compare new avl bom
        self.bom_compare_save = None  # Bom compare save location

        self.ccl_update_old = None  # CCL update old avl bom
        self.ccl_update_new = None  # CCL update old avl bom
        self.ccl_update_loc = None  # CCL update ccl save location
        self.ccl_update_save_loc = None  # CCL update new ccl save location

        self.docs_ccl = None  # Document Collect CCL document
        self.docs_paths = []  # Document collect ccl check paths
        self.docs_savedir = None  # Document collect document save directory
        self.docs_user = None  # Document collect enovia username
        self.docs_pass = None  # Document collect enovia password

        self.ill_ccl = None  # Illustration ccl document
        self.ill_cclsave = None  # Illustration ccl save location
        self.ill_save = None  # Illustration save location
        self.ill_scan = None  # Illustration scan location

    def main_options(self):
        self.mainframe = tk.Frame(self.main_notebook, bg='white')
        self.mainframe.pack(anchor='center', expand=True, fill='y')

        frame_compare = tk.Frame(self.mainframe, bg='white')
        frame_compare.pack(fill='x', pady=10, expand=True)
        compare_check = CustomCheckButton(frame_compare, variable=self.compare_select)
        compare_check.pack(side='left', padx=(0, 2))
        compare_button = DoubleTextButton(frame_compare, text_main='Bill of Materials Comparison',
                                          text_sub='Will generate a bill of materials comparison report',
                                          command=lambda: self.raiseframe(BomCompare))
        compare_button.pack(side='left', fill='x', expand=True)

        frame_update = tk.Frame(self.mainframe, bg='white')
        frame_update.pack(fill='x', pady=10, expand=True)
        update_check = CustomCheckButton(frame_update, variable=self.update_select)
        update_check.pack(side='left', padx=(0, 2))
        update_button = DoubleTextButton(frame_update, text_main='Update CCL',
                                         text_sub='Will output an updated CCL',
                                         command=lambda: self.raiseframe(UpdateCCL))
        update_button.pack(side='left', fill='x', expand=True)

        frame_docs = tk.Frame(self.mainframe, bg='white')
        frame_docs.pack(fill='x', pady=10, expand=True)
        docs_check = CustomCheckButton(frame_docs, variable=self.docs_select)
        docs_check.pack(side='left', padx=(0, 2))
        docs_button = DoubleTextButton(frame_docs, text_main='Collect CCL Documents',
                                       text_sub='Will collect all documents associated with CCL',
                                       command=lambda: self.raiseframe(DocumentCollector))
        docs_button.pack(side='left', fill='x', expand=True)

        frame_ills = tk.Frame(self.mainframe, bg='white')
        frame_ills.pack(fill='x', pady=10, expand=True)
        ills_check = CustomCheckButton(frame_ills, variable=self.ills_select)
        ills_check.pack(side='left', padx=(0, 2))
        ills_button = DoubleTextButton(frame_ills, text_main='Collect Illustrations',
                                       text_sub='Will collect all illustrations associated with CCL',
                                       command=lambda: self.raiseframe(Illustrations))
        ills_button.pack(side='left', fill='x', expand=True)

        run_button = ModernButton(self.mainframe, text='Press to Run', height=1, command=self.run)
        run_button.pack(expand=True, fill='x', pady=5)

    def settings_frame(self):
        self.settingsframe = tk.Frame(self.settings_notebook, bg='white')
        self.settingsframe.pack(anchor='center', pady=10)
        sharedprocess = ttk.Checkbutton(self.settingsframe,
                                        text='Share input between process',
                                        variable=self.shared)
        sharedprocess.pack(expand=True, pady=5, anchor='w')
        headless = ttk.Checkbutton(self.settingsframe,
                                   text='Enable headless Chrome',
                                   variable=self.headless)
        headless.pack(expand=True, pady=5, anchor='w')
        incomplete = ttk.Checkbutton(self.settingsframe,
                                     text='Allow run with incomplete fields',
                                     variable=self.incomplete_run)
        incomplete.pack(expand=True, pady=5, anchor='w')

        cpuframe = tk.Frame(self.settingsframe, background='white')
        cpuframe.pack(pady=5)
        cpu_label = ttk.Label(cpuframe, text='Enter desired CPU usage (0 for default): ')
        cpu_label.pack(side='left')
        self.cpu_entry = ttk.Entry(cpuframe, width=3)
        self.cpu_entry.insert(tk.END, 0)
        self.cpu_entry.pack(side='left')
        self.cpu_entry.bind('<FocusOut>', self._cpu_usage)
        percent = ttk.Label(cpuframe, text='%')
        percent.pack(side='left')

    def _cpu_usage(self, e):
        cores = os.cpu_count()
        try:
            cpu_usage = int(self.cpu_entry.get())
            if cpu_usage < 0 or cpu_usage > 100:
                self.invalid_input()
            elif cpu_usage == 0:
                self.controller.ccl.processes = 1
            else:
                self.controller.ccl.processes = round(cpu_usage / 100 * cores)
        except ValueError:
            self.invalid_input()

    @staticmethod
    def invalid_input():
        messagebox.showerror('Error', 'Invalid Input')

    def extra_frame(self):
        self.extraframe = tk.Frame(self.extra_notebook, bg='white')
        self.extraframe.pack(anchor='center', expand=True, fill='y')

        self.rohsframe = tk.Frame(self.extraframe, bg='#7093db')
        self.rohsframe.pack(pady=10, fill='x', expand=True)
        rohs = DoubleTextButton(self.rohsframe,
                                text_main='RoHS Bill of Materials Comparison',
                                text_sub='Output a delta report between two BOMS',
                                command=lambda: self.raiseframe_extra(ROHSCompare))
        rohs.pack(fill='x', expand=True, side='right', padx=(4, 0))

        self.filterframe = tk.Frame(self.extraframe, bg='#7093db')
        self.filterframe.pack(pady=10, fill='x', expand=True)
        filtercheck = DoubleTextButton(self.filterframe,
                                       text_main='Format Checker',
                                       text_sub='Will output filtered CCL to check CCL format',
                                       command=lambda: self.raiseframe_extra(FilterCompare))
        filtercheck.pack(fill='x', expand=True, side='right', padx=(4, 0))

        self.illtoolframe = tk.Frame(self.extraframe, bg='#7093db')
        self.illtoolframe.pack(pady=10, fill='x', expand=True)
        illustration_tool = DoubleTextButton(self.illtoolframe,
                                             text_main='Illustration Tool',
                                             text_sub='Used to insert and delete illustrations',
                                             command=lambda: self.raiseframe_extra(InsertDelIllustration))
        illustration_tool.pack(fill='x', expand=True, side='right', padx=(4, 0))

    def raiseframe(self, name):
        self.mainframe.forget()
        frame = self.frames[name]
        frame.pack(expand=True, fill='both', padx=10)
        frame.update()
        frame.event_generate('<<ShowFrame>>')

    def back(self, ontop):
        self.frames[ontop].forget()
        self.mainframe.pack(anchor='center', expand=True, fill='y')

    def raiseframe_extra(self, name):
        self.extraframe.forget()
        frame = self.frames[name]
        frame.pack(expand=True, fill='both', padx=10)
        frame.update()
        frame.event_generate('<<ShowFrame>>')

    def back_extra(self, ontop):
        self.frames[ontop].forget()
        self.extraframe.pack(anchor='center', expand=True, fill='y')

    def run(self):
        bom_compare = [self.bom_compare_old, self.bom_compare_new, self.bom_compare_save]
        ccl_update = [self.ccl_update_old, self.ccl_update_new, self.ccl_update_loc, self.ccl_update_save_loc]
        document = [self.docs_ccl, self.docs_paths, self.docs_savedir, self.docs_user, self.docs_pass]
        illustration = [self.ill_ccl, self.ill_cclsave, self.ill_save, self.ill_scan]
        full = True
        if not self.incomplete_run.get():

            if self.compare_select.get():
                for i in bom_compare:
                    if i is None:
                        full = False
                        messagebox.showerror(title='Missing Info',
                                             message='Missing Info in BOM Compare page')
                        break
            elif self.update_select.get():
                for i in ccl_update:
                    if i is None:
                        full = False
                        messagebox.showerror(title='Missing Info',
                                             message='Missing Info in CCL Update page')
                        break
            elif self.docs_select.get():
                for i in document:
                    if i is None:
                        full = False
                        messagebox.showerror(title='Missing Info',
                                             message='Missing Info in Document Collection page')
                        break
            elif self.ills_select.get():
                for i in illustration:
                    if i is None:
                        full = False
                        messagebox.showerror(title='Missing Info',
                                             message='Missing Info in Illustration Collection page')
                        break
        if full and (self.compare_select.get() or self.update_select.get() or
                     self.docs_select.get() or self.ills_select.get()):
            Run(self)


class BomCompare(tk.Frame):

    def __init__(self, root, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.root = root
        self['background'] = self.master['background']

        self.oldframe = tk.Frame(self, background=self.master['background'])
        self.oldframe.pack(expand=True, fill='x', pady=5)
        self.old_button = ModernButton(self.oldframe,
                                       text='Browse for Old BOM',
                                       width=20,
                                       command=self.oldsave)
        self.old_button.pack(side='left', anchor='center', padx=(0, 2))
        self.old_entry = ModernEntry(self.oldframe)
        self.old_entry.pack(side='right', fill='both', expand=True, anchor='center')

        self.newframe = tk.Frame(self, background=self.master['background'])
        self.newframe.pack(expand=True, fill='x', pady=5)
        self.new_button = ModernButton(self.newframe,
                                       text='Browse for New BOM',
                                       width=20,
                                       command=self.newsave)
        self.new_button.pack(side='left', anchor='center', padx=(0, 2))
        self.new_entry = ModernEntry(self.newframe)
        self.new_entry.pack(side='right', fill='both', expand=True, anchor='center')

        self.saveframe = tk.Frame(self, background=self.master['background'])
        self.saveframe.pack(expand=True, fill='x', pady=5)
        self.save_button = ModernButton(self.saveframe,
                                        text='Browse For Save Location',
                                        width=20,
                                        command=self.zipsave)
        self.save_button.pack(side='left', anchor='center', padx=(0, 2))
        self.save_entry = ModernEntry(self.saveframe)
        self.save_entry.pack(side='right', fill='both', expand=True, anchor='center')

        self.back_button = ModernButton(self, text='Back', width=20, command=lambda: self.root.back(BomCompare))
        self.back_button.pack(pady=5)

    def oldsave(self):
        filename = filedialog.askopenfile(initialdir=self.root.cache_dir,
                                          title='Select Old AVL Multilevel BOM',
                                          filetypes=[('Comma-Separated Values', '.csv')])
        self.old_entry.clear()
        self.old_entry.insert(tk.END, filename.name)
        self.root.cache_dir = filename
        self.root.bom_compare_old = filename

    def newsave(self):
        filename = filedialog.askopenfile(initialdir=self.root.cache_dir,
                                          title='Select New AVL Multilevel BOM',
                                          filetypes=[('Comma-Separated Values', '.csv')])
        self.new_entry.clear()
        self.new_entry.insert(tk.END, filename.name)
        self.root.cache_dir = filename
        self.root.bom_compare_new = filename

    def zipsave(self):
        filename = filedialog.asksaveasfilename(initialdir=self.root.cache_dir,
                                                title='Save As',
                                                filetypes=[('Zip', '.zip')],
                                                defaultextension='')
        self.save_entry.clear()
        self.save_entry.insert(tk.END, filename)
        self.root.cache_dir = filename
        self.root.bom_compare_save = filename


class UpdateCCL(tk.Frame):
    def __init__(self, root, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.root = root
        self.bind('<<ShowFrame>>', self.sharedvar)

        self['background'] = self.master['background']
        self.oldframe = tk.Frame(self, background=self.master['background'])
        self.oldframe.pack(expand='True', fill='x', pady=5)
        self.old_button = ModernButton(self.oldframe,
                                       text='Browse for Old BOM',
                                       width=20,
                                       command=self.oldsave)
        self.old_button.pack(side='left', anchor='center', padx=(0, 2))
        self.old_entry = ModernEntry(self.oldframe)
        self.old_entry.pack(side='right', fill='both', expand=True, anchor='center')

        self.newframe = tk.Frame(self, background=self.master['background'])
        self.newframe.pack(expand='True', fill='x', pady=5)
        self.new_button = ModernButton(self.newframe,
                                       text='Browse for New BOM',
                                       width=20,
                                       command=self.newsave)
        self.new_button.pack(side='left', anchor='center', padx=(0, 2))
        self.new_entry = ModernEntry(self.newframe)
        self.new_entry.pack(side='right', fill='both', expand=True, anchor='center')

        self.cclframe = tk.Frame(self, background=self.master['background'])
        self.cclframe.pack(expand=True, fill='x', pady=5)
        self.ccl_button = ModernButton(self.cclframe,
                                       text='Browse for CCL Location',
                                       width=20,
                                       command=self.ccl_open)
        self.ccl_button.pack(side='left', anchor='center', padx=(0, 2))
        self.ccl_entry = ModernEntry(self.cclframe)
        self.ccl_entry.pack(side='right', fill='both', expand=True, anchor='center')

        self.saveframe = tk.Frame(self, background=self.master['background'])
        self.saveframe.pack(expand='True', fill='x', pady=5)
        self.save_button = ModernButton(self.saveframe,
                                        text='Browse for Save Location',
                                        width=20,
                                        command=self.cclsave)
        self.save_button.pack(side='left', anchor='center', padx=(0, 2))
        self.save_entry = ModernEntry(self.saveframe)
        self.save_entry.pack(side='right', fill='both', expand=True, anchor='center')

        self.back_button = ModernButton(self, text='Back', width=20, command=lambda: self.root.back(UpdateCCL))
        self.back_button.pack(pady=5)

    def oldsave(self):
        filename = filedialog.askopenfile(initialdir=self.root.cache_dir,
                                          title='Select Old AVL Multilevel BOM',
                                          filetypes=[('Comma-Separated Values', '.csv')])
        self.old_entry.clear()
        self.old_entry.insert(tk.END, filename.name)
        self.root.cache_dir = filename
        self.root.ccl_update_old = filename

    def newsave(self):
        filename = filedialog.askopenfile(initialdir=self.root.cache_dir,
                                          title='Select New AVL Multilevel BOM',
                                          filetypes=[('Comma-Separated Values', '.csv')])
        self.new_entry.clear()
        self.new_entry.insert(tk.END, filename.name)
        self.root.cache_dir = filename
        self.root.ccl_update_new = filename

    def ccl_open(self):
        filename = filedialog.askopenfile(initialdir=self.root.cache_dir,
                                          title='Select CCL',
                                          filetypes=[('Word Document', '.docx')])
        self.ccl_entry.clear()
        self.ccl_entry.insert(tk.END, filename.name)
        self.root.cache_dir = filename
        self.root.ccl_update_loc = filename

    def cclsave(self):
        filename = filedialog.asksaveasfilename(initialdir=self.root.cache_dir,
                                                title='Save As',
                                                filetypes=[('Word Document', '.docx')],
                                                defaultextension='.docx')
        self.save_entry.clear()
        self.save_entry.insert(tk.END, filename)
        self.root.cache_dir = filename
        self.root.ccl_update_save_loc = filename

    def sharedvar(self, e):
        if self.root.shared.get() and self.root.ccl_update_old is None:
            try:
                self.old_entry.insert(tk.END, self.root.bom_compare_old.name)
                self.root.ccl_update_old = self.root.bom_compare_old
            except AttributeError:
                self.old_entry.clear()
                self.root.ccl_update_old = None

        if self.root.shared.get() and self.root.ccl_update_new is None:
            try:
                self.new_entry.insert(tk.END, self.root.bom_compare_new.name)
                self.root.ccl_update_new = self.root.bom_compare_new
            except AttributeError:
                self.new_entry.clear()
                self.root.ccl_update_new = None


class DocumentCollector(tk.Frame):
    def __init__(self, root, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.root = root
        self['background'] = self.master['background']
        self.bind('<<ShowFrame>>', self.sharedvar)

        self.cclframe = tk.Frame(self, background=self.master['background'])
        self.cclframe.pack(expand=False, fill='x', pady=5)
        self.ccl_button = ModernButton(self.cclframe,
                                       text='Browse for CCL Location',
                                       width=20,
                                       command=self.ccl_open)
        self.ccl_button.pack(side='left', anchor='center', padx=(0, 2))
        self.ccl_entry = ModernEntry(self.cclframe)
        self.ccl_entry.pack(side='right', fill='both', expand=True, anchor='center')

        self.saveframe = tk.Frame(self, background=self.master['background'])
        self.saveframe.pack(expand=False, fill='x', pady=5)
        self.save_button = ModernButton(self.saveframe,
                                        text='Select Save Folder',
                                        width=20,
                                        command=self.docsave)
        self.save_button.pack(side='left', anchor='center', padx=(0, 2))
        self.save_entry = ModernEntry(self.saveframe)
        self.save_entry.pack(side='right', fill='both', expand=True, anchor='center')

        self.enoviaframe = tk.Frame(self, background=self.master['background'])
        self.enoviaframe.pack(expand=False, fill='x', pady=5)
        self.user = ModernEntry(self.enoviaframe, text='Enter Enovia Username')
        self.user.label.bind("<Button-1>", self.clear_user)
        self.user.pack(expand=True, fill='both', side='left', padx=(0, 2))
        self.password = ModernEntry(self.enoviaframe, text='Enter Enovia Password')
        self.password.label.bind("<Button-1>", self.clear_password)
        self.password.pack(expand=True, fill='both', side='left', padx=(0, 2))
        self.login = ModernButton(self.enoviaframe, text='Login', width=9, command=self.start_thread)
        self.login.pack(side='right', fill='x')

        self.path_checks()
        self.back_button = ModernButton(self, text='Back', width=20,
                                        command=lambda: self.root.back(DocumentCollector))
        self.back_button.pack(pady=5)

    def clear_user(self, e):
        if self.user.get() == 'Enter Enovia Username':
            self.user.clear()

    def clear_password(self, e):
        self.password.label.config(show='*')
        if self.password.get() == 'Enter Enovia Password':
            self.password.clear()

    def ccl_open(self):
        filename = filedialog.askopenfile(initialdir=self.root.cache_dir,
                                          title='Select CCL',
                                          filetypes=[('Word Document', '.docx')])
        self.ccl_entry.clear()
        self.ccl_entry.insert(tk.END, filename.name)
        self.root.cache_dir = filename
        self.root.docs_ccl = filename

    def docsave(self):
        filename = filedialog.askdirectory(initialdir=self.root.cache_dir,
                                           title='Select Folder')
        self.save_entry.clear()
        self.save_entry.insert(tk.END, filename)
        self.root.cache_dir = filename
        self.root.docs_savedir = filename

    def path_checks(self):
        centerframe = tk.Frame(self, background='white')
        centerframe.pack(expand=True, fill='both')

        self.path_listbox = tk.Listbox(centerframe, height=1)
        self.path_listbox.pack(side='left', fill='both', expand=True)
        scroll = tk.Scrollbar(centerframe, orient='vertical', command=self.path_listbox.yview)
        scroll.pack(side='left', fill='y')
        self.path_listbox.configure(yscrollcommand=scroll.set)

        addpath = ModernButton(centerframe, text='Add Path', command=self.add_path, width=9)
        addpath.pack(pady=5, padx=(5, 0))
        delpath = ModernButton(centerframe, text='Delete Path', command=self.del_path, width=9)
        delpath.pack(pady=5, padx=(5, 0))
        moveup = ModernButton(centerframe, text='Move Up', command=self.move_up, width=9)
        moveup.pack(pady=5, padx=(5, 0))
        movedown = ModernButton(centerframe, text='Move Down', command=self.move_down, width=9)
        movedown.pack(pady=5, padx=(5, 0))

    def add_path(self):
        filename = filedialog.askdirectory(initialdir='/', title='Select Directory')
        self.path_listbox.insert(tk.END, filename)
        self.set_check_paths()

    def del_path(self):
        self.path_listbox.delete(self.path_listbox.curselection())
        self.set_check_paths()

    def move_up(self):
        selected = self.path_listbox.curselection()[0]
        text = self.path_listbox.get(selected)
        self.path_listbox.delete(selected)
        self.path_listbox.insert(selected - 1, text)
        self.path_listbox.select_set(selected - 1)
        self.set_check_paths()

    def move_down(self):
        selected = self.path_listbox.curselection()[0]
        text = self.path_listbox.get(selected)
        self.path_listbox.delete(selected)
        self.path_listbox.insert(selected + 1, text)
        self.path_listbox.select_set(selected + 1)
        self.set_check_paths()

    def set_check_paths(self):
        if self.path_listbox.size() > 0:
            self.root.docs_paths = [self.path_listbox.get(idx) for idx in range(self.path_listbox.size())]
        else:
            self.root.docs_paths = []

    def sharedvar(self, e):
        if self.root.shared.get() and self.root.docs_ccl is None:
            try:
                self.ccl_entry.insert(tk.END, self.root.ccl_update_loc.name)
                self.root.docs_ccl = self.root.ccl_update_loc
            except AttributeError:
                self.ccl_entry.clear()
                self.root.docs_ccl = None

    def enoviacheck(self):
        enovia = Enovia(self.user.get(), self.password.get(), headless=self.root.headless.get())
        try:
            enovia.create_env()
        except UnexpectedAlertPresentException:
            messagebox.showerror(title='Error', message='Invalid username or password')
            raise KeyError('Invalid username or password')
        except Exception as e:
            messagebox.showerror(title='Error', message=f'Error {e} has occured')
            raise e
        else:
            self.root.docs_user = self.user.get()
            self.root.docs_pass = self.password.get()
            messagebox.Message(title='Success', message='Login Successful')

    def start_thread(self):
        self.thread = threading.Thread(target=self.enoviacheck)

        self.progressframe = tk.Toplevel(self, background='white')
        self.progressframe.lift()
        self.progressframe.focus_force()
        self.progressframe.grab_set()
        self.progressframe.resizable(False, False)
        self.progressframe.minsize(width=200, height=50)
        progressbar = ttk.Progressbar(self.progressframe, mode='indeterminate', length=200)
        progressbar.pack(pady=(10, 0), padx=5)
        progressbar.start(10)
        progresslabel = tk.Label(self.progressframe, text='Logging into Enovia', background='white')
        progresslabel.pack(pady=(0, 10))

        self.thread.daemon = True
        self.thread.start()
        self.after(20, self.check_thread)

    def check_thread(self):
        if self.thread.is_alive():
            self.after(20, self.check_thread)
        else:
            self.progressframe.destroy()


class Illustrations(tk.Frame):
    def __init__(self, root, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.root = root
        self.bind('<<ShowFrame>>', self.sharedvar)
        self['background'] = self.master['background']

        self.cclframe = tk.Frame(self, background=self.master['background'])
        self.cclframe.pack(expand=True, fill='x', pady=5)
        self.ccl_button = ModernButton(self.cclframe,
                                       text='Browse for CCL Location',
                                       width=20,
                                       command=self.ccl_open)
        self.ccl_button.pack(side='left', anchor='center', padx=(0, 2))
        self.ccl_entry = ModernEntry(self.cclframe)
        self.ccl_entry.pack(side='right', fill='both', expand=True, anchor='center')

        self.saveframe = tk.Frame(self, background=self.master['background'])
        self.saveframe.pack(expand='True', fill='x', pady=5)
        self.save_button = ModernButton(self.saveframe,
                                        text='New CCL Save Location',
                                        width=20,
                                        command=self.cclsave)
        self.save_button.pack(side='left', anchor='center', padx=(0, 2))
        self.save_entry = ModernEntry(self.saveframe)
        self.save_entry.pack(side='right', fill='both', expand=True, anchor='center')

        self.saveframe_doc = tk.Frame(self, background=self.master['background'])
        self.saveframe_doc.pack(expand=True, fill='x', pady=5)
        self.save_button_doc = ModernButton(self.saveframe_doc,
                                            text='Document Scan Folder',
                                            width=20,
                                            command=self.docsave)
        self.save_button_doc.pack(side='left', anchor='center', padx=(0, 2))
        self.save_entry_doc = ModernEntry(self.saveframe_doc)
        self.save_entry_doc.pack(side='right', fill='both', expand=True, anchor='center')

        self.illsave_frame = tk.Frame(self, background=self.master['background'])
        self.illsave_frame.pack(expand='True', fill='x', pady=5)
        self.illsave_button = ModernButton(self.illsave_frame,
                                           text='Illustration Save',
                                           width=20,
                                           command=self.illsave)
        self.illsave_button.pack(side='left', anchor='center', padx=(0, 2))
        self.illsave_entry = ModernEntry(self.illsave_frame)
        self.illsave_entry.pack(side='right', fill='both', expand=True, anchor='center')

        self.back_button = ModernButton(self, text='Back', width=20, command=lambda: self.root.back(Illustrations))
        self.back_button.pack(pady=5)

    def ccl_open(self):
        filename = filedialog.askopenfile(initialdir=self.root.cache_dir,
                                          title='Select CCL',
                                          filetypes=[('Word Document', '.docx')])
        self.ccl_entry.clear()
        self.ccl_entry.insert(tk.END, filename.name)
        self.root.cache_dir = filename
        self.root.ill_ccl = filename

    def cclsave(self):
        filename = filedialog.asksaveasfilename(initialdir=self.root.cache_dir,
                                                title='Save As',
                                                filetypes=[('Word Document', '.docx')],
                                                defaultextension='.docx')
        self.save_entry.clear()
        self.save_entry.insert(tk.END, filename)
        self.root.cache_dir = filename
        self.root.ill_cclsave = filename

    def docsave(self):
        filename = filedialog.askdirectory(initialdir=self.root.cache_dir,
                                           title='Select Folder')
        self.save_entry.clear()
        self.save_entry.insert(tk.END, filename)
        self.root.cache_dir = filename
        self.root.ill_scan = filename

    def illsave(self):
        filename = filedialog.askdirectory(initialdir=self.root.cache_dir,
                                           title='Select Folder')
        self.illsave_entry.clear()
        self.illsave_entry.insert(tk.END, filename)
        self.root.cache_dir = filename
        self.root.ill_save = filename

    def sharedvar(self, e):
        if self.root.shared.get() and self.root.ill_ccl is None and self.root.docs_ccl is not None:
            try:
                self.ccl_entry.insert(tk.END, self.root.docs_ccl.name)
                self.root.ill_ccl = self.root.docs_ccl
            except AttributeError:
                self.ccl_entry.clear()
                self.root.ill_ccl = None

        elif self.root.shared.get() and self.root.ill_ccl is None and self.root.ccl_update_loc is not None:
            try:
                self.ccl_entry.insert(tk.END, self.root.ccl_update_loc.name)
                self.root.ill_ccl = self.root.ccl_update_loc
            except AttributeError:
                self.ccl_entry.clear()
                self.root.ill_ccl = None

        if self.root.shared.get() and self.root.ill_cclsave is None:
            try:
                self.save_entry.insert(tk.END, self.root.ccl_update_save_loc.name)
                self.root.ccl_update_new = self.root.ccl_update_save_loc
            except AttributeError:
                self.save_entry.clear()
                self.root.ccl_update_new = None

        if self.root.shared.get() and self.root.ill_scan is None:
            try:
                self.save_entry_doc.insert(tk.END, self.root.docs_savedir.name)
                self.root.ill_scan = self.root.docs_savedir
            except AttributeError:
                self.save_entry_doc.clear()
                self.root.ill_scan = None


class Run(tk.Toplevel):
    def __init__(self, root, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.root = root
        self.ccl = CCL()
        self.ccl.processes = self.root.processes
        self.prev_prog = None

        self.total_progress()
        self.prompt()
        self.controls()

        sys.stdout = TextRedirector(self.console, 'stdout')
        sys.stderr = TextRedirector(self.console, 'stderr')

        self.update()
        self.minsize(self.winfo_width(), self.winfo_height())

    def controls(self):
        framecrtl = tk.Frame(self)
        framecrtl.pack()

        run = ModernButton(framecrtl, text='Run', command=self.start_threading, width=5)
        run.pack(side='left', padx=5, pady=5)

        cancel = ModernButton(framecrtl, text='Cancel', command=self.destroy, width=5)
        cancel.pack(side='left', padx=5, pady=5)

        abort = ModernButton(framecrtl, text='Abort', command=self.root.destroy, width=5)
        abort.pack(side='left', padx=5, pady=5)

    def prompt(self):
        promptframe = tk.Frame(self)
        promptframe.pack(expand=True, fill=tk.BOTH, padx=5)

        self.progressbar = ttk.Progressbar(promptframe, mode='determinate', maximum=progressbar.total)
        self.progressbar.pack(fill='x', pady=(5, 0))

        self.progress_label = tk.Label(promptframe, text='Press Run to begin')
        self.progress_label.pack(anchor='w', pady=(0, 5))

        self.console = tk.Text(promptframe, wrap='word')
        self.console.pack(side='left', expand=True, fill=tk.BOTH)

        scroll = tk.Scrollbar(promptframe, orient='vertical', command=self.console.yview)
        scroll.pack(side='right', fill='y')
        self.console.configure(yscrollcommand=scroll.set)

    def run(self):
        if self.root.compare_select.get():
            print('Starting BOM Compare')
            self.ccl.set_bom_compare(self.root.bom_compare_old.name, self.root.bom_compare_new.name)
            self.ccl.save_compare(self.root.bom_compare_save)
            progressbar.add_current(1)
            print('BOM Compare finished')

        if self.root.update_select.get():
            print('Starting to update the CCL')
            self.ccl.ccl_docx = self.root.ccl_update_loc
            self.ccl.set_bom_compare(self.root.ccl_update_old.name, self.root.ccl_update_new.name)
            self.ccl.update_ccl(self.root.ccl_update_save_loc)
            print('CCL Has been updated and saved')
            progressbar.add_current(1)

        if self.root.docs_select.get():
            print('Collecting Documents')
            self.ccl.ccl_docx = self.root.docs_ccl
            self.ccl.path_checks = self.root.docs_paths
            self.ccl.path_ccl_data = self.root.docs_savedir
            self.ccl.username = self.root.docs_user
            self.ccl.password = self.root.docs_pass
            self.ccl.collect_documents(headless=self.root.headless.get())
            print('Documents have been successfully collected')
            # Progressbar progress will be updated in the filehandler module

        if self.root.ills_select.get():
            print('Starting to Collect Illustrations')
            self.ccl.ccl_docx = self.root.ill_ccl
            self.ccl.path_ccl_data = self.root.ill_scan
            self.ccl.path_illustration = self.root.ill_save
            self.ccl.collect_illustrations()
            self.ccl.insert_illustration_data(self.root.ill_cclsave)
            print('Illustrations have been collected and CCL has been updated')
            # Progressbar progress will be updated in the CCL module

        self.progressbar['value'] = progressbar.total
        self.progress_label.config(text='Done')
        print('FINISHED!')

    def total_progress(self):
        progressbar.reset()
        if self.root.compare_select.get():
            progressbar.add_total(1)
        if self.root.update_select.get():
            progressbar.add_total(1)
        if self.root.docs_select.get():
            progressbar.add_total(2)
        if self.root.ills_select.get():
            progressbar.add_total(1)

    def start_threading(self):
        self.progress_label.config(text='Estimating Time Reamining')
        self.prev_prog = progressbar.current
        self.submit_thread = threading.Thread(target=self.run)
        self.start_time = time.time()
        self.submit_thread.daemon = True
        self.submit_thread.start()
        self.after(1000, self.check_thread)

    def check_thread(self):
        if self.submit_thread.is_alive():
            if self.prev_prog != progressbar.current:
                self.time_remaining()
            self.after(1000, self.check_thread)

    def time_remaining(self):
        elapsed_time = time.time() - self.start_time
        self.progressbar['value'] = progressbar.current
        time_remaining = round((1 - progressbar.current) * elapsed_time)
        if time_remaining < 60:
            self.progress_label.config(text=f'Estimated Time Remaining: {time_remaining} seconds')
        elif 3600 > time_remaining > 60:
            time_remaining = round(time_remaining / 60)
            self.progress_label.config(text=f'Estimated TIme Remaining: {time_remaining} minutes')
        elif time_remaining > 3600:
            time_remaining = dt.timedelta(seconds=time_remaining)
            self.progress_label.config(text=f'Estimated Time Remaining: {time_remaining}')


class TextRedirector(object):
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, str):
        self.widget.configure(state="normal")
        self.widget.insert("end", str, (self.tag,))
        self.widget.see(tk.END)
        self.widget.configure(state="disabled")


class ROHSCompare(tk.Frame):
    def __init__(self, root, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.bom_a = None
        self.bom_b = None
        self.root = root
        self['background'] = self.master['background']

        self.aframe = tk.Frame(self, background=self.master['background'])
        self.aframe.pack(expand=True, fill='x', pady=5)
        self.a_button = ModernButton(self.aframe,
                                     text='Browse for BOM A',
                                     width=20,
                                     command=self.boma)
        self.a_button.pack(side='left', anchor='center', padx=(0, 2))
        self.a_entry = ModernEntry(self.aframe)
        self.a_entry.pack(side='right', fill='both', expand=True, anchor='center')

        self.bframe = tk.Frame(self, background=self.master['background'])
        self.bframe.pack(expand=True, fill='x', pady=5)
        self.b_button = ModernButton(self.bframe,
                                     text='Browse for BOM B',
                                     width=20,
                                     command=self.bomb)
        self.b_button.pack(side='left', anchor='center', padx=(0, 2))
        self.b_entry = ModernEntry(self.bframe)
        self.b_entry.pack(side='right', fill='both', expand=True, anchor='center')

        self.saveframe = tk.Frame(self, background=self.master['background'])
        self.saveframe.pack(expand=True, pady=5)
        self.exclusivea_button = ModernButton(self.saveframe,
                                              text='Generate Exclusive to BOM A',
                                              width=25,
                                              command=lambda: self.start_threading(self.bom_a, self.bom_b))
        self.exclusivea_button.pack(side='left', anchor='center', padx=(0, 2))
        self.exclusiveb_button = ModernButton(self.saveframe,
                                              text='Generate Exclusive to BOM B',
                                              width=25,
                                              command=lambda: self.start_threading(self.bom_b, self.bom_a))
        self.exclusiveb_button.pack(side='left', anchor='center', padx=(0, 2))

        self.back_button = ModernButton(self, text='Back', width=20, command=lambda: self.root.back_extra(ROHSCompare))
        self.back_button.pack(pady=5)

    def boma(self):
        filename = filedialog.askopenfile(initialdir=self.root.cache_dir,
                                          title='Select AVL Multilevel BOM',
                                          filetypes=[('Comma-Separated Values', '.csv')])
        self.a_entry.clear()
        self.a_entry.insert(tk.END, filename.name)
        self.root.cache_dir = filename
        self.bom_a = filename.name

    def bomb(self):
        filename = filedialog.askopenfile(initialdir=self.root.cache_dir,
                                          title='Select AVL Multilevel BOM',
                                          filetypes=[('Comma-Separated Values', '.csv')])
        self.b_entry.clear()
        self.b_entry.insert(tk.END, filename.name)
        self.root.b_entry = filename
        self.bom_b = filename.name

    def compare(self, a_avl, b_avl):
        a_avl = self.read_avl(a_avl, 0)
        b_avl = self.read_avl(b_avl, 0)

        a_pns, b_pns = set(a_avl['Name']), set(b_avl['Name'])
        exclusive = {pn for pn in a_pns if pn not in b_pns}
        df_exclusive = pd.DataFrame()
        for pn in exclusive:
            df_exclusive = pd.concat([a_avl.loc[a_avl['Name'] == pn], df_exclusive])

        sf = StyleFrame(a_avl)
        style = Styler(bg_color='yellow',
                       border_type=None,
                       shrink_to_fit=False,
                       wrap_text=False,
                       font='Calibri',
                       font_size=11)
        style_default = Styler(border_type=None,
                               fill_pattern_type=None,
                               shrink_to_fit=False,
                               wrap_text=False,
                               font='Calibri',
                               font_size=11)

        for idx in a_avl.index:
            sf.apply_style_by_indexes(sf.index[idx], styler_obj=style_default)

        for idx in df_exclusive.index:
            sf.apply_style_by_indexes(sf.index[idx], styler_obj=style)
        self.save_as(sf)

    @staticmethod
    def read_avl(path, skiprow):
        df = pd.read_csv(path, skiprows=skiprow)
        try:
            df['Name']
        except KeyError:
            skiprow += 1
            if skiprow < 10:
                df = CCL.read_avl(path, skiprow)
            else:
                raise TypeError('File is in wrong format')
        return df

    def save_as(self, sf):
        filename = filedialog.asksaveasfilename(initialdir=self.root.cache_dir,
                                                title='Save As',
                                                filetypes=[('Excel', '.xlsx')],
                                                defaultextension='.xlsx')
        self.root.cache_dir = filename
        sf.to_excel(filename).save()

    def show_progressbar(self):
        self.progressframe = tk.Toplevel(self, background='white')
        self.progressframe.lift()
        self.progressframe.focus_force()
        self.progressframe.grab_set()
        self.progressframe.resizable(False, False)
        self.progressframe.minsize(width=200, height=50)
        progressbar = ttk.Progressbar(self.progressframe, mode='indeterminate', length=200)
        progressbar.pack(pady=(10, 0), padx=5)
        progressbar.start(10)
        progresslabel = tk.Label(self.progressframe, text='Generating BOM Comparison', background='white')
        progresslabel.pack(pady=(0, 10))

    def start_threading(self, a_avl, b_avl):
        self.show_progressbar()
        self.submit_thread = threading.Thread(target=lambda: self.compare(a_avl, b_avl))
        self.submit_thread.daemon = True
        self.submit_thread.start()
        self.after(20, self.check_thread)

    def check_thread(self):
        if self.submit_thread.is_alive():
            self.after(20, self.check_thread)
        else:
            self.progressframe.destroy()


class FilterCompare(tk.Frame):
    def __init__(self, root, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.ccl_path = None
        self.root = root
        self['background'] = self.master['background']

        self.cclframe = tk.Frame(self, background=self.master['background'])
        self.cclframe.pack(expand=True, fill='x', pady=5)
        self.ccl_button = ModernButton(self.cclframe,
                                       text='Browse for CCL',
                                       width=20,
                                       command=self.ccl_file)
        self.ccl_button.pack(side='left', anchor='center', padx=(0, 2))
        self.ccl_entry = ModernEntry(self.cclframe)
        self.ccl_entry.pack(side='right', fill='both', expand=True, anchor='center')

        self.buttonframe = tk.Frame(self, background=self.master['background'])
        self.buttonframe.pack(expand=True, fill='y', pady=5)

        self.run_button = ModernButton(self.buttonframe,
                                       text='Run Format Checker',
                                       width=20,
                                       command=self.start_thread)
        self.run_button.pack(anchor='center', side='right', padx=(5, 0))

        self.back_button = ModernButton(self.buttonframe, text='Back', width=20,
                                        command=lambda: self.root.back_extra(FilterCompare))
        self.back_button.pack(side='left', padx=(0, 5))

    def ccl_file(self):
        filename = filedialog.askopenfile(initialdir=self.root.cache_dir,
                                          title='Select CCL',
                                          filetypes=[('Word Document', '.docx')])
        self.ccl_entry.clear()
        self.ccl_entry.insert(tk.END, filename.name)
        self.root.cache_dir = filename
        self.ccl_path = filename.name

    def getreport(self):
        filtered = Parser(self.ccl_path).filter()
        filtered['pn'].fillna('CRITIAL MISSING', inplace=True)
        filtered['desc'].fillna('CRITIAL MISSING', inplace=True)
        filtered['fn'].fillna('CRITIAL MISSING', inplace=True)
        filtered.to_csv(self.cclsave())

    def cclsave(self):
        filename = filedialog.asksaveasfilename(initialdir=self.root.cache_dir,
                                                title='Save As',
                                                filetypes=[('Comma-Separated Values', '.csv')],
                                                defaultextension='.csv')
        self.root.cache_dir = filename
        return filename

    def show_progressbar(self):
        self.progressframe = tk.Toplevel(self, background='white')
        self.progressframe.lift()
        self.progressframe.focus_force()
        self.progressframe.grab_set()
        self.progressframe.resizable(False, False)
        self.progressframe.minsize(width=200, height=50)
        progressbar = ttk.Progressbar(self.progressframe, mode='indeterminate', length=200)
        progressbar.pack(pady=(10, 0), padx=5)
        progressbar.start(10)
        progresslabel = tk.Label(self.progressframe, text='Generating filtered CCL', background='white')
        progresslabel.pack(pady=(0, 10))

    def start_thread(self):
        self.show_progressbar()
        self.submit_thread = threading.Thread(target=self.getreport)
        self.submit_thread.daemon = True
        self.submit_thread.start()
        self.after(20, self.check_thread)

    def check_thread(self):
        if self.submit_thread.is_alive():
            self.after(20, self.check_thread)
        else:
            self.progressframe.destroy()


class InsertDelIllustration(tk.Frame):
    def __init__(self, root, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.root = root
        self['background'] = self.master['background']
        self.ccl = CCL()
        self.new_ccl = None
        self.illustration = None
        self.illustration_dir = None
        self.ill_num = None

        self.cclframe = tk.Frame(self, background=self.master['background'])
        self.cclframe.pack(expand=True, fill='x', pady=5)
        self.ccl_button = ModernButton(self.cclframe,
                                       text='Browse for CCL',
                                       width=20,
                                       command=self.ccl_open)
        self.ccl_button.pack(side='left', anchor='center', padx=(0, 2))
        self.ccl_entry = ModernEntry(self.cclframe)
        self.ccl_entry.pack(side='right', fill='both', expand=True, anchor='center')

        self.saveframe = tk.Frame(self, background=self.master['background'])
        self.saveframe.pack(expand='True', fill='x', pady=5)
        self.save_button = ModernButton(self.saveframe,
                                        text='New CCL Save Location',
                                        width=20,
                                        command=self.cclsave)
        self.save_button.pack(side='left', anchor='center', padx=(0, 2))
        self.save_entry = ModernEntry(self.saveframe)
        self.save_entry.pack(side='right', fill='both', expand=True, anchor='center')

        self.illframe = tk.Frame(self, background=self.master['background'])
        self.illframe.pack(expand=True, fill='x', pady=5)
        self.ill_button = ModernButton(self.illframe,
                                       text='Browse for Illustration',
                                       width=20,
                                       command=self.ccl_open)
        self.ill_button.pack(side='left', anchor='center', padx=(0, 2))
        self.ill_entry = ModernEntry(self.illframe)
        self.ill_entry.pack(side='right', fill='both', expand=True, anchor='center')

        self.illdirframe = tk.Frame(self, background=self.master['background'])
        self.illdirframe.pack(expand=True, fill='x', pady=5)
        self.illdir_button = ModernButton(self.illdirframe,
                                          text='Illustration Save Location',
                                          width=20,
                                          command=self.ccl_open)
        self.illdir_button.pack(side='left', anchor='center', padx=(0, 2))
        self.illdir_entry = ModernEntry(self.illdirframe)
        self.illdir_entry.pack(side='right', fill='both', expand=True, anchor='center')

        self.runframe = tk.Frame(self, background=self.master['background'])
        self.runframe.pack(expand=True, pady=5)
        self.insert_button = ModernButton(self.runframe,
                                          text='Insert Illustration',
                                          width=25,
                                          command=self.insertcmd)
        self.insert_button.pack(side='left', anchor='center', padx=(0, 2))
        self.delete_button = ModernButton(self.runframe,
                                          text='Delete Illustration',
                                          width=25,
                                          command=self.delcmd)
        self.delete_button.pack(side='left', anchor='center', padx=(0, 2))

        self.back_button = ModernButton(self, text='Back', width=20,
                                        command=lambda: self.root.back_extra(InsertDelIllustration))
        self.back_button.pack(pady=5)

    def ccl_open(self):
        filename = filedialog.askopenfile(initialdir=self.root.cache_dir,
                                          title='Select CCL',
                                          filetypes=[('Word Document', '.docx')])
        self.ccl_entry.clear()
        self.ccl_entry.insert(tk.END, filename.name)
        self.root.cache_dir = filename
        self.ccl.ccl_docx = filename

    def cclsave(self):
        filename = filedialog.asksaveasfilename(initialdir=self.root.cache_dir,
                                                title='Save As',
                                                filetypes=[('Word Document', '.docx')],
                                                defaultextension='.docx')
        self.save_entry.clear()
        self.save_entry.insert(tk.END, filename)
        self.root.cache_dir = filename
        self.new_ccl = filename

    def ill_open(self):
        filename = filedialog.askopenfile(initialdir=self.root.cache_dir,
                                          title='Select Illustration')
        self.ccl_entry.clear()
        self.ccl_entry.insert(tk.END, filename.name)
        self.root.cache_dir = filename
        self.illustration = filename

    def ill_dir(self):
        filename = filedialog.askdirectory(initialdir=self.root.cache_dir,
                                           title='Select Directory')
        self.save_entry.clear()
        self.save_entry.insert(tk.END, filename)
        self.root.cache_dir = filename
        self.illustration_dir = filename

    def get_ill_num(self):
        num = re.findall(r'Ill?\s*?.?\s*(\d+)?.?\s*', self.illustration)
        self.ill_num = num[0] if num else None

    def insertcmd(self):
        self.get_ill_num()
        if self.ill_num is not None:
            self.ccl.insert_illustration(self.ill_num, self.illustration, self.new_ccl)
        else:
            messagebox.showerror(title='Error',
                                 message='Illustration number not detected, please check file name')

    def delcmd(self):
        self.get_ill_num()
        if self.ill_num is not None:
            self.ccl.delete_illustration(self.ill_num, self.illustration, self.new_ccl)
        else:
            messagebox.showerror(title='Error',
                                 message='Illustration number not detected, please check file name')

class ModernEntry(tk.Frame):
    BACKGROUND = '#CCCCCC'
    SELECTED = '#7A7A7A'

    def __init__(self, *args, **kwargs):
        self.text = kwargs.pop('text', '')
        super().__init__(*args, **kwargs)

        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)

        self['background'] = self.BACKGROUND
        self.label = tk.Entry(self, background='white', borderwidth=0)
        self.label.insert(tk.END, self.text)
        self.label.pack(expand=True, fill='both', padx=2, pady=2)

    def on_enter(self, e):
        self['background'] = self.SELECTED

    def on_leave(self, e):
        self['background'] = self.BACKGROUND

    def insert(self, loc, text):
        self.label.insert(loc, text)

    def clear(self):
        self.label.delete(0, 'end')

    def get(self):
        return self.label.get()


class ModernButton(tk.Frame):
    BACKGROUND = '#CCCCCC'
    SELECTED = '#7A7A7A'

    def __init__(self, *args, **kwargs):
        self.command = kwargs.pop('command', None)
        self.text_main = kwargs.pop('text', '')
        self.height = kwargs.pop('height', None)
        self.width = kwargs.pop('width', None)
        super().__init__(*args, **kwargs)
        self['background'] = self.BACKGROUND

        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)
        self.bind("Button-1", self.mousedown)

        self.button = tk.Button(self,
                                text=self.text_main,
                                font=('Segoe', 10),
                                highlightthickness=0,
                                borderwidth=0,
                                background=self.BACKGROUND,
                                state='disabled',
                                disabledforeground='black',
                                height=self.height,
                                width=self.width)
        self.button.bind('<ButtonPress-1>', self.mousedown)
        self.button.bind('<ButtonRelease-1>', self.mouseup)
        self.button.pack(pady=2, padx=2, expand=True, fill='both')

    def on_enter(self, e):
        self['background'] = self.SELECTED

    def on_leave(self, e):
        self['background'] = self.BACKGROUND

    def mousedown(self, e):
        self.button.config(relief='sunken')
        self.button.config(relief='sunken')

    def mouseup(self, e):
        self.button.config(relief='raised')
        self.button.config(relief='raised')
        if self.command is not None:
            self.command()


class DoubleTextButton(tk.Frame):
    # BACKGROUND = '#E9ECED'
    BACKGROUND = 'white'
    SELECTED = '#D8EAF9'

    def __init__(self, *args, **kwargs):
        self.command = kwargs.pop('command', None)
        self.text_main = kwargs.pop('text_main', '')
        self.text_sub = kwargs.pop('text_sub', '')

        super().__init__(*args, **kwargs)
        self['background'] = self.master['background']

        self.bind("<Enter>", self.on_enter)
        self.bind("<Leave>", self.on_leave)
        self.bind("Button-1", self.mousedown)

        self.label_main = tk.Button(self,
                                    text=self.text_main,
                                    font=('Segoe', 10),
                                    highlightthickness=0,
                                    borderwidth=0,
                                    background=self.BACKGROUND,
                                    state='disabled',
                                    disabledforeground='black',
                                    anchor='w')
        self.label_main.bind('<ButtonPress-1>', self.mousedown)
        self.label_main.bind('<ButtonRelease-1>', self.mouseup)
        self.label_main.pack(fill='both')

        self.label_sub = tk.Button(self,
                                   text=self.text_sub,
                                   font=('Segoe', 10),
                                   highlightthickness=0,
                                   borderwidth=0,
                                   background=self.BACKGROUND,
                                   state='disabled',
                                   disabledforeground='#666666',
                                   anchor='w')
        self.label_sub.bind('<ButtonPress-1>', self.mousedown)
        self.label_sub.bind('<ButtonRelease-1>', self.mouseup)
        self.label_sub.pack(fill='both')

    def on_enter(self, e):
        self['background'] = self.SELECTED
        self.label_main['background'] = self.SELECTED
        self.label_sub['background'] = self.SELECTED

    def on_leave(self, e):
        self['background'] = self.BACKGROUND
        self.label_main['background'] = self.BACKGROUND
        self.label_sub['background'] = self.BACKGROUND

    def mousedown(self, e):
        self.label_main.config(relief='sunken')
        self.label_sub.config(relief='sunken')

    def mouseup(self, e):
        self.label_main.config(relief='raised')
        self.label_sub.config(relief='raised')
        if self.command is not None:
            self.command()

    def changetext_main(self, text):
        self.label_main.config(text=text)

    def changetext_sub(self, text):
        self.label_sub.config(text=text)


class CustomCheckButton(tk.Checkbutton):
    IMGSIZE = 30

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.image_checkbox_off = ImageTk.PhotoImage(
            Image.open('icons\\checkbox_empty.png').resize((self.IMGSIZE, self.IMGSIZE), Image.ANTIALIAS)
        )
        self.image_checkbox_on = ImageTk.PhotoImage(
            Image.open('icons\\checkbox_full.png').resize((self.IMGSIZE, self.IMGSIZE), Image.ANTIALIAS)
        )
        self.config(image=self.image_checkbox_off,
                    selectimage=self.image_checkbox_on,
                    selectcolor=self.master['background'],
                    background=self.master['background'],
                    activebackground=self.master['background'],
                    activeforeground=self.master['background'],
                    highlightcolor='red',
                    indicatoron=False,
                    bd=0)


if __name__ == '__main__':
    freeze_support()
    tool = Root()
    # tool.style = ttk.Style()
    # tool.style.theme_use('vista')
    tool.mainloop()
