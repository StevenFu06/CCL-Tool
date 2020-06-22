from tkinter import ttk
import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk
from enovia import Enovia
import threading
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
        for F in (BomCompare, UpdateCCL, DocumentCollector):
            frame = F(self, self.main_notebook)
            self.frames[F] = frame

        self.main_options()
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
        self.shared = tk.BooleanVar()  # For shared variables in settings
        self.headless = tk.BooleanVar(value=True)  # For troubleshooting, toggles headless

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

    def main_options(self):
        self.mainframe = tk.Frame(self.main_notebook, bg='white')
        self.mainframe.pack(anchor='center')

        frame_compare = tk.Frame(self.mainframe, bg='white')
        frame_compare.pack(fill='x', pady=10)
        compare_check = CustomCheckButton(frame_compare, variable=self.compare_select)
        compare_check.pack(side='left', padx=(0, 2))
        compare_button = DoubleTextButton(frame_compare, text_main='Bill of Materials Comparison',
                                          text_sub='Will generate a bill of materials comparison report',
                                          command=lambda: self.raiseframe(BomCompare))
        compare_button.pack(side='left', fill='x', expand=True)

        frame_update = tk.Frame(self.mainframe, bg='white')
        frame_update.pack(fill='x', pady=10)
        update_check = CustomCheckButton(frame_update, variable=self.update_select)
        update_check.pack(side='left', padx=(0, 2))
        update_button = DoubleTextButton(frame_update, text_main='Update CCL',
                                         text_sub='Will output an updated CCL',
                                         command=lambda: self.raiseframe(UpdateCCL))
        update_button.pack(side='left', fill='x', expand=True)

        frame_docs = tk.Frame(self.mainframe, bg='white')
        frame_docs.pack(fill='x', pady=10)
        docs_check = CustomCheckButton(frame_docs, variable=self.docs_select)
        docs_check.pack(side='left', padx=(0, 2))
        docs_button = DoubleTextButton(frame_docs, text_main='Collect CCL Documents',
                                       text_sub='Will collect all documents associated with CCL',
                                       command=lambda: self.raiseframe(DocumentCollector))
        docs_button.pack(side='left', fill='x', expand=True)

        frame_ills = tk.Frame(self.mainframe, bg='white')
        frame_ills.pack(fill='x', pady=10)
        ills_check = CustomCheckButton(frame_ills, variable=self.ills_select)
        ills_check.pack(side='left', padx=(0, 2))
        ills_button = DoubleTextButton(frame_ills, text_main='Collect Illustrations',
                                       text_sub='Will collect all illustrations associated with CCL')
        ills_button.pack(side='left', fill='x', expand=True)

        run_button = ModernButton(self.mainframe, text='Press to Run', height=1)
        run_button.pack(expand=True, fill='x', pady=5)

    def settings_frame(self):
        self.settingsframe = tk.Frame(self.settings_notebook, bg='white')
        self.settingsframe.pack(anchor='center')
        sharedprocess = ttk.Checkbutton(self.settingsframe,
                                        text='Share input between process',
                                        variable=self.shared)
        sharedprocess.pack()
        headless = ttk.Checkbutton(self.settingsframe,
                                   text='Enable headless Chrome',
                                   variable=self.headless)
        headless.pack()

    def raiseframe(self, name):
        self.mainframe.forget()
        frame = self.frames[name]
        frame.pack(expand=True, fill='both', padx=10)
        frame.update()
        frame.event_generate('<<ShowFrame>>')

    def back(self, ontop):
        self.frames[ontop].forget()
        self.mainframe.pack()


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
                                                defaultextension='.zip')
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
            self.old_entry.insert(tk.END, self.root.bom_compare_old.name)
            self.root.ccl_update_old = self.root.bom_compare_old

        if self.root.shared.get() and self.root.ccl_update_new is None:
            self.new_entry.insert(tk.END, self.root.bom_compare_new.name)
            self.root.ccl_update_new = self.root.bom_compare_new


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
        self.user.label.config(foreground='#666666')
        self.user.label.bind("<Button-1>", self.clear_user)
        self.user.pack(expand=True, fill='both', side='left', padx=(0, 2))
        self.password = ModernEntry(self.enoviaframe, text='Enter Enovia Password')
        self.password.label.config(foreground='#666666')
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
        centerframe.pack(expand=True, fill='both', pady=5)

        self.path_listbox = tk.Listbox(centerframe, height=10)
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
            self.root.docs_path = [self.path_listbox.get(idx) for idx in range(self.path_listbox.size())]
        else:
            self.root.docs_path = []

    def sharedvar(self, e):
        if self.root.shared.get() and self.root.docs_ccl is None:
            self.ccl_entry.insert(tk.END, self.root.ccl_update_loc.name)
            self.root.docs_ccl = self.root.ccl_update_loc

    def enoviacheck(self):
        with Enovia(self.user.get(), self.password.get(), headless=self.root.headless) as enovia:
            pass

    def start_thread(self):
        self.thread = threading.Thread(target=self.enoviacheck)
        self.thread.daemon = True
        self.thread.start()
        self.after(20, self.check_thread)

    def check_thread(self):
        if self.thread.is_alive():
            self.after(20, self.check_thread)


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
    tool = Root()
    tool.mainloop()
