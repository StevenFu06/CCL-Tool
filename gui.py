import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

from ccl import *


class Root(tk.Tk):
    HEIGHT = 400
    WIDTH = 200
    TITLE = 'CCL Tool'

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.ccl = CCL()

        self.page = -1
        self.minsize(self.HEIGHT, self.WIDTH)
        self.title(self.TITLE)

        self.container = tk.Frame(self)
        self.container.pack(pady=(1, 0), expand=True, fill='both')
        self.container.grid_rowconfigure(0, weight=1)
        self.container.grid_columnconfigure(0, weight=1)

        self.frames_standby = {}
        self.frames_active = []
        for F in (BomInput, CCLInput, CCLDocuments, Illustration):
            frame = F(self.container, self)
            self.frames_standby[F] = frame
            frame.grid(row=0, column=0, stick='nsew')

        self.main_options()
        self.nav_buttons()

    def next_frame(self):
        if not self.frames_active:
            if self.bom_compare.get() or self.ccl_update.get():
                self.frames_active.append(self.frames_standby[BomInput])

            if self.ccl_update.get() or self.col_docs.get() or self.check_ill.get():
                self.frames_active.append(self.frames_standby[CCLInput])

            if self.col_docs.get() or self.check_ill.get():
                self.frames_active.append(self.frames_standby[CCLDocuments])

            if self.check_ill.get():
                self.frames_active.append(self.frames_standby[Illustration])

        if self.page < len(self.frames_active)-1:
            self.page += 1
            self.frames_active[self.page].tkraise()
        else:
            self.page -= 1

        if self.page >= len(self.frames_active)-1:
            self.button_next.configure(text='Start')

    def prev_frame(self):
        self.button_next.configure(text='Next')
        if self.page > 0:
            self.page -= 1
            self.frames_active[self.page].tkraise()
        else:
            self.page = -1
            self.frames_active = []
            self.options_frame.tkraise()

    def main_options(self):
        self.options_frame = tk.Frame(self.container, width=5, height=5)
        self.options_frame.grid(row=0, column=0, stick='nsew')

        center_check_button = tk.Frame(self.options_frame, width=5, height=5)
        center_check_button.pack(padx=5, pady=5, expand=True)

        self.bom_compare = tk.BooleanVar()
        self.ccl_update = tk.BooleanVar()
        self.col_docs = tk.BooleanVar()
        self.check_ill = tk.BooleanVar()

        check_bom_compare = ttk.Checkbutton(center_check_button,
                                            text='Bill of Material Comparison',
                                            variable=self.bom_compare)
        check_bom_compare.pack(anchor='w')

        check_ccl_update = ttk.Checkbutton(center_check_button,
                                           text='Update CCL',
                                           variable=self.ccl_update)
        check_ccl_update.pack(anchor='w')

        check_collect_documents = ttk.Checkbutton(center_check_button,
                                                  text='Collect CCL Documents',
                                                  variable=self.col_docs)
        check_collect_documents.pack(anchor='w')

        check_illustrations = ttk.Checkbutton(center_check_button,
                                              text='Collect Illustrations',
                                              variable=self.check_ill)
        check_illustrations.pack(anchor='w')

        button_insert_illustrations = ttk.Button(center_check_button, text='Insert / Delete Illustrations')
        button_insert_illustrations.pack(anchor='center')
        center_check_button.tkraise()

    def nav_buttons(self):
        self.button_prev = ttk.Button(self, text='Previous', width=10, command=self.prev_frame)
        self.button_prev.pack(side='left', padx=(110, 0), pady=5)
        self.button_next = ttk.Button(self, text='Next', width=10, command=self.next_frame)
        self.button_next.pack(side='right', padx=(0, 110), pady=5)

class BomInput(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.ccl = controller.ccl
        self.controller = controller

        info = tk.Button(self, text='?', width=1, height=1, borderwidth=0)
        info.pack(anchor='ne', side='right', padx=2)

        self.bom_input()

    def bom_input(self):
        centered_frame = tk.Frame(self, width=5, height=5)
        centered_frame.pack(padx=5, pady=5, expand=True)

        bom_old = ttk.Button(centered_frame, text='Browse Old AVL BOM', command=self.file_dialog_old)
        self.label_filename_old = tk.Label(centered_frame, text='')

        bom_old.pack(anchor='center', pady=5, padx=5)
        self.label_filename_old.pack(anchor='center')

        bom_new = ttk.Button(centered_frame, text='Browse New AVL BOM', command=self.file_dialog_new)
        self.label_filename_new = tk.Label(centered_frame, text='')

        bom_new.pack(anchor='center', pady=5, padx=5)
        self.label_filename_new.pack(anchor='center')

    def file_dialog_old(self):
        filename = filedialog.askopenfile(initialdir='/', title='Select the Old AVL BOM')
        self.label_filename_old.configure(text=filename.name)
        self.label_filename_old.pack()

    def file_dialog_new(self):
        filename = filedialog.askopenfile(initialdir='/', title='Select the New AVL BOM')
        self.label_filename_new.configure(text=filename.name)
        self.label_filename_new.pack()


class CCLInput(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        info = tk.Button(self, text='?', width=1, height=1, borderwidth=0)
        info.pack(anchor='ne', side='right', padx=2)

        self.ccl_input()

    def ccl_input(self):
        centered_frame = ttk.Frame(self, width=5, height=5)
        centered_frame.pack(padx=5, pady=5, expand=True)

        ccl = ttk.Button(centered_frame, text='Browse for CCL', command=self.file_dialog)
        self.label_filename = tk.Label(centered_frame, text='')

        ccl.pack(anchor='center', pady=5, padx=5)
        self.label_filename.pack(anchor='center')

    def file_dialog(self):
        filename = filedialog.askopenfile(initialdir='/', title='Select CCL')
        self.label_filename.configure(text=filename.name)
        self.label_filename.pack()


class CCLDocuments(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        info = tk.Button(self, text='?', width=1, height=1, borderwidth=0)
        info.pack(anchor='ne', side='right', padx=2)

        self.doc_input()

    def doc_input(self):
        centered_frame = tk.Frame(self, width=5, height=5)
        centered_frame.pack(padx=5, pady=5, expand=True)

        docs = ttk.Button(centered_frame, text='Browse for CCL Documents Save Location', command=self.file_dialog)
        self.label_filename = tk.Label(centered_frame, text='')

        docs.pack(anchor='center', pady=5, padx=5)
        self.label_filename.pack(anchor='center')

    def file_dialog(self):
        filename = filedialog.askopenfile(initialdir='/', title='Select Directory')
        self.label_filename.configure(text=filename.name)
        self.label_filename.pack()


class Illustration(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller

        info = tk.Button(self, text='?', width=1, height=1, borderwidth=0)
        info.pack(anchor='ne', side='right', padx=2)

        self.ill_input()

    def ill_input(self):
        centered_frame = tk.Frame(self, width=5, height=5)
        centered_frame.pack(padx=5, pady=5, expand=True)

        ills = ttk.Button(centered_frame, text='Browse for Illustration Save Location', command=self.file_dialog)
        self.label_filename = tk.Label(centered_frame, text='')

        ills.pack(anchor='center', pady=5, padx=5)
        self.label_filename.pack(anchor='center')

    def file_dialog(self):
        filename = filedialog.askopenfile(initialdir='/', title='Select Directory')
        self.label_filename.configure(text=filename.name)
        self.label_filename.pack()


if __name__ == '__main__':
    tool = Root()
    # tool.style = ttk.Style()
    # tool.style.theme_use('classic')
    tool.mainloop()
