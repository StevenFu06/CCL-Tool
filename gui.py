import tkinter as tk
from tkinter import ttk
from tkinter import filedialog
from tkinter import *
from tkinter import messagebox

from pandastable import Table
import pandas as pd

from ccl import *
from compare import *
from package import Parent
from StyleFrame import StyleFrame, Styler

from multiprocessing import freeze_support
import threading
import os


class Root(tk.Tk):
    TITLE = 'CCL Tool'

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.ccl = CCL()
        self.ccl_saveas = None
        self.compare_saveas = None
        self.processes = {}

        self.page = -1
        self.title(self.TITLE)
        self.settings()

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

        self.update()
        self.minsize(self.winfo_width(), self.winfo_height())

    def settings(self):
        settings = tk.Button(self, text='i', width=1, height=1, borderwidth=0, command=lambda: Settings(self))
        settings.pack(anchor='ne', side='right', padx=2)

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

        if self.page < len(self.frames_active) - 1:
            self.page += 1
            self.frames_active[self.page].tkraise()
        else:
            self.page -= 1

        if self.page >= len(self.frames_active) - 1:
            self.button_next.configure(text='Start', command=lambda: Run(self))

    def prev_frame(self):
        self.button_next.configure(text='Next', command=self.next_frame)
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

        button_insert_illustrations = ttk.Button(center_check_button,
                                                 text='Insert / Delete Illustrations',
                                                 command=InsertDelIll)
        button_insert_illustrations.pack(anchor='center')

        button_check_filtered = ttk.Button(center_check_button,
                                           text='CCL Format Checker',
                                           command=FilterCheck)
        button_check_filtered.pack()

        button_rohs_compare = ttk.Button(center_check_button,
                                         text='RoHS BOM Comparison',
                                         command=lambda: ROHSCompare(self))
        button_rohs_compare.pack()

        center_check_button.tkraise()

    def nav_buttons(self):
        self.button_prev = ttk.Button(self, text='Previous', width=10, command=self.prev_frame)
        self.button_prev.pack(side='left', padx=(110, 0), pady=5)
        self.button_next = ttk.Button(self, text='Next', width=10, command=self.next_frame)
        self.button_next.pack(side='right', padx=(0, 110), pady=5)


class Settings(tk.Toplevel):
    def __init__(self, controller, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.controller = controller
        self.lift()
        self.focus_force()
        self.grab_set()

        self.setprocess()
        self.done()

    def setprocess(self):
        processframe = tk.Frame(self)
        processframe.pack()

        cpu_usage_label = ttk.Label(processframe, text='Enter desired CPU usage (0 for default): ')
        cpu_usage_label.pack(side='left')

        default = IntVar()
        default.set(0)
        self.cpu_usage = Entry(processframe, width=3, text=default)
        self.cpu_usage.pack(side='left')

        percent = ttk.Label(processframe, text='%')
        percent.pack(side='left')

    def done(self):
        button = ttk.Button(self, text='Done', command=self.done_cmd)
        button.pack()

    def done_cmd(self):
        cores = os.cpu_count()
        cpu_usage = int(self.cpu_usage.get())
        if cpu_usage < 0 or cpu_usage > 100:
            self.invalid_input()
        elif cpu_usage == 0:
            self.controller.ccl.processes = 1
        else:
            self.controller.ccl.processes = round(cpu_usage / 100 * cores)
            print(self.controller.ccl.processes)

    @staticmethod
    def invalid_input():
        messagebox.showerror('Error', 'Invalid Input')


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

        bom_save = ttk.Button(centered_frame, text='Save As Directory Bom Compare (Optional)',
                              command=self.file_dialog_save)
        bom_save.pack(anchor='center')
        self.label_filename_save = ttk.Label(centered_frame, text='')
        self.label_filename_save.pack(anchor='center')

    def file_dialog_old(self):
        filename = filedialog.askopenfile(initialdir='/', title='Select the Old AVL BOM')
        self.label_filename_old.configure(text=filename.name)
        self.label_filename_old.pack()

        self.ccl.avl_bom_path = filename.name
        self.ccl.avl_bom = pd.read_csv(filename.name)

    def file_dialog_new(self):
        filename = filedialog.askopenfile(initialdir='/', title='Select the New AVL BOM')
        self.label_filename_new.configure(text=filename.name)
        self.label_filename_new.pack()

        self.ccl.avl_bom_updated_path = filename.name
        self.ccl.avl_bom_updated = pd.read_csv(filename.name)

    def file_dialog_save(self):
        filename = filedialog.askdirectory(initialdir='/', title='Save As')
        self.label_filename_save.configure(text=filename)
        self.label_filename_save.pack()

        self.controller.compare_saveas = filename


class CCLInput(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.ccl = controller.ccl

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

        ccl_save = ttk.Button(centered_frame, text='Save as new CCL', command=self.file_dialog_save)
        self.label_filename_save = tk.Label(centered_frame, text='')

        ccl_save.pack(anchor='center', pady=5, padx=5)
        self.label_filename_save.pack(anchor='center')

    def file_dialog(self):
        filename = filedialog.askopenfile(initialdir='/', title='Select CCL')
        self.label_filename.configure(text=filename.name)
        self.label_filename.pack()

        self.ccl.ccl_docx = filename.name

    def file_dialog_save(self):
        filename = filedialog.asksaveasfilename(initialdir='/', title='Save As')
        self.label_filename_save.configure(text=filename)
        self.label_filename_save.pack()

        self.controller.ccl_saveas = filename


class CCLDocuments(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.ccl = controller.ccl

        info = tk.Button(self, text='?', width=1, height=1, borderwidth=0)
        info.pack(anchor='ne', side='right', padx=2)

        self.doc_input()
        self.path_checks()
        self.enovia_user()

    def doc_input(self):
        centered_frame = tk.Frame(self)
        centered_frame.pack(expand=True)

        docs = ttk.Button(centered_frame, text='Browse for CCL Documents Save Location', command=self.file_dialog)
        self.label_filename = tk.Label(centered_frame, text='')

        docs.pack(anchor='center', pady=5, padx=5)
        self.label_filename.pack(anchor='center')

    def file_dialog(self):
        filename = filedialog.askdirectory(initialdir='/', title='Select Directory')
        self.label_filename.configure(text=filename)
        self.label_filename.pack()

        self.ccl.path_ccl_data = filename

    def path_checks(self):
        centerframe = tk.Frame(self)
        centerframe.pack(expand=True)

        self.path_listbox = Listbox(centerframe, height=5)
        self.path_listbox.pack(side='left', fill='both')

        scroll = Scrollbar(centerframe, orient='vertical', command=self.path_listbox.yview)
        scroll.pack(side='left', fill='y')

        self.path_listbox.configure(yscrollcommand=scroll.set)

        addpath = ttk.Button(centerframe, text='Add Path', command=self.add_path)
        addpath.pack()

        delpath = ttk.Button(centerframe, text='Delete Path', command=self.del_path)
        delpath.pack()

        moveup = ttk.Button(centerframe, text='Move Up', command=self.move_up)
        moveup.pack()

        movedown = ttk.Button(centerframe, text='Move Down', command=self.move_down)
        movedown.pack()

    def add_path(self):
        filename = filedialog.askdirectory(initialdir='/', title='Select Directory')
        self.path_listbox.insert(END, filename)
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
            self.ccl.path_checks = [self.path_listbox.get(idx) for idx in range(self.path_listbox.size())]
        else:
            self.ccl.path_checks = []

    def enovia_user(self):
        centerframe_user = tk.Frame(self)
        centerframe_user.pack(expand=True)

        user_label = ttk.Label(centerframe_user, text='Enter Enovia Username')
        user_label.pack(side='left')
        self.user = ttk.Entry(centerframe_user)
        self.user.pack(side='left')

        centerframe_pass = tk.Frame(self)
        centerframe_pass.pack(expand=True, pady=2)

        password_label = ttk.Label(centerframe_pass, text='Enter Enovia Password')
        password_label.pack(side='left')
        self.password = ttk.Entry(centerframe_pass, show='*')
        self.password.pack(side='left')

        center_login = tk.Frame(self)
        center_login.pack(expand='True')
        login = ttk.Button(center_login, text='Login', command=self.login)
        login.pack(side='left')
        self.check = Label(center_login, text='')
        self.check.pack(side='left')

    def login(self):
        self.ccl.username = self.user.get()
        self.ccl.password = self.password.get()
        if len(self.user.get()) != 0 and len(self.password.get()) != 0:
            self.check.config(text=u'\u2713')


class Illustration(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.ccl = controller.ccl

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
        filename = filedialog.askdirectory(initialdir='/', title='Select Directory')
        self.label_filename.configure(text=filename)
        self.label_filename.pack()

        self.ccl.path_illustration = filename


class InsertDelIll(tk.Toplevel):
    HEIGHT = 400
    WIDTH = 200
    TITLE = 'Illustration Tool'

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.ccl = CCL()
        self.ill_path = None
        self.save_ccl = None

        self.minsize(self.HEIGHT, self.WIDTH)
        self.title(self.TITLE)

        self.illnum()
        self.illpath()
        self.illdir()
        self.browseccl()
        self.saveccl()
        self.run_buttons()

        self.update()
        self.minsize(self.winfo_width(), self.winfo_height())

    def illnum(self):
        centerframe = tk.Frame(self)
        centerframe.pack(expand=True, pady=5)

        label = ttk.Label(centerframe, text='Enter illustration number: Ill.')
        label.pack(side='left')

        num = ttk.Entry(centerframe, width=3)
        num.pack(side='left')
        self.ill_num = num.get()

    def illpath(self):
        def browse(self):
            filename = filedialog.askopenfile(initialdir='/', title='Select Illustration')
            label.configure(text=filename.name)
            self.ill_path = filename.name

        centerframe = tk.Frame(self)
        centerframe.pack(expand=True, pady=5)

        button = ttk.Button(centerframe, text='Browse Illustration', command=lambda: browse(self))
        button.pack()
        label = ttk.Label(centerframe, text='Select illustration to be inserted / deleted')
        label.pack()

    def illdir(self):
        def browse(self):
            filename = filedialog.askopenfile(initialdir='/', title='Select Directory')
            label.configure(text=filename.name)
            self.ccl.path_illustration = filename.name

        centerframe = tk.Frame(self)
        centerframe.pack(expand=True, pady=5)

        button = ttk.Button(centerframe, text='Browse Directory', command=lambda: browse(self))
        button.pack()
        label = ttk.Label(centerframe, text='Select Illustration Directory Location')
        label.pack()

    def browseccl(self):
        def browse(self):
            filename = filedialog.askdirectory(initialdir='/', title='Select CCL')
            label.configure(text=filename.name)
            self.ccl.ccl_docx = filename.name

        centerframe = tk.Frame(self)
        centerframe.pack(expand=True, pady=5)

        button = ttk.Button(centerframe, text='Browse CCL', command=lambda: browse(self))
        button.pack()
        label = ttk.Label(centerframe, text='Select CCL Location')
        label.pack()

    def saveccl(self):
        def browse(self):
            filename = filedialog.asksaveasfilename(initialdir='/', title='Save CCL')
            label.configure(text=filename)
            self.save_ccl = filename

        centerframe = tk.Frame(self)
        centerframe.pack(expand=True, pady=5)

        button = ttk.Button(centerframe, text='Save CCL', command=lambda: browse(self))
        button.pack()
        label = ttk.Label(centerframe, text='Save CCL Location')
        label.pack()

    def run_buttons(self):
        centerframe = tk.Frame(self)
        centerframe.pack(expand=True, pady=5)

        insert = ttk.Button(centerframe, text='Insert', command=self.insertcmd)
        insert.pack(side='left')

        delete = ttk.Button(centerframe, text='Delete', command=self.delcmd)
        delete.pack(side='right')

    def insertcmd(self):
        self.ccl.insert_illustration(self.ill_num, self.ill_path, self.save_ccl)

    def delcmd(self):
        self.ccl.delete_illustration(self.ill_num, self.ill_path, self.save_ccl)


class FilterCheck(tk.Toplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.ccl_loc = None
        self.check_button()

    def check_button(self):
        button = ttk.Button(self, text='Browse for CCL', command=self.browse)
        button.pack()
        self.label = ttk.Label(self, text='')
        self.label.pack()

    def browse(self):
        filename = filedialog.askopenfile(initialdir='/', title='Select CCL')
        self.label.configure(text=filename.name)
        self.ccl_loc = filename.name
        self.filtered = Parser(self.ccl_loc).filter()
        self.display_df()

    def display_df(self):
        frame = Frame(self)
        frame.pack(fill=BOTH, expand=True)
        self.dfdisplay = Table(frame, dataframe=self.filtered)
        self.dfdisplay.show()


class Run(tk.Toplevel):
    def __init__(self, controller, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.controller = controller
        self.ccl = controller.ccl

        self.prompt()
        self.controls()

        sys.stdout = TextRedirector(self.console, 'stdout')
        sys.stderr = TextRedirector(self.console, 'stderr')

        self.update()
        self.minsize(self.winfo_width(), self.winfo_height())

    def controls(self):
        framecrtl = Frame(self)
        framecrtl.pack()

        run = ttk.Button(framecrtl, text='Run', command=self.start_threading)
        run.pack(side='left')

        abort = ttk.Button(framecrtl, text='Abort', command=self.controller.destroy)
        abort.pack(side='right')

    def prompt(self):
        promptframe = Frame(self)
        promptframe.pack(expand=True, fill=BOTH, padx=5)

        self.progressbar = ttk.Progressbar(promptframe, mode='indeterminate')
        self.progressbar.pack(fill='x', pady=5)

        self.console = Text(promptframe, wrap='word')
        self.console.pack(side='left', expand=True, fill=BOTH)

        scroll = Scrollbar(promptframe, orient='vertical', command=self.console.yview)
        scroll.pack(side='right', fill='y')
        self.console.configure(yscrollcommand=scroll.set)

    def run(self):
        if self.controller.bom_compare.get():
            if self.controller.compare_saveas is None:
                raise ValueError('Missing BOM Comparison Save Location')

            print('Starting BOM Compare')
            self.ccl.save_compare(self.controller.compare_saveas)
            print('BOM Compare finished')

        if self.controller.ccl_update.get():
            if self.controller.ccl_saveas is None:
                raise ValueError('Missing CCL Input')
            else:
                print('Starting to update the CCL')
                self.ccl.update_ccl(self.controller.ccl_saveas)
                print('CCL Has been updated and saved')

        if self.controller.col_docs.get():
            print('Collecting Documents')
            self.ccl.collect_documents()
            print('Documents have been successfully collected')

        if self.controller.check_ill.get():
            if self.controller.ccl_saveas is None:
                raise ValueError('Missing CCL Input')
            print('Starting to Collect Illustrations')
            self.ccl.collect_illustrations()
            self.ccl.insert_illustration_data(self.controller.ccl_saveas)
            print('Illustrations have been collected and CCL has been updated')

        print('FINISHED!')

    def start_threading(self):
        self.submit_thread = threading.Thread(target=self.run)
        self.submit_thread.daemon = True
        self.submit_thread.start()
        self.progressbar.start()
        self.after(20, self.check_thread)

    def check_thread(self):
        if self.submit_thread.is_alive():
            self.after(20, self.check_thread)
        else:
            self.progressbar.stop()


class TextRedirector(object):
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, str):
        self.widget.configure(state="normal")
        self.widget.insert("end", str, (self.tag,))
        self.widget.configure(state="disabled")


class ROHSCompare(tk.Toplevel):
    def __init__(self, controller, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.controller = controller
        self.a_avl = None
        self.b_avl = None

        self.lift()
        self.focus_force()
        self.grab_set()

        self.input()
        self.output()

        self.update()
        self.minsize(self.winfo_width(), self.winfo_height())

    def input(self):
        frame_a = Frame(self)
        frame_a.pack(expand=True, fill='both')
        button_a = ttk.Button(frame_a, text='Browse for BOM A', width=17, command=self.fileinput_a)
        button_a.pack(side='left')
        self.label_a = tk.Label(frame_a, text='', bg='white', width=20)
        self.label_a.pack(side='right', expand=True, fill='x')

        frame_b = Frame(self)
        frame_b.pack(expand=True, fill='both')
        button_b = ttk.Button(frame_b, text='Browse for BOM B', width=17, command=self.fileinput_b)
        button_b.pack(side='left')
        self.label_b = tk.Label(frame_b, text='', bg='white')
        self.label_b.pack(side='right', expand=True, fill='x')

    def fileinput_a(self):
        file = filedialog.askopenfile(initialdir='/', title='Browse for BOM A')
        self.label_a.config(text=file.name)
        self.a_avl = pd.read_csv(file)

    def fileinput_b(self):
        file = filedialog.askopenfile(initialdir='/', title='Browse for BOM B')
        self.label_b.config(text=file.name)
        self.b_avl = pd.read_csv(file)

    def output(self):
        frame_save = Frame(self)
        frame_save.pack()

        save_a = ttk.Button(frame_save,
                            text='Generate Exclusive to BOM A',
                            command=lambda: self.start_threading(self.a_avl, self.b_avl))
        save_a.pack(side='left')

        save_b = ttk.Button(frame_save,
                            text='Generate Exclusive to BOM B',
                            command=lambda: self.start_threading(self.b_avl, self.a_avl))
        save_b.pack(side='right')

    def compare(self, a_avl, b_avl):
        a_bom = Bom(a_avl, Parent(a_avl).build_tree())
        b_bom = Bom(b_avl, Parent(b_avl).build_tree())
        a_only = [int(idx) for idx, pn in (a_bom - b_bom)]

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

        for idx in a_only:
            sf.apply_style_by_indexes(sf.index[idx], styler_obj=style)
        self.save_as(sf)

    @staticmethod
    def save_as(styleframe):
        file = filedialog.asksaveasfilename(initialdir='/',
                                            title='Save As',
                                            filetypes=(('Excel', '.xlsx'),),
                                            defaultextension='.xlsx')
        styleframe.to_excel(file).save()

    def show_progressbar(self):
        popup = tk.Toplevel(self)
        self.progressbar = ttk.Progressbar(popup, mode='indeterminate')
        self.progressbar.pack(expand=True, fill='x')

    def start_threading(self, a_avl, b_avl):
        self.show_progressbar()
        self.submit_thread = threading.Thread(target=lambda: self.compare(a_avl, b_avl))
        self.submit_thread.daemon = True
        self.submit_thread.start()
        self.progressbar.start()
        self.after(20, self.check_thread)

    def check_thread(self):
        if self.submit_thread.is_alive():
            self.after(20, self.check_thread)
        else:
            self.progressbar.stop()


if __name__ == '__main__':
    freeze_support()
    tool = Root()
    # tool.style = ttk.Style()
    # tool.style.theme_use('classic')
    tool.mainloop()
