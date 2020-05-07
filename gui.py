import tkinter as tk
from tkinter import ttk
from tkinter import filedialog


def info_button():
    popup = tk.Tk()
    popup.wm_title("Help")
    label = ttk.Label(popup, text='Example info, directions / explinations go here')
    label.pack(side="top", fill="x", pady=10)
    b1 = tk.Button(popup, text="Okay", command=popup.destroy)
    b1.pack()
    popup.mainloop()


class Tool(tk.Tk):
    def __init__(self, *args, **kwargs):
        super(Tool, self).__init__(*args, **kwargs)
        self.width = 200
        self.height = 400
        self.minsize(self.height, self.width)
        self.title('CCL Tool')

        container = tk.Frame(self)
        container.pack(pady=(1, 0), expand=True, fill='both')
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        self.frames = {}
        for F in (OptionsPage, BomInput, CCLInput, CCLDocuments, Illustration):
            frame = F(container, self)
            self.frames[F] = frame
            frame.grid(row=0, column=0, stick='nsew')
        self.show_frame(OptionsPage)

    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()


class OptionsPage (tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg='white')
        self.controller = controller

        info = tk.Button(self, text='?', width=1, height=1, bg='white', borderwidth=0, command=info_button)
        info.pack(anchor='ne', side='right', padx=2)

        self.main_options()
        self.nav_buttons()

    def main_options(self):
        center_check_button = tk.Frame(self, bg='white', width=5, height=5)
        center_check_button.pack(padx=5, pady=5, expand=True)

        check_bom_compare = tk.Checkbutton(center_check_button, text='Bill of Material Comparison', bg='white')
        check_bom_compare.pack(anchor='w')

        check_ccl_update = tk.Checkbutton(center_check_button, text='Update CCL', bg='white')
        check_ccl_update.pack(anchor='w')

        check_collect_documents = tk.Checkbutton(center_check_button, text='Collect CCL Documents', bg='white')
        check_collect_documents.pack(anchor='w')

        check_illustrations = tk.Checkbutton(center_check_button, text='Collect Illustrations', bg='white')
        check_illustrations.pack(anchor='w')

        button_insert_illustrations = tk.Button(center_check_button, text='Insert / Delete Illustrations')
        button_insert_illustrations.pack(anchor='center')

    def nav_buttons(self):
        button_prev = tk.Button(self, text='Previous',  width=10, height=1, state='disabled')
        button_prev.pack(side='left', padx=(110, 0), pady=5)
        button_next = tk.Button(self, text='Next', width=10, height=1, command=self.next_cmd)
        button_next.pack(side='right', padx=(0, 110), pady=5)

    def next_cmd(self):
        self.controller.show_frame(BomInput)


class BomInput(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg='white')
        self.controller = controller

        info = tk.Button(self, text='?', width=1, height=1, bg='white', borderwidth=0, command=info_button)
        info.pack(anchor='ne', side='right', padx=2)

        self.bom_input()
        self.nav_buttons()

    def bom_input(self):
        centered_frame = tk.Frame(self, bg='white', width=5, height=5)
        centered_frame.pack(padx=5, pady=5, expand=True)

        bom_old = tk.Button(centered_frame, text='Browse Old AVL BOM', command=self.file_dialog_old)
        self.label_filename_old = tk.Label(centered_frame, text='')

        bom_old.pack(anchor='center', pady=5, padx=5)
        self.label_filename_old.pack(anchor='center')

        bom_new = tk.Button(centered_frame, text='Browse New AVL BOM', command=self.file_dialog_new)
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

    def nav_buttons(self):
        button_prev = tk.Button(self, text='Previous',  width=10, height=1, command=self.prev_cmd)
        button_prev.pack(side='left', anchor='s', padx=(110, 0), pady=5)
        button_next = tk.Button(self, text='Next', width=10, height=1, command=self.next_cmd)
        button_next.pack(side='right', anchor='s', padx=(0, 110), pady=5)

    def next_cmd(self):
        self.controller.show_frame(CCLInput)

    def prev_cmd(self):
        self.controller.show_frame(OptionsPage)


class CCLInput(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg='white')
        self.controller = controller

        info = tk.Button(self, text='?', width=1, height=1, bg='white', borderwidth=0, command=info_button)
        info.pack(anchor='ne', side='right', padx=2)

        self.ccl_input()
        self.nav_buttons()

    def ccl_input(self):
        centered_frame = tk.Frame(self, bg='white', width=5, height=5)
        centered_frame.pack(padx=5, pady=5, expand=True)

        ccl = tk.Button(centered_frame, text='Browse for CCL', command=self.file_dialog)
        self.label_filename = tk.Label(centered_frame, text='')

        ccl.pack(anchor='center', pady=5, padx=5)
        self.label_filename.pack(anchor='center')

    def file_dialog(self):
        filename = filedialog.askopenfile(initialdir='/', title='Select CCL')
        self.label_filename.configure(text=filename.name)
        self.label_filename.pack()

    def nav_buttons(self):
        button_prev = tk.Button(self, text='Previous',  width=10, height=1, command=self.prev_cmd)
        button_prev.pack(side='left', anchor='s', padx=(110, 0), pady=5)
        button_next = tk.Button(self, text='Next', width=10, height=1, command=self.next_cmd)
        button_next.pack(side='right', anchor='s', padx=(0, 110), pady=5)

    def next_cmd(self):
        self.controller.show_frame(CCLDocuments)

    def prev_cmd(self):
        self.controller.show_frame(BomInput)


class CCLDocuments(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg='white')
        self.controller = controller

        info = tk.Button(self, text='?', width=1, height=1, bg='white', borderwidth=0, command=info_button)
        info.pack(anchor='ne', side='right', padx=2)

        self.doc_input()
        self.nav_buttons()

    def doc_input(self):
        centered_frame = tk.Frame(self, bg='white', width=5, height=5)
        centered_frame.pack(padx=5, pady=5, expand=True)

        docs = tk.Button(centered_frame, text='Browse for CCL Documents Save Location', command=self.file_dialog)
        self.label_filename = tk.Label(centered_frame, text='')

        docs.pack(anchor='center', pady=5, padx=5)
        self.label_filename.pack(anchor='center')

    def file_dialog(self):
        filename = filedialog.askopenfile(initialdir='/', title='Select Directory')
        self.label_filename.configure(text=filename.name)
        self.label_filename.pack()

    def nav_buttons(self):
        button_prev = tk.Button(self, text='Previous',  width=10, height=1, command=self.prev_cmd)
        button_prev.pack(side='left', anchor='s', padx=(110, 0), pady=5)
        button_next = tk.Button(self, text='Next', width=10, height=1, command=self.next_cmd)
        button_next.pack(side='right', anchor='s', padx=(0, 110), pady=5)

    def next_cmd(self):
        self.controller.show_frame(Illustration)

    def prev_cmd(self):
        self.controller.show_frame(CCLInput)

class Illustration(tk.Frame):
    def __init__(self, parent, controller):
        super().__init__(parent, bg='white')
        self.controller = controller

        info = tk.Button(self, text='?', width=1, height=1, bg='white', borderwidth=0, command=info_button)
        info.pack(anchor='ne', side='right', padx=2)

        self.ill_input()
        self.nav_buttons()

    def ill_input(self):
        centered_frame = tk.Frame(self, bg='white', width=5, height=5)
        centered_frame.pack(padx=5, pady=5, expand=True)

        ills = tk.Button(centered_frame, text='Browse for Illustration Save Location', command=self.file_dialog)
        self.label_filename = tk.Label(centered_frame, text='')

        ills.pack(anchor='center', pady=5, padx=5)
        self.label_filename.pack(anchor='center')

    def file_dialog(self):
        filename = filedialog.askopenfile(initialdir='/', title='Select Directory')
        self.label_filename.configure(text=filename.name)
        self.label_filename.pack()

    def nav_buttons(self):
        button_prev = tk.Button(self, text='Previous', width=10, height=1, command=self.prev_cmd)
        button_prev.pack(side='left', anchor='s', padx=(110, 0), pady=5)
        button_start = tk.Button(self, text='Start Process', width=10, height=1)
        button_start.pack(side='right', anchor='s', padx=(0, 110), pady=5)

    def prev_cmd(self):
        self.controller.show_frame(CCLDocuments)


if __name__ == '__main__':
    tool = Tool()
    # tool.style = ttk.Style()
    # tool.style.theme_use('classic')
    tool.mainloop()
