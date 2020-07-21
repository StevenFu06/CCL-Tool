import tkinter as tk
from tkinter import ttk


class Root(tk.Tk):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        frame = tk.Frame(self, bg='red', width=50, height=50)
        frame.pack(fill='both', expand=True, pady=5, padx=5)

        button = ModernButton(frame, text='Help', command=self.test_command)
        button.pack()

    def test_command(self):
        print('test')


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


if __name__ == '__main__':
    root = Root()
    root.mainloop()
