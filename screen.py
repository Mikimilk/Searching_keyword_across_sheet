import tkinter as tk
from tkinter import *

class Draw(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        inp = StringVar()
        out = StringVar()
        self.geometry('400x250')
        self.title('Searching keyword across sheet [COMPUTERSYSTEM]')
        self.desc = tk.Label(None, text="Please fill in an information." ,font=('bold',16))
        self.desc.place(x=70,y=20)

        self.label_inp = Label(None, text="input file's name")
        self.label_inp.place(x=70, y=70)
        self.input_box = tk.Entry(textvariable = inp)
        self.input_file = raw_input(inp.get())
        self.input_box.place(x=190, y=70)

        self.label_out = Label(None, text="output file's name")
        self.label_out.place(x=70, y=100)
        self.out_box = tk.Entry(textvariable = out)
        self.out_file = raw_input(out.get())
        self.out_box.place(x=190, y=100)

        self.button = tk.Button(self, text="Go", command=self.on_button)
        self.button.place(x=190, y=140)

    def on_button(self):
        print(self.input_file)
        print(self.out_file)

app = Draw()
app.mainloop()