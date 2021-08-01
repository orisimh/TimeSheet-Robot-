from tkinter import filedialog
from tkinter import *
import TimeSheetProject
from PIL import ImageTk, Image
from tkinter import messagebox
import sys
# import Classifier
# import Validation

import os
import os.path
from os import path
# import pandas as pd
# import numpy as np
import operator
from importlib import reload


class GUI:

    def __init__(self, root):

        # self.validate = Validation.Validation()

        logo = self.resource_path("cp.ico")
        root.wm_iconbitmap(logo)  # 'cp.ico'
        root.title("TimeSheet Robot")
        root.configure(background="white")
        root.geometry("750x350+200+200")
        self.folder_path = StringVar()

        image = self.resource_path("cp.png")

        image = Image.open(image)  #"cp.png"
        image = image.resize((750, 350))

        bg = ImageTk.PhotoImage(image)
        pic = Label(root, image=bg)
        pic.place(x=0, y=0  , relwidth=1 , relheig=1  )


        # # Create Frame
        # frame1 = Frame(root)
        # frame1.pack(pady=20)


        # horizonaly space
        # Label(root, text="", bg="white").grid(row=0, column=0)
        # Label(root, text="", bg="white").grid(row=1, column=0)
        # Label(root, text="", bg="white").grid(row=2, column=0)

        root.rowconfigure(0, minsize=20)
        root.rowconfigure(1, minsize=20)

        root.columnconfigure(0, minsize=20)
        root.columnconfigure(1, minsize=20)

        Directory_Path = Label(root, text="Excel Path:", bg="white", width=15)
        Directory_Path.grid(row=3, column=0, ipady=10, sticky=W)

        self.ThePath = Label(master=root, textvariable=self.folder_path, bd=1, relief="solid", bg="white", width=20)
        self.ThePath.grid(row=3, column=1, ipadx=200, sticky=SW, pady=10, ipady=2)
        #
        Label(root, text="  ", bg="white").grid(row=3, column=2)  # vertically space
        #
        # Open Button
        self.Open = Button(text="Open", command=self.Open, bd=1, relief="solid", width=10)
        self.Open.grid(row=3, column=3)



        self.var1 = IntVar()
        Checkbutton(root, text="ניקוי האקסל לאחר סיום", variable=self.var1 ,bg="white" ).grid(row=5,column=1 , sticky=W)

        # Label(root, text="", bg="white").grid(row=6, column=0)  # horizonaly space

        root.rowconfigure(6, minsize=120)
        root.rowconfigure(7, minsize=120)

        root.columnconfigure(6, minsize=120)
        root.columnconfigure(7, minsize=120)

        # Upload Button
        self.Upload = Button(text="Upload", command=self.Upload, bd=1, relief="solid", width=15, )
        self.Upload.configure(state="disabled")
        self.Upload.grid(row=7, column=1, padx=200)

        root.mainloop()




    def Open(self):


        filename = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])

        self.folder_path.set(filename)
        self.Upload.configure(state="normal")



    def Upload(self):


        self.ts = TimeSheetProject.TimeSheet(self.folder_path.get())

        rs = self.ts.execute_driver()

        if(rs == 1):
            self.folder_path.set("")
            self.Upload.configure(state="disabled")

        # status = self.ts.clean_empty_rows()
        else:
            self.ts.play()
            self.ts.clean_excel(self.var1.get())
            # else:
            self.folder_path.set("")
            self.Upload.configure(state="disabled")




    def resource_path(self, relative_path):
        """ Get absolute path to resource, works for dev and for PyInstaller """
        base_path = getattr(sys, '_MEIPASS', os.getcwd()) #os.path.dirname(os.path.abspath(__file__)))
        print(base_path)
        return os.path.join(base_path, relative_path)

if __name__ == '__main__':


    root = Tk()
    my_gui = GUI(root)
    root.mainloop()
