import time

import customtkinter as ctk
from tkinter import filedialog, messagebox
import os
from helper import send_mails_to_all


FONT = ("Calibri", 25, 'normal')


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode('dark')
        self.iconbitmap('email.ico')
        self.title = "Email Sender"
        self.current_doc = None
        self.current_excel = None

        self.container = ctk.CTkFrame(master=self)
        self.container.grid(row=0, column=0, padx=40, pady=40)

        self.add_files_button = ctk.CTkButton(master=self.container, width=150, font=FONT, text='Add files', command=self.add_files)
        self.add_files_button.grid(row=2, column=0, padx=30, pady=20)

        self.files_label = ctk.CTkLabel(master=self.container, font=(FONT[0], 16, FONT[2]))

        self.send_button = ctk.CTkButton(master=self.container, width=150, font=FONT, text='Send', state='disabled', command=self.send_mails)
        self.send_button.grid(row=2, column=1, padx=30, pady=20)

    def send_mails(self):
        self.files_label.grid_remove()
        self.send_button.configure(state='disabled')
        self.add_files_button.configure(state='disabled')
        self.container.grid_remove()

        time.sleep(2)

        with open('data.txt') as file:
            lines = file.readlines()
            email = lines[0]
            password = lines[1]

        send_mails_to_all(
            doc_file=self.current_doc,
            excel_file=self.current_excel,
            sender=email,
            password=password,
        )

        self.add_files_button.configure(state='normal')
        self.container.grid(row=0, column=0, padx=40, pady=40)
        # show popup for done

    def add_files(self):
        self.current_doc = filedialog.askopenfilename(
            initialdir='/Desktop',
            title='Select the word document',
            filetypes=(('Word Document', "*.docx*"), ('All files', "*.*"))
        )

        self.current_excel = filedialog.askopenfilename(
            initialdir='/Desktop',
            title='Select the word document',
            filetypes=(('Excel Sheet', "*.xlsx*"), ('All files', "*.*"))
        )

        doc_extension = os.path.splitext(self.current_doc)[1]
        excel_extension = os.path.splitext(self.current_excel)[1]

        if (doc_extension != '.docx') or (excel_extension != '.xlsx'):
            return

        self.files_label.configure(text=f'{self.current_doc}\n{self.current_excel}')
        self.files_label.grid(row=3, column=0, columnspan=2, padx=20, pady=5)
        self.send_button.configure(state='normal')

    def set_const_size(self, width, height):
        self.minsize(width, height)
        self.maxsize(width, height)

