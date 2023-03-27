import ctypes
import os.path
import smtplib
import ssl
import tkinter as tk
from tkinter import filedialog
import docx
import pandas

ctypes.windll.shcore.SetProcessDpiAwareness(1)

win = tk.Tk()

current_word_file = None
current_excel_file = None


def get_files():
    global current_excel_file, current_word_file
    doc_file = filedialog.askopenfilename(
        initialdir='/',
        title='Select the word document',
        filetypes=(('Word Document', "*.docx*"), ('All files', "*.*"))
    )
    excel_file = filedialog.askopenfilename(
        initialdir='/',
        title='Select the excel sheet',
        filetypes=(('Excel Sheet', "*.xlsx*"), ('All files', "*.*"))
    )

    doc_extn = os.path.splitext(doc_file)[1]
    xl_extn = os.path.splitext(excel_file)[1]

    if (doc_extn != '.docx') or (xl_extn != '.xlsx'):
        return

    current_excel_file = excel_file
    current_word_file = doc_file

    send_btn.config(state='normal')


def send_all_mails():
    doc = docx.Document(current_word_file)
    full_text = []
    text = ''
    for para in doc.paragraphs:
        full_text.append(para.text)
        text = '\n'.join(full_text)

    for i, row in pandas.read_excel(current_excel_file):
        name = row['Names']
        send_text = text.replace('[Name]', name)
        send_mail(send_text, row["Emails"])

    send_btn.config(state='disabled')


def send_mail(body, to):
    gmail = "bhuvanshgoyal2004@gmail.com"
    password = "jweypmgzzneabrmt"
    context = ssl.create_default_context()
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.ehlo()
        server.starttls(context=context)
        server.ehlo()
        server.login(gmail, password)
        server.sendmail(gmail, to, body)
    print("mail-sent")


button = tk.Button(text="Go", command=get_files)
button.pack()


send_btn = tk.Button(text="send", state='disabled', command=send_mail)
send_btn.pack()

win.mainloop()
