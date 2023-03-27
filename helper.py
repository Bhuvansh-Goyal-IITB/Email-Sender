import smtplib
import ssl
import docx
import pandas


def send_mails_to_all(doc_file, excel_file, sender, password):
    doc = docx.Document(doc_file)
    full_text = []
    text = ''
    for para in doc.paragraphs:
        full_text.append(para.text)
        text = '\n\n'.join(full_text)

    df = pandas.read_excel(excel_file)
    if 'Email' not in df:
        return
    unsent_msg = []
    for i, row in df.iterrows():
        send_text = text
        for column in df:
            send_text = send_text.replace(f"[{column}]", row[column])
        try:
            send_mail(sender, password, row["Email"], send_text)
        except Exception as e:
            unsent_msg.append(f"Couldn't send email to {row['Email']} due to {e}")


def send_mail(sender, password, receiver, msg):
    context = ssl.create_default_context()
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.ehlo()
        server.starttls(context=context)
        server.ehlo()
        server.login(sender, password)
        server.sendmail(sender, receiver, msg)
