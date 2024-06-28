from customtkinter import *
import os
from datetime import date
import json
import smtplib, ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from win32api import *


odt_name = date.today().strftime("%d.%m.%Y") + ".odt"

pack_path = "C:\\Users\\Czerwiec\\Desktop\\VS Code work\\csv\\2024-06-28 x64"

file_version = ""

is_file = True

def send_email():
    sender_email = "tomasz.czerwinski@cadprojekt.com.pl"
    emails = ["czerrwiec@gmail.com"]
    password = "tomczer23"

    message = MIMEMultipart("multipart")
    # message = MIMEMultipart("alternative")
    message["Subject"] = "Myszy latają kluczem"
    message["From"] = sender_email
    message["To"] = ', '.join(emails) 

    # Create the plain-text and HTML version of your message
    text = "Myszowy test majla grupowego."

    footer_tomasz = """\

    <html>
        <head>
            <meta name="viewport" content="width=device-width,initial-scale=1,user-scalable=no">
        </head>
        <body style="margin: 0; padding: 20px;">
        <div style="font: normal normal normal 12px/14px Helvetica; color: #000; padding: 5px 0;">Pozdrawiam / Kind regards</div>
            <div style="padding: 10px 0;">
                <div style="display: flex; flex-wrap: wrap; align-items: center; margin: 0 -20px;">
                    <div style="width: 340px;">
                        <div style="padding: 0 20px">						
                            <div style="font: normal normal bold 20px/24px Helvetica; color: #000; padding: 5px 0;">Tomasz Czerwiński</div>
                            <div style="font: normal normal normal 14px/16px Helvetica; color: #000;">Junior Software Tester</span></div>
                        </div>
                    </div>
                    <div style="width: 330px;display: flex;align-items: center;align-self: stretch;height: 56px;border-left: 1px solid #000;margin-left: -1px;">
                        <div style="padding-left: 20px;">
                                <a style="display: block; text-decoration: none; font: normal normal normal 11px/12px Helvetica; color: #000; padding-top: 5px;" href="mailto:tomasz.czerwinski@cadprojekt.com.pl">tomasz.czerwinski@cadprojekt.com.pl</a>
                            <a style="display: block; text-decoration: none; font: normal normal normal 11px/12px Helvetica; color: #000; padding-top: 5px;" href="mailto:testy@cadprojekt.com.pl">testy@cadprojekt.com.pl</a>
                        </div>
                    </div>
                    <div style="padding: 0 20px; align-self: flex-end;">
                        <img style="height: 24px; width: auto; display: block; margin-top: 10px;" src="https://cadprojekt.com.pl/zasoby/email/logo_spjawna.png" alt="CAD PROJEKT">
                    </div>
                </div>
            </div>
            <div style="font: normal normal normal 10px/14px Helvetica; color: #666; padding: 10px 0; width: 815px; max-width: 100%; border-top: 3px solid #F3951A;">
                CAD Projekt K&A Sp.j. Dąbrowski, Sterczała, Sławek  | Poznański Park Naukowo-Technologiczny <br>
                ul. Rubież 46 | 61-612 Poznań | Poland | NIP 779-00-34-266 | REGON 632223660 | <a style="text-decoration: none; color: #666" href="https://www.cadprojekt.com.pl">www.cadprojekt.com.pl</a> 
            </div>
        
            <a style="display: block; font: normal normal normal 10px/14px Helvetica; color: #666; padding: 10px 0; text-decoration: none;" href="https://cadprojekt.com.pl/polityka-prywatnosci-sp-j" target="_blank">
                Zapoznaj się z zasadami przetwarzania Twoich danych osobowych.
            </a>
        </body>
    </html>

    """

    # Turn these into plain/html MIMEText objects
    part1 = MIMEText(text, "plain")
    part2 = MIMEText(footer_tomasz, "html")

    # Add HTML/plain-text parts to MIMEMultipart message
    # The email client will try to render the last part first
    message.attach(part1)
    message.attach(part2)

    # Create secure connection with server and send email
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("poczta.cadprojekt.com.pl", 465, context=context) as server:
        server.login(sender_email, password)
        server.sendmail(
            sender_email, emails, message.as_string()
        )

def get_user():
    print(var_0.get())

def get_receivers():
    print(var_1.get())

def get_version_number(file_path):
    try:
        File_information = GetFileVersionInfo(file_path, "\\")
        ms_file_version = File_information["FileVersionMS"]
        ls_file_version = File_information["FileVersionLS"]

        list_of_versions = [
            str(HIWORD(ms_file_version)),
            str(LOWORD(ms_file_version)),
            str(HIWORD(ls_file_version)),
            str(LOWORD(ls_file_version)),
        ]
        return ".".join(list_of_versions)
    except:
        return "-"


app = CTk()

app.title("email sender")
app.geometry("380x400")
app.eval('tk::PlaceWindow . center')

var_0 = IntVar()
var_1 = IntVar()

f = open('//home//czerwiec//Pulpit//VS Code workplaces//csv_reader//data.json')
data = json.load(f)
print(data["users"][0])
print(data["users"][0]["displayname"])

if os.path.isfile(pack_path + "\\" + odt_name):
    odt_file = pack_path + "\\" + odt_name
    label01 = CTkLabel(app, text=odt_name, fg_color="transparent", font=("Consolas", 14))
    label01.pack(padx=(0,0), pady=(10,0), anchor=CENTER)
else:
    label01 = CTkLabel(app, text="Dodaj plik z listą", fg_color="transparent", font=("Consolas", 14))
    label01.pack(padx=(0,0), pady=(10,0), anchor=CENTER)
    is_file = False

e_button01 = CTkButton(app, text="Dodaj listę", width=160, height=40, font=("Consolas", 14))
e_button01.pack(padx=(0,0), pady=(15,20), anchor=CENTER)


e_frame = CTkFrame(app)
e_frame.configure(border_width=2, fg_color="transparent")
e_frame.pack()


e_radiobutton01 = CTkRadioButton(e_frame, text="Kinga", variable=var_0, value=1, font=("Consolas", 14), command=get_user)
e_radiobutton01.pack(pady=20, padx=(30, 5), side="left", anchor=N)
if is_file == False:
    e_radiobutton01.configure(state="disabled")

e_radiobutton02 = CTkRadioButton(e_frame, text="Tomasz", variable=var_0, value=2, font=("Consolas", 14), command=get_user)
e_radiobutton02.pack(pady=20, padx=(5, 15), side="left", anchor=N)
if is_file == False:
    e_radiobutton02.configure(state="disabled")


e_frame02 = CTkFrame(app)
e_frame02.configure(border_width=2, fg_color="transparent")
e_frame02.pack(pady=(20,0), padx=(0,0))

e_radiobutton03 = CTkRadioButton(e_frame02, text="Serwis", variable=var_1, value=1, font=("Consolas", 14), command=get_receivers)
e_radiobutton03.pack(pady=20, padx=(30, 5), side="left", anchor=N)
if is_file == False:
    e_radiobutton03.configure(state="disabled")

e_radiobutton04 = CTkRadioButton(e_frame02, text="Wszyscy", variable=var_1, value=2, font=("Consolas", 14), command=get_receivers)
e_radiobutton04.pack(pady=20, padx=(5, 15), side="left", anchor=N)
if is_file == False:
    e_radiobutton04.configure(state="disabled")


x = pack_path + "\\" + "MainFiles\\V4_I10x64\\kafle.dll"

if os.path.isfile(x):
    label02 = CTkLabel(app, text="Wersja programu: " + get_version_number(x), fg_color="transparent", font=("Consolas", 14))
else:
    label02 = CTkLabel(app, text="Wersja programu: " + "bez zmian.", fg_color="transparent", font=("Consolas", 14))

label02.pack(padx=(0,0), pady=(15,0), anchor=CENTER)

e_button02 = CTkButton(app, text="Wyślij email", width=160, height=40, font=("Consolas", 14), command=send_email)
e_button02.pack(padx=(0,0), pady=(20,20), anchor=CENTER)
if is_file == False:
    e_button02.configure(state="disabled")


app.mainloop()