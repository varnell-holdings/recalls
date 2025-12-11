import argparse
import csv
import datetime
from dateutil.parser import parse
import os
import pickle
import re
import shutil
import subprocess
import sys
import time

from tkinter import ttk, StringVar, Tk, W, E, N, S
from tkinter import FALSE, Menu, Frame  # , messagebox
from tkinter.filedialog import askopenfilename

from docx import Document
from docx.shared import Pt

from jinja2 import Environment, FileSystemLoader
from striprtf.striprtf import rtf_to_text
import pyautogui as pya
import pyperclip
from pyisemail import is_email
import win32com.client as win32  # pip install pywin32

pya.PAUSE = 0.1

# Get today's date
today = datetime.date.today()
today_str = today.strftime("%d-%m-%Y")

parser = argparse.ArgumentParser(description="Recalls - First")
parser.add_argument("-t", "--test", action="store_true", help="Run the test")
parser.add_argument("-n", "--nopickle", action="store_true",
                    help="No pickle option")
args = parser.parse_args()
if args.test:
    print("Test mode activated")
    csv_address = "d:\\john tillet\\source\\active\\recalls\\csv\\test_csv.csv"
    csv_address_2 = "D:\\Nobue\\test_recalls_csv.csv"

else:
    print("Not in  test mode")
    csv_address = "D:\\JOHN TILLET\\source\\active\\recalls\\csv\\recalls_csv.csv"
    csv_address_2 = "D:\\Nobue\\recalls_csv.csv"


if args.nopickle:
    print("No pickling mode")
else:
    pickle_address = "D:\\JOHN TILLET\\source\\active\\recalls\\pickled_list"

full_path = ""
old_path = ""
num_to_do = 0
pat = []  # eg ['Mr Alan MATHISON', 'Stoita', '0432-876-980', 'Colonoscopy']
email = ""
phone = ""
mrn = ""
dob = ""
first_run = True
manual = False


output_list_4 = []

ocd_doc_set = {
    "Bariol",
    "Feller",
    "Stoita",
    "Mill",
}


full_doc_dict = {
    "Bariol": "Carolyn",
    "Feller": "Robert",
    "Stoita": "Alina",
    "Mill": "Justine",
    "Sanagapalli": "Santosh",
    "Williams": "David",
    "Wettstein": "Antony",
    "Vivekanandarajah": "Suhirdan",
    "Ghaly": "Simon",
    "Vickers": "Chris",
}

proc_dict = {"Colonoscopy": "c", "COL/PE": "d", "Panendoscopy": "p"}

user = os.getenv("USERNAME")

if user == "John":
    RED_BAR_POS = (280, 790)
    TITLE_POS = (230, 170)
    MRN_POS = (740, 315)
    POST_CODE_POS = (610, 355)
    DOB_POS = (750, 220)
    FUND_NO_POS = (770, 703)
    CLOSE_POS = (1020, 120)
    SMS_POS = (485, 735)
    EMAIL_POS = (600, 510)
elif user == "John2":
    RED_BAR_POS = (160, 630)
    TITLE_POS = (200, 134)
    MRN_POS = (600, 250)
    POST_CODE_POS = (490, 284)
    DOB_POS = (600, 174)
    FUND_NO_POS = (580, 548)
    CLOSE_POS = (774, 96)
    SMS_POS = (360, 630)
    EMAIL_POS = (450, 395)
elif user == "Typing2":
    RED_BAR_POS = (160, 630)
    TITLE_POS = (200, 134)
    MRN_POS = (575, 250)
    POST_CODE_POS = (480, 280)
    DOB_POS = (600, 174)
    FUND_NO_POS = (580, 548)
    CLOSE_POS = (780, 96)
    SMS_POS = (360, 630)
    EMAIL_POS = (450, 400)


def get_pickled_list():
    global output_list_4
    with open(pickle_address, "rb") as f:
        output_list_4 = pickle.load(f)


def set_pickled_list():
    global output_list_4
    with open(pickle_address, "wb") as f:
        pickle.dump(output_list_4, f)


def collect_file():
    global full_path
    global old_path
    full_path = askopenfilename()

    f.set("    Open Blue Chip now.")
    old_path = os.path.join("D:\\john tillet\\source\\active\\recalls\\old")
    button2.config(state="normal", style="Normal.TButton")
    root.update_idletasks()


def extract():
    global output_list_4
    global num_to_do   # int
    text2 = ""

    with open(full_path, "r") as rtf_file:
        rtf_content = rtf_file.read()

    text = rtf_to_text(rtf_content)

    for i, line in enumerate(text.splitlines()):
        if i == 0:
            continue
        elif "|||||" in line or not line:
            continue
        else:
            text2 += line
            text2 += "\n"

    for line in text2.splitlines():
        local_list = []
        for i, field in enumerate(line.split("|")):
            if i == 0:
                local_list.append(field)
            elif i == 1:
                doc_name = field.split()[2]
                if doc_name == "Vickers":
                    doc_name = "Mill"
                local_list.append(doc_name)
            elif i == 2:
                local_list.append(field)
            elif i == 3:
                local_list.append(field)
        if local_list != [""]:
            output_list_4.append(local_list)

    output_list_4.reverse()
    print(output_list_4)
    num_to_do = len(output_list_4)
    n.set(f"{num_to_do} patients to do.")
    filename = os.path.splitext(os.path.basename(full_path))[0]
    f.set(f"Working on {filename}.")
    if not args.nopickle:
        set_pickled_list()
    print("line 191")
    button1.config(state="disabled", style="Disabled.TButton")
    button2.config(state="disabled", style="Disabled.TButton")
    next_patient()


def next_patient():
    global full_path
    global num_to_do
    global output_list_4
    global pat
    global phone
    global first_run
    global manual
    if not args.nopickle:
        get_pickled_list()  # this gets output_list_4
    try:
        pat = output_list_4.pop()
        manual = False
        

    except IndexError:
        p.set("Finished!")
        num_to_do = 0
        n.set(f"{num_to_do} patients to do.")
        # shutil.move(full_path, old_path)
        full_path = ""
        button1.config(state="normal", style="Normal.TButton")
        button2.config(state="normal", style="Normal.TButton")

        output_list_4 = []
    
    print(pat)
    name = pat[0]
    phone = pat[2]
    name_for_label = f"{name}\n{phone}"
    p.set(name_for_label)

    num_to_do = len(output_list_4)
    n.set(f"{num_to_do} patients to do.")
    scrape_info_label.set("")
    root.update_idletasks()
    print("line 218")
    phone = pat[2].replace("-", "")
    
    if first_run:
        pya.alert("Make sure Blue Chip is open then press OK.")
        first_run = False
        if phone and phone[0] == "0":
            open_bc_by_phone()
        else:
            open_bc_by_name()
    else:
        if phone and phone[0] == "0":
            open_bc_by_phone()
        else:
            open_bc_by_name()


def open_bc_by_name():
    global pat
    name_as_list = pat[0].split()
    name_for_bc = name_as_list[-1] + "," + name_as_list[1]
    pass
    pya.moveTo(100, 450, duration=0.3)
    pya.click()
    pya.hotkey("ctrl", "o")
    pya.typewrite(name_for_bc)
    pya.press("enter")
    time.sleep(1.6)
    pya.press("enter")
    pya.moveTo(100, 450, duration=0.3)
    pya.click()
    pya.press("up", presses=4)
    pya.press("enter")


def open_bc_by_name_short():
    """Stops at list of names."""
    global pat
    name_as_list = pat[0].split()
    name_for_bc = name_as_list[-1] + "," + name_as_list[1]
    pass
    pya.moveTo(100, 450, duration=0.3)
    pya.click()
    pya.hotkey("ctrl", "o")
    pya.typewrite(name_for_bc)
    pya.press("enter")


def open_bc_by_phone():
    phone = pat[2].replace("-", "")
    pya.moveTo(100, 450, duration=0.3)
    pya.click()
    pya.hotkey("ctrl", "o")
    pya.hotkey("alt", "b")
    time.sleep(0.2)

    pya.press("down", presses=4)

    pya.hotkey("shift", "tab")
    pya.hotkey("shift", "tab")
    pya.typewrite(phone)
    pya.press("enter")
    pya.moveTo(100, 450, duration=0.3)
    pya.click()
    pya.press("up", presses=4)
    pya.press("enter")


def scraper(email=False):
    """'"""
    result = "na"
    pya.hotkey("ctrl", "c")
    result = pyperclip.paste()
    if email:
        result = re.split(r"[\s,:/;\\*]", result)[0]
        for object in result:
            if is_email(object):
                return object
            return ""
    return result


def scrape():
    global mrn
    global email
    global dob

    scrape_info_label.set("")
    root.update_idletasks()

    pya.moveTo(100, 450, duration=0.1)
    pya.click()
    pya.press("up", presses=9)
    pya.press("enter")
    time.sleep(0.5)

    pya.moveTo(EMAIL_POS)
    pya.doubleClick()
    email = scraper()
    print(email)

    pya.moveTo(MRN_POS)
    pya.doubleClick()
    mrn = scraper()
    print(mrn)

    pya.moveTo(DOB_POS)
    pya.doubleClick()
    dob = scraper()
    print(dob)

    if not mrn.isdigit():
        scrape_info_label.set("\u274C Error in MRN\nTry again.")
        root.update_idletasks()
        return
    elif not parse_dob():
        scrape_info_label.set("\u274C Error in DOB\nTry again.")
        root.update_idletasks()
        return
    elif (not is_email(email)) or (email in {"", "na"}):
        scrape_info_label.set("\u274C Problem with email")
        root.update_idletasks()
        return
    else:
        scrape_info_label.set("\u2705")
        root.update_idletasks()

        scrape_info_label.set("Sending text")
        root.update_idletasks()
        send_text()

        recall_compose()
        return


def parse_dob():
    try:
        parse(dob, dayfirst=True)
        return True
    except Exception:
        return False


def is_over_75(date_of_birth):
    """
    Check if a person is 75 years old or older based on their date of birth.

    Args:
        date_of_birth (str): Date of birth in format "dd/mm/yyyy"

    Returns:
        bool: True if aged 75 or over, False otherwise
    """
    # Parse the date of birth
    dob = datetime.datetime.strptime(date_of_birth, "%d/%m/%Y")

    # Calculate age
    age = today.year - dob.year

    # Adjust if birthday hasn't occurred this year yet
    if (today.month, today.day) < (dob.month, dob.day):
        age -= 1

    return age > 75


def make_html_body(our_content_id):
    """Uses jinja2 to construct the html body of the email. Saves in body_.html"""
    full_name = pat[0]
    title = full_name.split()[0]
    first_name = full_name.split()[1]
    last_name = full_name.split()[-1].title()
    full_name = f"{title} {first_name} {last_name}"

    doctor = pat[1]

    doc_first_name = full_doc_dict[doctor]

    if doctor in ocd_doc_set:
        ocd = True
    else:
        ocd = False

    procedure = pat[3]
    if procedure == "COL/PE":
        procedure = "Gastroscopy and Colonoscopy"
    if procedure == "Panendoscopy":
        procedure = "Gastroscopy"

    over_75 = is_over_75(dob)

    path_to_template = "D:\\JOHN TILLET\\source\\active\\recalls\\templates"
    loader = FileSystemLoader(path_to_template)
    env = Environment(loader=loader)
    template_name = "email_1_template.html"
    template = env.get_template(template_name)
    page = template.render(
        today_date=today_str,
        full_name=full_name,
        title=title,
        last_name=last_name,
        doc_first_name=doc_first_name,
        doctor=doctor,
        procedure=procedure,
        over_75=over_75,
        ocd=ocd,
        our_content_id=our_content_id,
    )

    with open("D:\\JOHN TILLET\\source\\active\\recalls\\body_1.html", "wt") as f:
        f.write(page)


def write_csv(attended):
    global email
    day_sent = today.isoformat()
    name = pat[0]
    doctor = pat[1]
    procedure = pat[3]
    if not is_email(email):
        email = ""
    if attended == "yes":
        first = ""
    else:
        first = day_sent
    with open(csv_address, "a") as f:
        writer = csv.writer(f, dialect="excel", lineterminator="\n")
        # name, doctor, mrn, dob, procedure, email, first, second, third, attended
        # - first second and third are dates
        entry = [
            name,
            doctor,
            mrn,
            dob,
            procedure,
            email,
            first,
            "",
            "",
            attended,
        ]
        writer.writerow(entry)
    shutil.copy(csv_address, csv_address_2)


def recall_compose():
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # 0 represents an email item
    mail.Subject = "Procedure reminder"

    mail.To = email

    logo_path = r"D:\\JOHN TILLET\\source\\active\\recalls\\dec_logo.jpg"
    attachment = mail.Attachments.Add(logo_path)
    CONTENT_ID_PROPERTY = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
    our_content_id = "my_logo_123"
    attachment.PropertyAccessor.SetProperty(
        CONTENT_ID_PROPERTY, our_content_id)

    make_html_body(our_content_id)

    with open(
        "D:\\JOHN TILLET\\source\\active\\recalls\\body_1.html", "rt", encoding="cp1252"
    ) as f:
        html_content = f.read()

    mail.HTMLBody = html_content
    if not manual:
        mail.Send()
    else:
        mail.Display()
    # write to csv
    write_csv(attended="no")

    # config
    scrape_info_label.set("Email made")
    root.update_idletasks()


def make_letter_text(pat, dob):
    full_name = pat[0]
    title = full_name.split()[0]
    first_name = full_name.split()[1]
    last_name = full_name.split()[-1].title()
    full_name = f"{title} {first_name} {last_name}"

    doctor = pat[1]
    procedure = pat[3]
    if procedure == "COL/PE":
        procedure = "Gastroscopy and Colonoscopy"
    if procedure == "Panendoscopy":
        procedure = "Gastroscopy"

    over_75 = is_over_75(dob)

    if doctor in ocd_doc_set:
        ocd = True
    else:
        ocd = False

    path_to_template = "D:\\JOHN TILLET\\source\\active\\recalls\\templates"
    loader = FileSystemLoader(path_to_template)
    env = Environment(loader=loader)
    template_name = "letter_1_template.txt"
    template = env.get_template(template_name)
    page = template.render(
        today_date=today_str,
        full_name=full_name,
        title=title,
        last_name=last_name,
        doctor=doctor,
        procedure=procedure,
        over_75=over_75,
        ocd=ocd,
    )
    return page


def letter_compose():
    global recall_type
    recall_type = "letter"
    full_name = pat[0]
    title = full_name.split()[0]
    first_name = full_name.split()[1]
    last_name = full_name.split()[-1].title()
    full_name = f"{title} {first_name} {last_name}"
    doctor = pat[1]

    page = make_letter_text(pat, dob)

    doc = Document(
        f"d:\\john tillet\\source\\active\\recalls\\headers\\{doctor}.docx")

    style = doc.styles["Normal"]
    font = style.font
    font.name = "Bookman Old Style"
    font.size = Pt(12)

    paragraph = doc.add_paragraph()

    # Split the text and add line breaks
    lines = page.split("\n")

    for i, line in enumerate(lines):
        print(i, line)
        run = paragraph.add_run(line)
        if i < len(lines) - 1:  # Don't add break after last line
            run.add_break()

    doc.save(f"d:\\john tillet\\source\\active\\recalls\\letters\\{
             last_name}.docx")
    write_csv(attended="no")
    scrape_info_label.set("Letter made\nCancel manually")
    root.update_idletasks()
    os.startfile(f"d:\\john tillet\\source\\active\\recalls\\letters\\{
        last_name}.docx")


def send_text():
    p.set("")
    full_name = pat[0]
    title = full_name.split()[0]
    last_name = full_name.split()[-1].title()
    doctor = pat[1]
    message = f"Dear {title} {last_name} just advising you that an email will be sent to you from Dr {
        doctor} with a reminder that you are now due for your procedure. Please review email and contact our office on 83826622 if you have any queries and for all bookings."

    pya.moveTo(100, 450, duration=0.3)
    pya.click()
    pya.press("up", presses=3)
    pya.press("enter")
    pya.hotkey("alt", "n")
    pya.moveTo(SMS_POS[0], SMS_POS[1])
    pya.click()
    pya.typewrite(message)
    time.sleep(3)
    pya.press("tab")
    pya.press("enter")
    time.sleep(3)
    pya.press("enter")
    scrape_info_label.set("Text sent - open Outlook")
    root.update_idletasks()

def manual_send():
    global manual
    manual = True
    send_text()
    recall_compose()
    return


def no_recall():
    global recall_number
    global recall_type
    write_csv(attended="yes")
    scrape_info_label.set("No recall sent -> finish")
    root.update_idletasks()
    recall_number = "none"
    recall_type = "none"
    pya.click(100, 400)
    pya.press("up", presses=3)
    pya.press("enter")
    pya.hotkey("alt", "n")
    pya.press("enter")

    p.set("")
    finish_recall()


def close_out():
    if not args.nopickle:
        set_pickled_list()
    full_name = pat[0]
    scrape_info_label.set(f"{full_name} finished")
    root.update_idletasks()
    pya.moveTo(CLOSE_POS[0], CLOSE_POS[1])
    pya.click()
    pya.hotkey("alt", "n")


def finish_recall():
    close_out()
    if output_list_4:
        # button3.config(state="normal", style="Normal.TButton")
        p.set("")
        next_patient()
    else:
        p.set("Finished")


def finish_exit():
    close_out()
    sys.exit(1)


def open_letters():
    os.startfile("d:\\john tillet\\source\\active\\recalls\\letters")


def reset_program():
    subprocess.run(
        [sys.executable, "D:\\JOHN TILLET\\source\\active\\recalls\\pickler.py"])
    sys.exit(1)


if not args.nopickle:
    get_pickled_list()  # this gets output_list_4

root = Tk()

f = StringVar()  # label to show which blue chip file we are working on
n = StringVar()  # number of patients left to process
p = StringVar()  # current patient name
scrape_info_label = StringVar()

root.geometry("350x450+840+50")
root.title("First Recalls")
root.option_add("*(tearOff)", FALSE)

menubar = Menu(root)
root.config(menu=menubar)
menu_extras = Menu(menubar)
menubar.add_cascade(menu=menu_extras, label="Extras")
menu_extras.add_command(label="Letters folder", command=open_letters)
menu_extras.add_command(label="Reset Program", command=reset_program)

# Create a style
style = ttk.Style()

# Configure styles for different states
style.configure("Normal.TButton", background="lightgray", foreground="blue")
style.configure("Disabled.TButton", background="lightgray",
                foreground="darkgray")

mainframe = ttk.Frame(root, padding="3 3 12 12")
mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
mainframe.columnconfigure(0, weight=1)
mainframe.rowconfigure(0, weight=1)

topframe = Frame(mainframe)
topframe.grid(column=0, row=0, sticky=(N, W, E, S))
topframe.columnconfigure(0, weight=1)
topframe.rowconfigure(0, weight=1)

midframe = Frame(mainframe, height=2, bg="black")
midframe.grid(column=0, row=2, sticky=(N, W, E, S))
midframe.columnconfigure(0, weight=1)
midframe.rowconfigure(0, weight=1)

bottomframe = Frame(mainframe)
bottomframe.grid(column=0, row=3, sticky=(N, W, E, S))
bottomframe.columnconfigure(0, weight=1)
bottomframe.rowconfigure(0, weight=1)


button1 = ttk.Button(topframe, text="Open Blue Chip File",
                     command=collect_file)
button1.grid(column=0, row=0, sticky=W)

label1 = ttk.Label(topframe, textvariable=f)
label1.grid(column=1, row=0, sticky=E)

button2 = ttk.Button(topframe, text="Create Datafile", command=extract)
button2.grid(column=0, row=1, sticky=W)

label2 = ttk.Label(topframe, textvariable=n)
label2.grid(column=1, row=1, sticky=E)


label3 = ttk.Label(bottomframe, textvariable=p)
label3.grid(column=0, row=0, sticky=W)

open_by_name_button = ttk.Button(
    bottomframe, text="Open by name", command=open_bc_by_name_short
)
open_by_name_button.grid(column=1, row=0, sticky=E)


scrape_button = ttk.Button(bottomframe, text="Text & Email", command=scrape)
scrape_button.grid(column=0, row=1, sticky=W)

scrape_label = ttk.Label(bottomframe, textvariable=scrape_info_label)
scrape_label.grid(column=0, row=2, sticky=E)


manual_button = ttk.Button(
    bottomframe, text="Manual Send", command=manual_send
)
manual_button.grid(column=1, row=1, sticky=E)

letter_button = ttk.Button(
    bottomframe, text="Recall letter", command=letter_compose
)
letter_button.grid(column=1, row=2, sticky=E)


button5 = ttk.Button(bottomframe, text="No Recall", command=no_recall)
button5.grid(column=1, row=3, sticky=E)

button_6 = ttk.Button(bottomframe, text="Finish & new", command=finish_recall)
button_6.grid(column=0, row=3, sticky=W)

button_7 = ttk.Button(bottomframe, text="Finish & exit", command=finish_exit)
button_7.grid(column=1, row=4, sticky=E)

for child in mainframe.winfo_children():
    child.grid_configure(padx=5, pady=10)

for child in bottomframe.winfo_children():
    child.grid_configure(padx=5, pady=20)


if output_list_4:
    button1.config(state="disabled", style="Disabled.TButton")
    button2.config(state="disabled", style="Disabled.TButton")
    num_to_do = len(output_list_4)
    n.set(f"{num_to_do} patients to do.")
    next_patient()
else:
    button1.config(state="normal", style="Normal.TButton")
    button2.config(state="normal", style="Normal.TButton")
    n.set("")


scrape_info_label.set("")

root.attributes("-topmost", True)
root.mainloop()
