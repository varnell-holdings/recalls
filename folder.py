# import win32com.client as win32  # pip install pywin32
import argparse
import csv
import datetime
from dateutil.parser import parse
import os
import re
import shutil
import time

from tkinter import ttk, StringVar, Tk, W, E, N, S
from tkinter import FALSE, Menu, Frame, messagebox


from jinja2 import Environment, FileSystemLoader

# import pyautogui as pya
# import pyperclip
from pyisemail import is_email

# pya.PAUSE = 0.1

# Get today's date
today = datetime.date.today()

parser = argparse.ArgumentParser(description="This is a test script")
parser.add_argument("-t", "--test", action="store_true", help="Run the test")
args = parser.parse_args()
if args.test:
    print("Test mode activated")
    csv_address = "d:\\john tillet\\source\\active\\recalls\\test_csv.csv"
    csv_address_2 = "D:\\Nobue\\test_recalls_csv.csv"

else:
    print("No test mode activated")
    csv_address = "D:\\JOHN TILLET\\source\\active\\recalls\\recalls_csv.csv"
    csv_address_2 = "D:\\Nobue\\recalls_csv.csv"


full_path = ""
print_length = 0
pat = []  # eg ['Mr Alan MATHISON', 'Stoita', '9998877', 'Colonoscopy']
email = ""
phone = ""
mrn = ""
dob = ""


ocd_doc_set = {
    "Bariol",
    "Feller",
    "Stoita",
    "Mill",
}
# non_ocd_doc_set = {"Sanagapalli", "Williams",
#                    "Wettstein", "Vivekanandahrajah", "Ghaly"}

DOCTORS = [
    "Bariol",
    "Feller",
    "Ghaly",
    "Mill",
    "Sanagapalli",
    "Stoita",
    "Vivekanandahrajah",
    "Wettstein",
    "Williams",
]


full_doc_dict = {
    "Bariol": "Carolyn",
    "Feller": "Robert",
    "Stoita": "Alina",
    "Mill": "Justine",
    "Sanagapalli": "Santosh",
    "Williams": "David",
    "Wettstein": "Antony",
    "Vivekanandahrajah": "Suhirdan",
    "Ghaly": "Simon",
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
    EMAIL_POS = (380, 510)
elif user == "John2":
    RED_BAR_POS = (160, 630)
    TITLE_POS = (200, 134)
    MRN_POS = (600, 250)
    POST_CODE_POS = (490, 284)
    DOB_POS = (600, 174)
    FUND_NO_POS = (580, 548)
    CLOSE_POS = (774, 96)
    SMS_POS = (360, 630)
    EMAIL_POS = (335, 395)
elif user == "Recept5":
    RED_BAR_POS = (160, 630)
    TITLE_POS = (200, 134)
    MRN_POS = (575, 250)
    POST_CODE_POS = (480, 280)
    DOB_POS = (600, 174)
    FUND_NO_POS = (580, 548)
    CLOSE_POS = (780, 96)
elif user == "Typing2":
    RED_BAR_POS = (160, 630)
    TITLE_POS = (200, 134)
    MRN_POS = (575, 250)
    POST_CODE_POS = (480, 280)
    DOB_POS = (600, 174)
    FUND_NO_POS = (580, 548)
    CLOSE_POS = (780, 96)


def open_bc():
    pass
    # pya.moveTo(100, 450, duration=0.3)
    # pya.click()
    # pya.hotkey("ctrl", "o")


def scraper(email=False):
    """'"""
    pass
    # result = "na"
    # pya.hotkey("ctrl", "c")
    # result = pyperclip.paste()
    # if email:
    #     result = re.split(r"[\s,:/;\\]", result)[0]
    #     if not is_email(result):
    #         result = ""
    # return result


def scrape():
    pass
    # global mrn
    # global email
    # global dob

    # pya.moveTo(100, 450, duration=0.1)
    # pya.click()
    # pya.press("up", presses=9)
    # pya.press("enter")
    # time.sleep(0.5)

    # pya.moveTo(EMAIL_POS)
    # pya.doubleClick()
    # email = scraper()
    # print(email)

    # pya.moveTo(MRN_POS)
    # pya.doubleClick()
    # mrn = scraper()
    # print(mrn)

    # pya.moveTo(DOB_POS)
    # pya.doubleClick()
    # dob = scraper()
    # print(dob)

    # if (not mrn.isdigit()) or (not parse(dob)):
    #     scrape_info_label.set("Error in data")
    #     root.update_idletasks()
    #     return

    # if is_email(email) and email not in {"", "na"}:
    #     scrape_info_label.set("OK")
    # else:
    #     scrape_info_label.set("Problem with email")
    # root.update_idletasks()


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

    today_str = today.strftime("%d-%m-%Y")

    over_75 = is_over_75(dob)

    path_to_template = "D:\\JOHN TILLET\\source\\active\\recalls"
    loader = FileSystemLoader(path_to_template)
    env = Environment(loader=loader)
    template_name = "body_1_template.html"
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


def letter_compose():
    pass


def send_text():
    full_name = pat[0]
    title = full_name.split()[0]
    last_name = full_name.split()[-1].title()
    doctor = pat[1]
    message = f"Dear {title} {last_name} just advising you that an email will be sent to you from Dr {doctor} with a reminder that you are now due for your procedure. Please review email and contact our office on 83826622 if you have any queries and for all bookings."

    pya.moveTo(100, 450, duration=0.3)
    pya.click()
    pya.press("up", presses=3)
    pya.press("enter")
    pya.hotkey("alt", "n")
    pya.moveTo(SMS_POS[0], SMS_POS[1])
    pya.click()
    pya.typewrite(message)
    pya.press("tab")
    pya.press("enter")
    pya.press("enter")


def recall_type(event):
    recall_type = rec.get()
    if recall_type == "GP":
        button4.config(state="disabled", style="Disabled.TButton")
        send_text_button.config(state="disabled", style="Disabled.TButton")
    else:
        button4.config(state="normal", style="Normal.TButton")
        send_text_button.config(state="normal", style="Normal.TButton")
    root.update_idletasks()


def recall_compose():
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # 0 represents an email item
    mail.Subject = "Procedure reminder"

    mail.To = email

    logo_path = r"D:\\JOHN TILLET\\source\\active\\recalls\\dec_logo.jpg"
    attachment = mail.Attachments.Add(logo_path)
    CONTENT_ID_PROPERTY = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"
    our_content_id = "my_logo_123"
    attachment.PropertyAccessor.SetProperty(CONTENT_ID_PROPERTY, our_content_id)

    make_html_body(our_content_id)

    with open(
        "D:\\JOHN TILLET\\source\\active\\recalls\\body_1.html", "rt", encoding="cp1252"
    ) as f:
        html_content = f.read()

    mail.HTMLBody = html_content
    # # Uncomment to actually send the email
    # # mail.Send()

    # # Or display it for review before sending
    mail.Display()
    # write to csv
    day_sent = today.isoformat()
    with open(csv_address, "a") as f:
        writer = csv.writer(f, dialect="excel", lineterminator="\n")
        # name, doctor, phone, mrn, dob, phone, email, first, second, third, gp, attended
        # - first second and third are dates
        entry = [
            pat[0],
            pat[1],
            pat[2],
            mrn,
            dob,
            phone,
            email,
            day_sent,
            "",
            "",
            "no",
            "",
        ]
        writer.writerow(entry)
    shutil.copy(csv_address, csv_address_2)

    # config gui
    p.set("")
    # button3.config(state="normal", style="Normal.TButton")
    # button4.config(state="disabled", style="Disabled.TButton")
    # button5.config(state="disabled", style="Disabled.TButton")
    # root.update_idletasks()


def finish():
    """This will close Blue Chip and write to
    disposal.csv and reset buttons above the line"""
    doc.set("Doctor")
    rec.set("Recall Type")
    proc.set("Procedure")


root = Tk()

doc = StringVar()  # doctor for recall
rec = StringVar()  # type of recall ie 2 or 3
proc = StringVar()  # type of procedure
scrape_info_label = StringVar()

root.geometry("450x550+840+50")
root.title("Recalls from Folder")
root.option_add("*(tearOff)", FALSE)

# Create a style
style = ttk.Style()

# Configure styles for different states
style.configure("Normal.TButton", background="lightgray", foreground="blue")
style.configure("Disabled.TButton", background="lightgray", foreground="darkgray")

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


doc_box = ttk.Combobox(topframe, textvariable=doc)
doc_box["values"] = DOCTORS
doc_box["state"] = "readonly"
doc_box.grid(column=0, row=0, sticky=W)

rec_box = ttk.Combobox(topframe, textvariable=rec)
rec_box["values"] = ["2", "3"]
rec_box["state"] = "readonly"
rec_box.grid(column=0, row=1, sticky=W)

proc_box = ttk.Combobox(topframe, textvariable=proc)
proc_box["values"] = ["Pe", "Col", "Double"]
proc_box["state"] = "readonly"
proc_box.grid(column=1, row=1, sticky=W)

open_button = ttk.Button(bottomframe, text="Open Blue Chip", command=open_bc)
open_button.grid(column=0, row=0, sticky=W)


scrape_button = ttk.Button(bottomframe, text="Get info", command=scrape)
scrape_button.grid(column=0, row=1, sticky=W)

scrape_label = ttk.Label(bottomframe, textvariable=scrape_info_label)
scrape_label.grid(column=1, row=1, sticky=E)


button4 = ttk.Button(bottomframe, text="Send email", command=recall_compose)
button4.grid(column=0, row=2, sticky=W)

letter_button = ttk.Button(
    bottomframe, text="Make Recall letter", command=letter_compose
)
letter_button.grid(column=1, row=2, sticky=E)

send_text_button = ttk.Button(bottomframe, text="Send text", command=send_text)
send_text_button.grid(column=0, row=3, sticky=W)


finish_button = ttk.Button(bottomframe, text="Finish patient", command=finish)
finish_button.grid(column=0, row=4, sticky=W)

for child in topframe.winfo_children():
    child.grid_configure(padx=5, pady=20)

for child in bottomframe.winfo_children():
    child.grid_configure(padx=5, pady=20)


scrape_info_label.set("")

doc.set("Doctor")
rec.set("Recall Type")
proc.set("Procedure")

root.attributes("-topmost", True)
root.mainloop()
