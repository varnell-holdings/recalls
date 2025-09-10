import win32com.client as win32  # pip install pywin32
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
from tkinter.filedialog import askopenfilename


from jinja2 import Environment, FileSystemLoader
import mammoth

import pyautogui as pya
import pyperclip
from pyisemail import is_email

pya.PAUSE = 0.1

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
pat = []  # eg ['Mr Alan MATHISON', 'Stoita', 'Colonoscopy']
email = ""
phone = ""
mrn = ""
dob = ""

output_list_1 = []
output_list_2 = []
output_list_3 = []

ocd_doc_set = {
    "Bariol",
    "Feller",
    "Stoita",
    "Mill",
}
# non_ocd_doc_set = {"Sanagapalli", "Williams",
#                    "Wettstein", "Vivekanandahrajah", "Ghaly"}


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


def has_alpha(inputString):
    return any(char.isalpha() for char in inputString)


def has_numbers(inputString):
    return any(char.isdigit() for char in inputString)


def docx_to_text_mammoth(docx_file):
    with open(docx_file, "rb") as docx_file:
        result = mammoth.extract_raw_text(docx_file)
        return result.value


def collect_file():
    global full_path
    full_path = askopenfilename()

    filename = os.path.splitext(os.path.basename(full_path))[0]
    f.set(f"Working on {filename}")
    print(full_path)
    button2.config(state="normal", style="Normal.TButton")
    root.update_idletasks()


def extract():
    global print_length
    global output_list_1
    global output_list_2
    global output_list_3

    text_content = docx_to_text_mammoth(full_path)
    with open("output.txt", "w", encoding="utf-8") as f:
        f.write(text_content)
    with open("output.txt", "r") as f:
        for line in f:
            if (
                not (has_numbers(line) and has_alpha(line))
                and not (has_numbers(line) and "/" in line)
                and not line.isspace()
            ):
                # f2.write(f"{i}   {line}")
                output_list_1.append(line.strip())

        for i, element in enumerate(output_list_1):
            if i < 5:
                continue
            else:
                output_list_2.append(element)
        with open("output.csv", "w") as f:
            writer = csv.writer(f)
            csv_row = []
            patient = []
            for i, element in enumerate(output_list_2):
                if i % 4 == 0:
                    csv_row.append(element)
                    patient.append(element)
                elif i % 4 == 1:
                    doc = element.split()[2]
                    csv_row.append(doc)
                    patient.append(doc)
                elif i % 4 == 2:
                    csv_row.append(element)
                    patient.append(element)
                else:
                    csv_row.append(element)
                    patient.append(element)
                    writer.writerow(csv_row)
                    output_list_3.append(patient)
                    csv_row = []
                    patient = []

    output_list_3.reverse()
    print(output_list_3)
    print_length = len(output_list_3)
    num_to_do.set(str(print_length))
    n.set(f"{print_length} patients to do.")

    # button1.config(state="disabled", style="Disabled.TButton")
    # button2.config(state="disabled", style="Disabled.TButton")
    # button3.config(state="normal", style="Normal.TButton")
    # button4.config(state="disabled", style="Disabled.TButton")
    # button5.config(state="disabled", style="Disabled.TButton")
    # root.update_idletasks()


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

    # button3.config(state="disabled", style="Disabled.TButton")
    # button4.config(state="normal", style="Normal.TButton")
    # button5.config(state="normal", style="Normal.TButton")
    # root.update_idletasks()


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


def next_patient():
    global full_path
    global print_length
    global output_list_1
    global output_list_2
    global output_list_3
    global pat
    global phone
    try:
        pat = output_list_3.pop()
        print(pat)
        name = pat[0]
        phone = pat[2]
        name_for_label = f"{name}  {phone}"
        p.set(name_for_label)
        print(f"Print length - {print_length}")
        num_to_do.set(str(print_length))
        n.set(f"{print_length} patients to do.")
        print_length -= 1
        # button3.config(state="disabled", style="Disabled.TButton")
        # button5.config(state="normal", style="Normal.TButton")
        scrape_info_label.set("")
        root.update_idletasks()

    except IndexError:
        p.set("Finished!")
        num_to_do.set("0")
        n.set(f"{print_length} patients to do.")
        os.remove(full_path)
        full_path = ""
        # button1.config(state="normal", style="Normal.TButton")
        # button2.config(state="normal", style="Normal.TButton")
        # button3.config(state="disabled", style="Disabled.TButton")
        # button4.config(state="disabled", style="Disabled.TButton")
        # button5.config(state="disabled", style="Disabled.TButton")
        # root.update_idletasks()
        output_list_1 = []
        output_list_2 = []
        output_list_3 = []


def scraper(email=False):
    """'"""
    result = "na"
    pya.hotkey("ctrl", "c")
    result = pyperclip.paste()
    if email:
        result = re.split(r"[\s,:/;\\]", result)[0]
        if not is_email(result):
            result = ""
    return result


def scrape():
    global mrn
    global email
    global dob

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

    if (not mrn.isdigit()) or (not parse(dob)):
        scrape_info_label.set("Error in data")
        root.update_idletasks()
        return

    if is_email(email) and email not in {"", "na"}:
        scrape_info_label.set("OK")
    else:
        scrape_info_label.set("Problem with email")
    root.update_idletasks()


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
    pya.press("tab")
    pya.press("enter")
    pya.press("enter")


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


def no_recall():
    pya.click(100, 400)
    pya.press("up", presses=3)
    pya.press("enter")
    pya.hotkey("alt", "c")
    pyperclip.copy("")

    message = "- no recall sent"
    pya.write(message)
    pya.press("enter")
    # pya.hotkey("shift", "tab")
    # pya.press("enter")
    # pya.hotkey("alt", "m")
    # pya.press("enter")
    # pya.moveTo(CLOSE_POS[0], CLOSE_POS[1])
    # pya.click()
    # pya.hotkey("alt", "n")
    # pya.moveTo(50, 200)
    # reset button3
    p.set("")
    # button3.config(state="normal", style="Normal.TButton")
    # button4.config(state="disabled", style="Disabled.TButton")
    # button5.config(state="disabled", style="Disabled.TButton")
    # root.update_idletasks()


root = Tk()

f = StringVar()  # label to show which blue chip file we are working on
n = StringVar()  # number of patients left to process
p = StringVar()  # current patient name
num_to_do = StringVar()
scrape_info_label = StringVar()

root.geometry("450x550+840+50")
root.title("First Recalls")
root.option_add("*(tearOff)", FALSE)

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

button3 = ttk.Button(bottomframe, text="Next patient", command=next_patient)
button3.grid(column=0, row=0, sticky=W)

label3 = ttk.Label(bottomframe, textvariable=p)
label3.grid(column=1, row=0, sticky=E)

open_by_name_button = ttk.Button(
    bottomframe, text="Open by name", command=open_bc_by_name
)
open_by_name_button.grid(column=0, row=1, sticky=W)


open_by_phone_button = ttk.Button(
    bottomframe, text="Open by phone", command=open_bc_by_phone
)
open_by_phone_button.grid(column=1, row=1, sticky=E)

scrape_button = ttk.Button(bottomframe, text="Get info", command=scrape)
scrape_button.grid(column=0, row=2, sticky=W)

scrape_label = ttk.Label(bottomframe, textvariable=scrape_info_label)
scrape_label.grid(column=1, row=2, sticky=E)


button4 = ttk.Button(bottomframe, text="Send email", command=recall_compose)
button4.grid(column=0, row=3, sticky=W)

letter_button = ttk.Button(
    bottomframe, text="Make Recall letter", command=letter_compose
)
letter_button.grid(column=1, row=3, sticky=E)

send_text_button = ttk.Button(bottomframe, text="Send text", command=send_text)
send_text_button.grid(column=0, row=4, sticky=W)


button5 = ttk.Button(bottomframe, text="No Recall", command=no_recall)
button5.grid(column=1, row=4, sticky=E)

for child in mainframe.winfo_children():
    child.grid_configure(padx=5, pady=10)

for child in bottomframe.winfo_children():
    child.grid_configure(padx=5, pady=20)


# button1.config(state="normal", style="Normal.TButton")
# button2.config(state="disabled", style="Disabled.TButton")
# button3.config(state="disabled", style="Disabled.TButton")
# button4.config(state="disabled", style="Disabled.TButton")
# button5.config(state="disabled", style="Disabled.TButton")

scrape_info_label.set("")

root.attributes("-topmost", True)
root.mainloop()
