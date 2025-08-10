# import tkinter as tk
from tkinter import ttk, StringVar, Tk, W, E, N, S
import csv
import datetime

import os
import re
from tkinter import FALSE, Menu, Frame, messagebox
from tkinter.filedialog import askopenfilename

from docx import Document  # pip  install python-docs
from docx.shared import Pt
import mammoth
import pyautogui as pya
import pyperclip
from pyisemail import is_email

pya.PAUSE = 0.6

import win32com.client as win32  # pip install pywin32


full_path = ""
print_length = 0
pat = []  # eg ['Mr Alan MATHISON', 'Stoita', 'Colonoscopy']
email = ""
mrn = ""
output_list_1 = []
output_list_2 = []
output_list_3 = []
doc_dict = {
    "Bariol": "cb",
    "Feller": "rf",
    "Sanagapalli": "ss",
    "Williams": "dw",
    "Stoita": "as",
    "Wettstein": "aw",
    "Vivekanandahrajah": "sv",
    "Mill": "jm",
    "Ghaly": "sg",
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
elif user == "John2":
    RED_BAR_POS = (160, 630)
    TITLE_POS = (200, 134)
    MRN_POS = (600, 250)
    POST_CODE_POS = (490, 284)
    DOB_POS = (600, 174)
    FUND_NO_POS = (580, 548)
    CLOSE_POS = (774, 96)
elif user == "Regina":
    RED_BAR_POS = (160, 630)
    TITLE_POS = (200, 134)
    MRN_POS = (575, 250)
    POST_CODE_POS = (480, 280)
    DOB_POS = (600, 174)
    FUND_NO_POS = (580, 548)
    CLOSE_POS = (774, 96)
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
    print("Button 1 clicked")
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
        print(output_list_2)
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

    print(output_list_3)
    print_length = len(output_list_3)
    num_to_do.set(str(print_length))
    n.set(f"{print_length} patients to do.")

    button1.config(state="disabled", style="Disabled.TButton")
    button2.config(state="disabled", style="Disabled.TButton")
    button3.config(state="normal", style="Normal.TButton")
    button4.config(state="disabled", style="Disabled.TButton")
    button5.config(state="disabled", style="Disabled.TButton")
    button6.config(state="disabled", style="Disabled.TButton")
    button7.config(state="disabled", style="Disabled.TButton")
    root.update_idletasks()


def next_patient():
    global full_path
    global print_length
    global output_list_1
    global output_list_2
    global output_list_3
    global pat
    try:
        pat = output_list_3.pop()
        print(pat)
        # name_for_label = name_as_list[-1]
        name_for_label = f"{pat[0]}  {pat[2]}"
        p.set(name_for_label)
        print(f"Print length - {print_length}")
        num_to_do.set(str(print_length))
        n.set(f"{print_length} patients to do.")
        print_length -= 1
        button3_label.set("Working On")
        button3.config(state="disabled", style="Disabled.TButton")
        button4.config(state="normal", style="Normal.TButton")
        button7.config(state="normal", style="Normal.TButton")
        root.update_idletasks()

    except IndexError:
        p.set("Finished!")
        num_to_do.set("0")
        n.set(f"{print_length} patients to do.")
        os.remove(full_path)
        full_path = ""
        button1.config(state="normal", style="Normal.TButton")
        button2.config(state="normal", style="Normal.TButton")
        button3.config(state="disabled", style="Disabled.TButton")
        button4.config(state="disabled", style="Disabled.TButton")
        button5.config(state="disabled", style="Disabled.TButton")
        button6.config(state="disabled", style="Disabled.TButton")
        button7.config(state="disabled", style="Disabled.TButton")
        root.update_idletasks()
        output_list_1 = []
        output_list_2 = []
        output_list_3 = []


def open_bc():
    global pat
    name_as_list = pat[0].split()
    name_for_bc = name_as_list[-1] + "," + name_as_list[1]
    pass
    pya.moveTo(100, 450, duration=0.3)
    pya.click()
    pya.hotkey("ctrl", "o")
    pya.typewrite(name_for_bc)
    pya.press("enter")
    pya.press("enter")
    button3.config(state="disabled", style="Disabled.TButton")
    button4.config(state="disabled", style="Disabled.TButton")
    button5.config(state="normal", style="Normal.TButton")
    button7.config(state="normal", style="Normal.TButton")
    root.update_idletasks()


def scraper(email=False):
    """'"""
    print("scraper called")
    result = "na"
    print(result)
    pya.hotkey("ctrl", "c")
    result = pyperclip.paste()
    print("after scrape", result)
    if email:
        result = re.split(r"[\s,:/;\\]", result)[0]
        if not is_email(result):
            result = ""
    print(result)
    return result


def postcode_to_state(postcode):
    post_dic = {"3": "VIC", "4": "QLD", "5": "SA", "6": "WA", "7": "TAS"}

    try:
        if postcode[0] == "0":
            if postcode[:2] in {"08", "09"}:
                return "NT"
            else:
                return ""
        elif postcode[0] in {"0", "1", "8", "9"}:
            return ""
        elif postcode[0] == "2":
            if (2600 <= int(postcode) <= 2618) or postcode[:2] == 29:
                return "ACT"
            else:
                return "NSW"
        else:
            return post_dic[postcode[0]]
    except Exception:
        return ""


def scrape():
    global mrn
    global email

    pya.moveTo(100, 450, duration=0.1)
    pya.click()
    pya.moveTo(TITLE_POS)
    pya.click()

    pya.press("tab", presses=9)
    pyperclip.copy("")
    street = scraper()
    street = street.replace(",", "")

    print(street)

    pya.press("tab")
    pya.press("tab")
    pyperclip.copy("")
    suburb = scraper()

    print(suburb)

    pya.press("tab", presses=6)
    email = scraper()

    pya.moveTo(POST_CODE_POS, duration=0.1)
    x1, y1 = POST_CODE_POS
    # disable_mouse(x1, y1, x1 + 1, y1 + 1)
    pya.doubleClick()
    postcode = scraper()

    state = postcode_to_state(postcode)

    suburb_state = f"{suburb} {state} {postcode}"

    pya.moveTo(MRN_POS)
    pya.doubleClick()
    mrn = scraper()

    print(f"Addresses {street}  {suburb_state}")

    return street, suburb_state


def make_letter(street, suburb_state):
    today = datetime.date.today()
    today_str = today.strftime("%A, %d %B %Y")
    full_name = pat[0]
    title = full_name.split()[0]
    first_name = full_name.split()[1]
    last_name = full_name.split()[-1].title()
    full_name = f"{title} {first_name} {last_name}"

    text = f"{today_str} \n\n{full_name}\n{street}\n{suburb_state}\n\nDear {pat[0].split()[0]} {pat[0].split()[-1].title()},\n\n"
    doc_abr = doc_dict[pat[1]]
    proc_abr = proc_dict[pat[3]]
    document = Document(
        f"D:\\JOHN TILLET\\source\\active\\recalls\\recall_letters\\{doc_abr}{proc_abr}1.docx"
    )

    for p in document.iter_inner_content():
        if p.text != "":  # == "Re: Your overdue procedure":
            # p.insert_paragraph_before(today_str)
            this_bit = p.insert_paragraph_before()
            this_run = this_bit.add_run(text)
            font = this_run.font
            font.name = "Bookman Old Style"
            font.size = Pt(10)
            # print(font)
            break

    document.save("D:\\JOHN TILLET\\source\\active\\recalls\\current.docx")


def recall_compose():
    # scrape details
    street, suburb_state = scrape()
    # compose recall letter
    make_letter(street, suburb_state)
    os.startfile("D:\\JOHN TILLET\\source\\active\\recalls\\current.docx")
    # config gui
    button5.config(state="normal", style="Normal.TButton")
    button6.config(state="normal", style="Normal.TButton")
    button7.config(state="disabled", style="Disabled.TButton")
    root.update_idletasks()


def send_email():
    # send email
    word_app = win32.Dispatch("Word.Application")
    word_app.Visible = False
    doc = word_app.Documents.Open(
        "D:\\JOHN TILLET\\source\\active\\recalls\\current.docx"
    )

    html_path = r"D:\\JOHN TILLET\\source\\active\\recalls\\temp_email.html"
    try:
        doc.SaveAs2(html_path, FileFormat=8)
    except Exception as e:
        print(f"{e}")
        pass
    finally:
        doc.Close()
        word_app.Quit()

    with open(html_path, "r", encoding="cp1252") as f:
        html_content = f.read()

    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)  # 0 represents an email item
    mail.Subject = "Procedure reminder"
    mail.HTMLBody = html_content
    mail.To = email

    # # Uncomment to actually send the email
    # # mail.Send()

    # # Or display it for review before sending
    mail.Display()
    # write to csv
    day_sent = datetime.date.today().isoformat()
    with open("recalls.csv", "a") as f:
        writer = csv.writer(f)
        entry = [pat[0], pat[1], pat[2], mrn, email, day_sent, "", "", "", "no"]
        writer.writerow(entry)

    # config gui
    button3_label.set("Get Next Patient")
    p.set("")
    button3.config(state="normal", style="Normal.TButton")
    button4.config(state="disabled", style="Disabled.TButton")
    button5.config(state="disabled", style="Disabled.TButton")
    button6.config(state="disabled", style="Disabled.TButton")
    button7.config(state="disabled", style="Disabled.TButton")
    root.update_idletasks()


def no_recall():
    # pya.click(100, 400)
    # pya.press("up", presses=3)
    # pya.press("enter")
    # pya.hotkey("alt", "c")
    # pyperclip.copy("")
    # message = scraper()
    # message = message + " no recall sent"
    # print(message)
    # pya.write(message)
    # pya.press("enter")
    # pya.hotkey("shift", "tab")
    # pya.press("enter")
    # pya.hotkey("alt", "m")
    # pya.press("enter")
    # pya.moveTo(CLOSE_POS[0], CLOSE_POS[1])
    # pya.click()
    # pya.hotkey("alt", "n")
    # pya.moveTo(50, 200)
    # reset button3
    button3_label.set("Get Next Patient")
    p.set("")
    button3.config(state="normal", style="Normal.TButton")
    button4.config(state="disabled", style="Disabled.TButton")
    button5.config(state="disabled", style="Disabled.TButton")
    button6.config(state="disabled", style="Disabled.TButton")
    button7.config(state="disabled", style="Disabled.TButton")
    root.update_idletasks()


root = Tk()

f = StringVar()  # label to show which blue chip file we are working on
n = StringVar()  # number of patients left to process
p = StringVar()  # current patient name
num_to_do = StringVar()
button3_label = StringVar()  # either 'next patient' or 'working on'

root.geometry("450x450+840+50")
root.title("First Recalls")
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


button1 = ttk.Button(topframe, text="Open Blue Chip File", command=collect_file)
button1.grid(column=0, row=0, sticky=W)

label1 = ttk.Label(topframe, textvariable=f)
label1.grid(column=1, row=0, sticky=E)

button2 = ttk.Button(topframe, text="Create Datafile", command=extract)
button2.grid(column=0, row=1, sticky=W)

label2 = ttk.Label(topframe, textvariable=n)
label2.grid(column=1, row=1, sticky=E)

button3 = ttk.Button(bottomframe, textvariable=button3_label, command=next_patient)
button3.grid(column=0, row=2, sticky=W)

label3 = ttk.Label(bottomframe, textvariable=p)
label3.grid(column=1, row=2, sticky=E)

button4 = ttk.Button(bottomframe, text="Open Blue Chip", command=open_bc)
button4.grid(column=0, row=3, sticky=W)

button5 = ttk.Button(bottomframe, text="Make Recall letter", command=recall_compose)
button5.grid(column=0, row=4, sticky=W)

button6 = ttk.Button(bottomframe, text="Send Email", command=send_email)
button6.grid(column=1, row=4, sticky=E)

button7 = ttk.Button(bottomframe, text="No Recall", command=no_recall)
button7.grid(column=0, row=5, sticky=W)

for child in mainframe.winfo_children():
    child.grid_configure(padx=5, pady=10)

for child in bottomframe.winfo_children():
    child.grid_configure(padx=5, pady=20)

button3_label.set("Get Next Patient")

button1.config(state="normal", style="Normal.TButton")
button2.config(state="disabled", style="Disabled.TButton")
button3.config(state="disabled", style="Disabled.TButton")
button4.config(state="disabled", style="Disabled.TButton")
button5.config(state="disabled", style="Disabled.TButton")
button6.config(state="disabled", style="Disabled.TButton")
button7.config(state="disabled", style="Disabled.TButton")

button3_label.set("Get Next Patient")

root.attributes("-topmost", True)
root.mainloop()
