import csv
import mammoth
import os
import tkinter as tk
from tkinter.filedialog import askopenfilename

import pyautogui as pya

filename = ""
output_list_1 = []
output_list_2 = []
output_list_3 = []

def has_numbers(inputString):
    return any(char.isdigit() for char in inputString)


def docx_to_text_mammoth(docx_file):
    with open(docx_file, "rb") as docx_file:
        result = mammoth.extract_raw_text(docx_file)
        return result.value

def extract():
    global output_list_1
    global output_list_2
    global output_list_3
    
    text_content = docx_to_text_mammoth(filename)
    with open("output.txt", "w", encoding="utf-8") as f:
        f.write(text_content)
    with open("output.txt", "r") as f:
        for line in f:
            if not has_numbers(line) and not line.isspace():
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
                if i % 3 == 0:
                    csv_row.append(element)
                    patient.append(element)
                elif i % 3 == 1:
                    doc = element.split()[2]
                    csv_row.append(doc)
                    patient.append(doc)
                else:
                    csv_row.append(element)
                    patient.append(element)
                    writer.writerow(csv_row)
                    output_list_3.append(patient)
                    csv_row = []
                    patient = []
                    
    print(output_list_3)




def collect_files():
    print("Button 1 clicked")
    global filename
    filename = askopenfilename()

    print(filename)


# def button2_click():
#     print("Button 2 clicked")
#     # Add your function code here
#     doc_dict, unknown_doc_dict_1, unknown_doc_dict_2 = doc_dict_maker()
#     iter_over_files(doc_dict, unknown_doc_dict_1, unknown_doc_dict_2)


def button3_click():
    print("Button 3 clicked")
    # Add your function code here
    os.startfile("output.csv")

def open_bc(name):
    pya.moveTo(100, 450, duration=0.3)
    pya.click()
    pya.hotkey("ctrl", "o")
    pya.typewrite(name)
    pya.press('enter')
    pya.press('enter')
    
def button4_click():
    global output_list_3
    print("Button 4 clicked")
    # Add your function code here
    try:
        pat = output_list_3.pop()
    except IndexError:
        print("list empty")
    print(pat)
    # print(output_list_3)
    name_as_list = pat[0].split()
    name_for_bc = name_as_list[-1] + "," + name_as_list[1]
    print(name_for_bc)
    open_bc(name_for_bc)
    


# Create the main window
root = tk.Tk()
root.title("Rec1")
root.geometry("200x300")

# Create and pack the buttons vertically
button1 = tk.Button(root, text="Open File", command=collect_files, width=15, height=2)
button1.pack(pady=10)

button2 = tk.Button(
    root, text="Create datafile", command=extract, width=15, height=2
)
button2.pack(pady=10)

button3 = tk.Button(root, text="Open Datafile", command= button3_click, width=15, height=2)
button3.pack(pady=10)

button4 = tk.Button(root, text="Nothing yet", command=button4_click, width=15, height=2)
button4.pack(pady=10)



# Start the GUI event loop
root.attributes("-topmost", True)
root.mainloop()
