import csv

import mammoth

output_list_1 = []
output_list_2 = []
output_list_3 = []
output_list_3 = []


def has_alpha(inputString):
    return any(char.isalpha() for char in inputString)


def has_numbers(inputString):
    return any(char.isdigit() for char in inputString)


def docx_to_text_mammoth(docx_file):
    with open(docx_file, "rb") as docx_file:
        result = mammoth.extract_raw_text(docx_file)
        return result.value


def extract():
    global print_length
    global output_list_1
    global output_list_2
    global output_list_3
    global output_list_4

    text_content = docx_to_text_mammoth(full_path)
    for line in text_content.splitlines():
        if line and (
            not (has_numbers(line) and has_alpha(line))
            and not (has_numbers(line) and "/" in line)
        ):
            output_list_1.append(line.strip())

    # remove first 5 lines
    for i, element in enumerate(output_list_1):
        if i < 5:
            continue
        else:
            output_list_2.append(element)

    # replace absent phone number with 0000 to keep order correct
    for i, element in enumerate(output_list_2):
        if i % 4 == 0:
            output_list_3.append(element)
        elif i % 4 == 1:
            output_list_3.append(element)
        elif i % 4 == 2:
            if not has_numbers(element):
                output_list_3.append("0000")
                output_list_3.append(element)
                i += 1
            else:
                output_list_3.append(element)
        else:
            output_list_3.append(element)
    # group the list into patients lists with 4 elements each
    patient = []
    for i, element in enumerate(output_list_3):
        if i % 4 == 0:
            patient.append(element)
        elif i % 4 == 1:
            doc = element.split()[2]
            patient.append(doc)
        elif i % 4 == 2:
            patient.append(element)
        else:
            patient.append(element)
            output_list_4.append(patient)
            patient = []

    print(output_list_4)
    # print_length = len(output_list_3)
    # num_to_do.set(str(print_length))
    # n.set(f"{print_length} patients to do.")


if __name__ == "__main__":
    full_path = "./bc_reports/bc1_no_phone.docx"
    extract()
