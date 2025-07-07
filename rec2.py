import csv
import mammoth


def has_numbers(inputString):
    return any(char.isdigit() for char in inputString)


def docx_to_text_mammoth(docx_file):
    with open(docx_file, "rb") as docx_file:
        result = mammoth.extract_raw_text(docx_file)
        return result.value


output_list_1 = []
output_list_2 = []

text_content = docx_to_text_mammoth("bc1.docx")
with open("output.txt", "w", encoding="utf-8") as f:
    f.write(text_content)
with open("output.txt", "r") as f, open("output2.txt", "w") as f2:
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
        for i, element in enumerate(output_list_2):
            if i % 3 == 0:
                csv_row.append(element)
            elif i % 3 == 1:
                doc = element.split()[2]
                csv_row.append(doc)
            else:
                csv_row.append(element)
                writer.writerow(csv_row)
                csv_row = []


# print(output_list_2)
