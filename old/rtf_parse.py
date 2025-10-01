from striprtf.striprtf import rtf_to_text


rtf_file_path = "./bc_reports/recall_rtf.rtf"


def rtf_to_txt(rtf_file_path):
    text2 = ""
    output_list = []
    with open(rtf_file_path, "r") as rtf_file:
        rtf_content = rtf_file.read()

    text = rtf_to_text(rtf_content)

    for i, line in enumerate(text.split("\n")):
        if i == 0:
            continue
        elif "|||||" in line or not line:
            continue
        else:
            text2 += line
            text2 += "\n"

    for line in text2.split("\n"):
        local_list = []
        for i, field in enumerate(line.split("|")):
            if i == 0:
                local_list.append(field)
            elif i == 1:
                local_list.append(field.split()[2])
            elif i == 2:
                local_list.append(field)
            elif i == 3:
                local_list.append(field)
        if local_list != [""]:
            output_list.append(local_list)

    print(output_list)
