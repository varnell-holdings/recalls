import datetime
from jinja2 import Environment, FileSystemLoader

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

pat = ["Ms Mayumi TORGERSEN", "Stoita", "0406-627-052", "Colonoscopy"]
street = "7 Jones St"
suburb_state = "Smithfield, NSW, 2099"


def make_html_body(street, suburb_state):
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

    path_to_template = "D:\\JOHN TILLET\\source\\active\\recalls"
    loader = FileSystemLoader(path_to_template)
    env = Environment(loader=loader)
    template_name = "body_1_template.html"
    template = env.get_template(template_name)
    page = template.render(
        today_date=today_str,
        full_name=full_name,
    )

    with open("D:\\JOHN TILLET\\source\\active\\recalls\\recall_1.html", "wt") as f:
        f.write(page)


if __name__ == "__main__":
    street = "7 Jones St"
    suburb_state = "Smithfield, NSW, 2099"
    make_html_body(street, suburb_state)
