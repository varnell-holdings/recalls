import datetime as dt
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

# we need 1.date 2. patient name  3. pat age 4. address 5. doctor 6. procedure


pat = ["Ms Mayumi TORGERSEN", "Stoita", "0406-627-052", "Colonoscopy"]
address1 = "7 Jones St"
address2 = "Smithfield, NSW, 2099"
today = dt.date.today()
today_str = today.strftime("%A, %d %B %Y")


def is_over_75(date_of_birth):
    """
    Check if a person is 75 years old or older based on their date of birth.

    Args:
        date_of_birth (str): Date of birth in format "dd/mm/yyyy"

    Returns:
        bool: True if aged 75 or over, False otherwise
    """
    # Parse the date of birth
    dob = dt.datetime.strptime(date_of_birth, "%d/%m/%Y")

    # Get today's date
    today = dt.datetime.now()

    # Calculate age
    age = today.year - dob.year

    # Adjust if birthday hasn't occurred this year yet
    if (today.month, today.day) < (dob.month, dob.day):
        age -= 1

    return age > 75


def make_html_body(pat, dob, address1, address2):
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

    path_to_template = "."
    loader = FileSystemLoader(path_to_template)
    env = Environment(loader=loader)
    template_name = "body_template.html"
    template = env.get_template(template_name)
    page = template.render(
        today_date=today_str,
        full_name=full_name,
        title=title,
        last_name=last_name,
        address1=address1,
        address2=address2,
        doctor=doctor,
        procedure=procedure,
        over_75=over_75,
    )

    with open("body.html", "wt") as f:
        f.write(page)


if __name__ == "__main__":
    pat = ["Ms Mayumi TORGERSEN", "Stoita", "0406-627-052", "Colonoscopy"]
    dob = "30/08/1900"
    address1 = "7 Jones St"
    address2 = "Smithfield, NSW, 2099"
    make_html_body(pat, dob, address1, address2)
