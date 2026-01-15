"""
Fix date formats in recalls_csv.csv.
Converts DD/MM/YYYY and YYYY-MM-DD to DD-MM-YYYY format.
Also strips leading spaces from column headers.
Outputs to fixed_recalls_csv.csv.
"""

import csv
from dateutil.parser import parse


def fix_date(date_str):
    """Convert date string to DD-MM-YYYY format."""
    if not date_str.strip():
        return ""

    try:
        # Check if it's ISO format (YYYY-MM-DD)
        if len(date_str) == 10 and date_str[4] == '-':
            d = parse(date_str).date()
        else:
            # Assume DD/MM/YYYY or similar day-first format
            d = parse(date_str, dayfirst=True).date()

        return d.strftime("%d-%m-%Y")
    except Exception:
        print(f"  Warning: Could not parse date '{date_str}'")
        return date_str


def fix_recalls_csv(input_path, output_path):
    """Read recalls CSV, fix dates and headers, write to new file."""
    rows = []

    with open(input_path, 'r') as f:
        reader = csv.DictReader(f)
        # Strip spaces from fieldnames
        fieldnames = [name.strip() for name in reader.fieldnames]

        for row in reader:
            # Create new row with stripped keys
            new_row = {k.strip(): v for k, v in row.items()}

            # Fix date columns
            for date_col in ['first', 'second', 'third']:
                if date_col in new_row and new_row[date_col]:
                    original = new_row[date_col]
                    fixed = fix_date(original)
                    if original != fixed:
                        print(f"  Fixed: '{original}' -> '{fixed}'")
                    new_row[date_col] = fixed

            rows.append(new_row)

    # Write to output file
    with open(output_path, 'w', newline='') as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)

    print(f"Wrote {len(rows)} rows to {output_path}")


if __name__ == "__main__":
    input_path = "recalls_csv.csv"
    output_path = "fixed_recalls_csv.csv"

    print(f"Fixing dates in {input_path}...")
    fix_recalls_csv(input_path, output_path)
    print("Done!")
