import pandas as pd
import re
import fitz
from collections import Counter

# Reading the Excel file
excel_file = 'C:\\Users\\recepce\\Documents\\NI\\Duplicita nebo chybejici\\stat.xlsx'
df = pd.read_excel(excel_file)

# Ensure that reservation numbers from Excel are strings and remove any trailing '.0'
excel_reservation_numbers = set(df['cislo Rez na BDC'].astype(str).str.replace('.0', '', regex=False).dropna())

# Inputting reservation numbers from AV
pdf_doc = fitz.open('av.pdf')
pdf_text = ""
for page in pdf_doc:
    pdf_text += page.get_text()

# Find all numbers in the PDF and convert them to strings
all_pdf_numbers = re.findall(r'\d+', pdf_text)
filtered_pdf_numbers = {num for num in all_pdf_numbers if len(num) > 8}  # Ensure set of strings

# Find duplicates in the PDF numbers
duplicate_reservation_numbers = [item for item, count in Counter(filtered_pdf_numbers).items() if count > 1]

# Compare sets to find numbers in Excel but not in PDF, and vice versa
missing_in_pdf = {num for num in excel_reservation_numbers - filtered_pdf_numbers if num.isdigit()}
missing_in_excel = {num for num in filtered_pdf_numbers - excel_reservation_numbers if num.isdigit()}

# Print the result
if duplicate_reservation_numbers:
    print("Duplicate reservation numbers found:")
    for number in duplicate_reservation_numbers:
        print(number)
else:
    print("No duplicate reservation numbers found.")

# Print missing and additional reservation numbers in one line
print(f"Rezervace chybí: {', '.join(missing_in_pdf) if missing_in_pdf else 'None'}")
print(f"Rezervace navíc: {', '.join(missing_in_excel) if missing_in_excel else 'None'}")
