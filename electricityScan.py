import pdfplumber
import pandas as pd
import argparse
import os
import openpyxl


def parse_pdf(list_file_name, root):
    # One record per service location, so amount/meter/usage can never get out of step
    records = []
    for file_name in list_file_name:
        if not file_name.lower().endswith('.pdf'):
            continue
        path = os.path.join(root, file_name)
        record = None
        with pdfplumber.open(path) as pdf:
            for p0 in pdf.pages:
                text = p0.extract_text(keep_blank_chars=True) or ''
                save = -1
                for num, line in enumerate(text.splitlines()):
                    list_line = line.split()
                    if not list_line:
                        continue

                    if 'Status:' in list_line:
                        # A new service location starts here; store the previous one
                        if record is not None:
                            records.append(record)
                        slice = list_line.index('Status:')
                        address = " ".join(list_line[:slice])
                        record = {'Amount': None, 'Address': address, 'Meter': 'NA', 'Usage': 'NA'}
                    if record is None:
                        continue

                    if 'Total' in list_line and 'Current' in list_line and 'Activity' in list_line:
                        price = list_line[-1]
                        record['Amount'] = float(price.replace('$', '').replace(',', ''))

                    if list_line[:5] == ['Meter', '#', 'Rate', 'Cycle', 'Days']:
                        save = num + 2
                    if num == save and len(list_line[0]) == 5 and list_line[0].isdigit():
                        record['Meter'] = int(list_line[0])

                    if list_line[:2] == ['Current', 'Month'] and len(list_line) > 3:
                        record['Usage'] = int(list_line[3].replace(',', ''))
        if record is not None:
            records.append(record)
    return pd.DataFrame(records, columns=['Amount', 'Address', 'Meter', 'Usage'])

def update_excel(workbook, new_values, month):
    sheet = workbook["Electricity"]

    # Look up both headers exactly, ignoring extra or non-breaking spaces
    month = " ".join(month.split())
    headers = {" ".join(str(cell.value).split()): count
               for count, cell in enumerate(sheet[4]) if cell.value is not None}
    amount_header = f"{month} Amount"
    usage_header = f"{month} Usage"
    if amount_header not in headers or usage_header not in headers:
        raise ValueError(f"Could not find '{amount_header}' and '{usage_header}' in row 4 of the Electricity sheet")
    index_amount = headers[amount_header]
    index_usage = headers[usage_header]

    unmatched = []
    for _, j in new_values.iterrows():
        found = False
        for row in sheet.iter_rows(min_row=5):
            if row[4].value == j['Meter']:
                row[index_amount].value = j['Amount']
                row[index_usage].value = j['Usage']
                found = True
        if not found:
            unmatched.append(j)

    if unmatched:
        print(f"{len(unmatched)} bill(s) had no matching meter in column E:")
        for j in unmatched:
            print(f"  {j['Address']}: meter {j['Meter']}, ${j['Amount']}, {j['Usage']} kWh")

    workbook.save('2025-26 Electricity Billing.xlsx')

        


if __name__ == "__main__":
    #parse arguments
    parser = argparse.ArgumentParser("Parse info in pdf files")
    parser.add_argument("-f", help="folder name") 
    parser.add_argument("-e", help='existing excel file')
    parser.add_argument("-m", help="month") 
    args = parser.parse_args()

    files = os.listdir(args.f)
    dic_data = parse_pdf(files, args.f)

    new_values = pd.DataFrame(dic_data)
    workbook = openpyxl.load_workbook(args.e)
    update_excel(workbook, new_values, args.m)
    print("Successfuly updated the values")
