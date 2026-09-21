import pdfplumber
import pandas as pd
import calendar
import argparse
import os
import openpyxl


data = {
    'Amount': [],
    'Address': [],
    'Meter': [],
    'Usage': [],
    # 'Month': []
}

def parse_pdf(list_file_name, root):
    for file_name in list_file_name:
        path = root + "/" + file_name
        with pdfplumber.open(path) as pdf:
            size_doc = len(pdf.pages)
            for page in range(size_doc):
                p0 = pdf.pages[page]
                text = p0.extract_text(keep_blank_chars=True)
                output = open("output.txt",'w', encoding='utf-8')
                output.write(text)
                output = open('output.txt', 'r', encoding='utf-8')
                save = -1
                for num, line in enumerate(output.readlines()):
                    list_line = str(line).strip().split(' ')
                    # print(list_line)
                    if 'Total' in list_line and 'Current' in list_line and 'Activity' in list_line:
                        price = list_line[-1]
                        data['Amount'].append(price[1:])
                        # print(f'Price: {price} \n')

                    if 'Status:' in list_line:
                        slice = list_line.index('Status:')
                        address = " ".join(list_line[:slice])
                        data['Address'].append(address)

                        # print(f'Address: {address}')
                    if list_line == ['Meter', '#', 'Rate', 'Cycle', 'Days'] or list_line == ['Meter', '#', 'Rate', 'Cycle', 'Days', 'Demand']:
                        save = num + 2
                    if num == save:
                        if len(list_line[0]) == 5:
                            meter = list_line[0]
                            data['Meter'].append(int(meter))
                            # print(f'Meter: {meter}')
                        else: 
                            data['Meter'].append('NA')
                            data['Usage'].append('NA')
                            # print('No Meter associated')

                    if 'Current' in list_line and 'Month' in list_line and 'Previous' not in list_line:
                        consumption = list_line[3]
                        if consumption==None:
                            data['Usage'].append('NA')
                            
                        else:
                            data['Usage'].append(int(consumption))
                        # print(f'Consumption: {consumption} kW')

                    # if 'Billing' in list_line and 'Date:' in list_line:
                    #     month = calendar.month_abbr[int(list_line[-1][:2])-1]
            
            # tot = len(data['Amount'])
            # for i in range(tot):
            #     data['Month'].append(month + ".") 
    return data

def update_excel(workbook, new_values, month):
    #specific_month = new_values["Month"][1]
    specific_month = month
    sheet = workbook["Electricity"]

    for count, cell in enumerate(sheet[4]):
        if cell.value is not None:
            value = str(cell.value.split(" ")[0])
            if value == specific_month:
                index_amount = count - 1
                index_usage = count 
    
    for i, j in new_values.iterrows():
        for row in sheet.iter_rows(min_row=4):

            if row[4].value == j.iloc[2]:
                row[index_amount].value = j.iloc[0]
                row[index_usage].value = j.iloc[3]

    workbook.save('2025-26 Utility Billing.xlsx')

        


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
    
