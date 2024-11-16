import sys
import os
import math
import pandas as pd
from fillpdf import fillpdfs        


def filter_data(data, form):
    data_dict = {}
    count = 0
    page_num = 1
    for index, row in data.iterrows():
        analyte = row[0]
        # kinda of annoying need all these if statements because they didn't name the 
        # form fields consistently throughout all 14 pages -_-
        # page 2 field forms
        if page_num == 2:
            if count == 0:
                print(f'starting page: {page_num}')
            # print(index)
            # print(f"'P{page_num}C1R{count+1}' : '{analyte}'")
            analyte_field = f'P{page_num}C1R{count+1}'
            data_dict.update({analyte_field : analyte})

            if count < 6:
                count += 1
            else:
                count = 0
                page_num += 1
        # page 6 field forms
        elif page_num == 6:
            if count == 0:
                print(f'starting page: {page_num}')
            # print(index)
            # print(f"'P{page_num}C1R1.0.{count}' : '{analyte}'")
            analyte_field = f'P{page_num}C1R1.0.{count}'
            data_dict.update({analyte_field : analyte})
            
            if count < 6:
                count += 1
            else:
                count = 0
                page_num += 1
        # page 7 field forms
        elif page_num == 7:
            if count == 0:
                print(f'starting page: {page_num}')
            # print(index)
            # print(f"'P{page_num}C1R1.0.0.{count}' : '{analyte}'")
            analyte_field = f'P{page_num}C1R1.0.0.{count}'
            data_dict.update({analyte_field : analyte})
            
            if count < 6:
                count += 1
            else:
                count = 0
                page_num += 1
        # page 8-9 field forms
        elif page_num >= 8 and page_num < 10:
            if count == 0:
                print(f'starting page: {page_num}')
            analyte_field = f'{count + 44 * (page_num - 8) + 1}'
            print(f"'{analyte_field}' : '{analyte}'")
            data_dict.update({analyte_field : analyte})
            if count < 6:
                count += 1
            else:
                count = 0
                page_num += 1
        # page 10 field forms
        elif page_num == 10:
            if count == 0:
                print(f'starting page: {page_num}')
            analyte_field = f'{count + 89}'
            print(f"'{analyte_field}' : '{analyte}'")
            data_dict.update({analyte_field : analyte})
            if count < 6:
                count += 1
            else:
                count = 0
                page_num += 1
        # page 11 field forms
        elif page_num == 11:
            if count == 0:
                print(f'starting page: {page_num}')
            analyte_field = f'{count + 134}'
            print(f"'{analyte_field}' : '{analyte}'")
            data_dict.update({analyte_field : analyte})
            if count < 6:
                count += 1
            else:
                count = 0
                page_num += 1
        # page 12 field forms
        elif page_num == 12:
            if count == 0:
                print(f'starting page: {page_num}')
            analyte_field = f'{count + 178}'
            print(f"'{analyte_field}' : '{analyte}'")
            data_dict.update({analyte_field : analyte})
            if count < 6:
                count += 1
            else:
                count = 0
                page_num += 1
        # page 14 field forms
        elif page_num == 14:
            if count == 0:
                print(f'starting page: {page_num}')
            analyte_field = f'{count + 284}'
            print(f"'{analyte_field}' : '{analyte}'")
            data_dict.update({analyte_field : analyte})
            if count < 6:
                count += 1
            else:
                count = 0
                page_num += 1
        # pages 1, 3, 4, 5 field forms
        else:
            if count == 0:
                print(f'starting page: {page_num}')
            # print(index)
            # print(f"'P{page_num}C1R1.{count}' : '{analyte}'")
            analyte_field = f'P{page_num}C1R1.{count}'
            even_result_field = f'P{page_num}C2R1.{count*2}'
            odd_result_field = f'P{page_num}C2R1.{count*2+1}'
            print(even_result_field,odd_result_field)
            data_dict.update({analyte_field : analyte})
            
            if count < 6:
                count += 1
            else:
                count = 0
                page_num += 1

    #print(data_dict)

    #fillpdfs.write_fillable_pdf(form, 'new.pdf', data_dict)


if __name__ == '__main__':
    # filename of SPCC field form
    form = "2023 DMR Form template.pdf"
    #fields = fillpdfs.print_form_fields(form)
    #print(fields)

    # filename of stormwater data
    spreadsheet = "14 CNRH DATA BLDG19.xlsx"
    toxic_df = pd.read_excel(spreadsheet, sheet_name = "Toxic Parameters", header = 2, skipfooter = 5)
    field_df = pd.read_excel(spreadsheet, sheet_name = "Field Data", header = 1, skipfooter = 2)
    
    filter_data(toxic_df, form)
