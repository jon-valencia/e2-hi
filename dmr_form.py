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
        #print(row)
        analyte = row[0]
        if row['Result'] == 'U':
            result = 'ND'
        else:
            result = row['Result']    
        limit = row[2]
        # kinda of annoying need all these if statements because they didn't name the 
        # form fields consistently throughout all 14 pages -_-
        # page 2 field forms
        if page_num == 2:
            if count == 0:
                print(f'starting page: {page_num}')
                result_field = f'P{page_num}C2R1'
            else:
                result_field = f'P{page_num}C2R2.0.{count * 2 - 1}'
            
            
            analyte_field = f'P{page_num}C1R{count + 1}'
            limit_field = f'P2C2R2.0.{count * 2}'

            data_dict.update({analyte_field : analyte})
            data_dict.update({limit_field : limit})
            data_dict.update({result_field : result})

            print(f'{analyte_field} : {analyte}')
            print(f'{result_field} : {result}')
            print(f'{limit_field} : {limit}')
            
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
            analyte_field = f'P6C1R1.0.{count}'
            result_field = f'P6C2R1.0.{count * 2}'
            limit_field = f'P6C2R1.0.{count * 2 + 1}'
            
            data_dict.update({analyte_field : analyte})
            data_dict.update({limit_field : limit})
            data_dict.update({result_field : result})

            print(f'{analyte_field} : {analyte}')
            print(f'{result_field} : {result}')
            print(f'{limit_field} : {limit}')
            
            if count < 6:
                count += 1
            else:
                count = 0
                page_num += 1
        # page 7 field forms
        elif page_num == 7:
            if count == 0:
                print(f'starting page: {page_num}')
            
            analyte_field = f'P{page_num}C1R1.0.0.{count}'
            result_field = f'P{page_num}C2R1.0.0.{count * 2}'
            limit_field = f'P{page_num}C2R1.0.0.{count * 2 + 1}'
            
            data_dict.update({analyte_field : analyte})
            data_dict.update({limit_field : limit})
            data_dict.update({result_field : result})

            print(f'{analyte_field} : {analyte}')
            print(f'{result_field} : {result}')
            print(f'{limit_field} : {limit}')
            
            if count < 6:
                count += 1
            else:
                count = 0
                page_num += 1
        # page 8-9 field forms
        elif page_num >= 8 and page_num < 10:
            if count == 0:
                print(f'starting page: {page_num}')
            
            analyte_field = f'{44 * (page_num - 8) + 1 + count}'
            result_field = f'{44 * (page_num - 8) + 8 + count * 2}'
            limit_field = f'{44 * (page_num - 8) + 9 + count * 2}'
            
            data_dict.update({analyte_field : analyte})
            data_dict.update({limit_field : limit})
            data_dict.update({result_field : result})

            print(f'{analyte_field} : {analyte}')
            print(f'{result_field} : {result}')
            print(f'{limit_field} : {limit}')
            
            if count < 6:
                count += 1
            else:
                count = 0
                page_num += 1
        # page 10 field forms
        elif page_num == 10:
            if count == 0:
                print(f'starting page: {page_num}')
            
            analyte_field = f'{89 + count}'
            result_field = f'{96 + count * 2}'
            limit_field = f'{96  + count * 2 + 1}'
            
            data_dict.update({analyte_field : analyte})
            data_dict.update({limit_field : limit})
            data_dict.update({result_field : result})

            print(f'{analyte_field} : {analyte}')
            print(f'{result_field} : {result}')
            print(f'{limit_field} : {limit}')
            
            if count < 6:
                count += 1
            else:
                count = 0
                page_num += 1
        # page 11 field forms
        elif page_num == 11:
            if count == 0:
                print(f'starting page: {page_num}')

            analyte_field = f'{134 + count}'
            result_field = f'{141 + count * 2}'
            limit_field = f'{141 + count * 2 + 1}'
            
            data_dict.update({analyte_field : analyte})
            data_dict.update({limit_field : limit})
            data_dict.update({result_field : result})

            print(f'{analyte_field} : {analyte}')
            print(f'{result_field} : {result}')
            print(f'{limit_field} : {limit}')
            
            if count < 6:
                count += 1
            else:
                count = 0
                page_num += 1
        # page 12 field forms
        elif page_num == 12:
            if count == 0:
                print(f'starting page: {page_num}')
            
            analyte_field = f'{178 + count}'
            result_field = f'{185 + count * 2}'
            limit_field = f'{185 + count * 2 + 1}'
            
            data_dict.update({analyte_field : analyte})
            data_dict.update({limit_field : limit})
            data_dict.update({result_field : result})

            print(f'{analyte_field} : {analyte}')
            print(f'{result_field} : {result}')
            print(f'{limit_field} : {limit}')
            
            if count < 6:
                count += 1
            else:
                count = 0
                page_num += 1

        # page 13 field forms
        elif page_num == 13:
            if count == 0:
                print(f'starting page: {page_num}')
            # extra logic to account for going from 233 to 245 because chronological
            # numbering apparently doesnt exist anymore
            if count == 5:
                analyte_field = f'{245}'
            elif count == 6:
                analyte_field = f'{246}'
            else:
                analyte_field = f'{229 + count}'
            result_field = f'{247 + count * 2}'
            limit_field = f'{247 + count * 2 + 1}'
            
            data_dict.update({analyte_field : analyte})
            data_dict.update({limit_field : limit})
            data_dict.update({result_field : result})

            print(f'{analyte_field} : {analyte}')
            print(f'{result_field} : {result}')
            print(f'{limit_field} : {limit}')
            
            if count < 6:
                count += 1
            else:
                count = 0
                page_num += 1

        # page 14 field forms
        elif page_num == 14:
            if count == 0:
                print(f'starting page: {page_num}')
            analyte_field = f'{284 + count}'
            result_field = f'{291 + count * 2}'
            limit_field = f'{291 + count * 2 + 1}'
            
            data_dict.update({analyte_field : analyte})
            data_dict.update({limit_field : limit})
            data_dict.update({result_field : result})

            print(f'{analyte_field} : {analyte}')
            print(f'{result_field} : {result}')
            print(f'{limit_field} : {limit}')

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
            result_field = f'P{page_num}C2R1.{count*2}'
            limit_field = f'P{page_num}C2R1.{count*2+1}'
            
            data_dict.update({analyte_field : analyte})
            data_dict.update({limit_field : limit})
            data_dict.update({result_field : result})

            print(f'{analyte_field} : {analyte}')
            print(f'{result_field} : {result}')
            print(f'{limit_field} : {limit}')
            
            if count < 6:
                count += 1
            else:
                count = 0
                page_num += 1

    # print(data_dict['P2C2R2.0.0'])
    fillpdfs.print_form_fields(form, page_number=2)
    # fillpdfs.print_form_fields(form, page_number=6)
    fillpdfs.write_fillable_pdf(form, 'new.pdf', data_dict)


if __name__ == '__main__':
    # filename of SPCC field form
    form = "2023 DMR Form template.pdf"
    #fields = fillpdfs.print_form_fields("new.pdf", page_number=1)
    # print(fields)

    # filename of stormwater data
    spreadsheet = "14 CNRH DATA BLDG19.xlsx"
    toxic_df = pd.read_excel(spreadsheet, sheet_name = "Toxic Parameters", header = 2, skipfooter = 5)
    field_df = pd.read_excel(spreadsheet, sheet_name = "Field Data", header = 1, skipfooter = 2)
    
    filter_data(toxic_df, form)
