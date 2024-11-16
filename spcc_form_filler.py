import sys
import os
import math
import pandas as pd
from fillpdf import fillpdfs

"""
 This function fills out the SPCC form with JBPHH tank data
 tank_id, building, capacity, contents, location, poc, and poc_phone are all linked to the respect tank list column names in the google sheet
"""
def get_nav_tanks(tank_list, form):
    tanks = pd.read_csv(tank_list)
    form = form
    
    for index, row in tanks.iterrows():
        tank_id = row['Tank/Facility No.']
        building = row['Bldg No.']
        capacity = row['Volume Stored (gallons)']
        contents = row['Product Stored']
        location = row['Location']
        poc = row['POC']
        poc_phone = row['PHONE']
        if math.isnan(row['Reference (N-)']):
            ref_num = 'XX'
        else:
            ref_num = round(row['Reference (N-)'])

        data_dict = {
            'ACTIVITY' : 'NAVFAC HI',
            'TANK / FACILITY ID' : tank_id,
            'BUILDING' : building,
            'CAPACITY' : capacity,
            'CONTENTS' : contents,
            'LOCATION' : location,
            'POC' : poc,
            'POC PHONE' : poc_phone,
        }
        # print(data_dict)
        
        if not os.path.exists(os.path.join('NAVFAC', f'N-{ref_num}_{tank_id}.pdf')):
            filename = os.path.join('NAVFAC',f'N-{ref_num}_{tank_id}.pdf')
            print(f'Working on Tank {tank_id}')
            # print(filename)

            fillpdfs.write_fillable_pdf(form, filename, data_dict)

"""
This function fills out the SPCC form with JBPHH tank data
tank_id, building, capacity, contents, location, poc, and poc_phone are all linked to the respect tank list column names in the google sheet
"""
def get_jbphh_tanks(tank_list, form):
    tanks = pd.read_csv(tank_list)
    form = form
    
    for index, row in tanks.iterrows():
        tank_id = row['Tank/Facility No.']
        building = row['Bldg No.']
        capacity = row['Volume Stored (gallons)']
        contents = row['Product Stored']
        location = row['Location']
        poc = row['POC']
        poc_phone = row['PHONE']
        # print(type(row['tankHelper']))
        if math.isnan(row['tankHelper']):
            ref_num = 'XX'
        else:
            ref_num = round(row['tankHelper'])

        data_dict = {
            'ACTIVITY' : 'JBPHH',
            'TANK / FACILITY ID' : tank_id,
            'BUILDING' : building,
            'CAPACITY' : capacity,
            'CONTENTS' : contents,
            'LOCATION' : location,
            'POC' : poc,
            'POC PHONE' : poc_phone,
        }
        # print(data_dict)
        
        if not os.path.exists(os.path.join('JBPHH', f'J-{ref_num}_{tank_id}.pdf')):
            filename = os.path.join('JBPHH', f'J-{ref_num}_{tank_id}.pdf')
            print(f'Working on Tank {tank_id}')
            print(filename)

            fillpdfs.write_fillable_pdf(form, filename, data_dict)
        

if __name__ == '__main__':
    # filename of SPCC field form
    form = "2024 NEW Blank SPCC - Inspection Form Fillable.pdf"
    
    # Make sure to change google sheet link from: https://docs.google.com/spreadsheets/d/{id}/edit?gid={id} -> https://docs.google.com/spreadsheets/d/{id}/export?format=csv&gid={id}
    # link to NAVFAC tank list google sheet - link needs to be public
    navfac = "https://docs.google.com/spreadsheets/d/1ygqmqV5uvLjqacryKz0rFOh2dNqFgMP-iamx2aSw9VA/export?format=csv&gid=1654911570"
    # link to JBPHH tank list google sheet - link needs to be public
    jbphh = "https://docs.google.com/spreadsheets/d/1ygqmqV5uvLjqacryKz0rFOh2dNqFgMP-iamx2aSw9VA/export?format=csv&gid=657100141"

    get_nav_tanks(navfac, form)
    get_jbphh_tanks(jbphh, form)
