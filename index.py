import sys
# excel reader
import openpyxl
# url encoder
from urllib.parse import quote
# request module
import requests
# os module
import os
# import csv module
import csv

# debuger module
# import json

# ++++++++++++ variable init ++++++++++#
baseUrl = 'https://maps.googleapis.com/maps/api/geocode/json?address='
API_key = '&key='


# check the argument is provide (arg must be the input filename)

if len(sys.argv) < 2 or not sys.argv[1]:
    print('Requested params is missing')
    sys.exit()
arg_param = sys.argv[1]

print(arg_param)

# defining a funtion for read excel file and write to an array
def read_xl_write_arr(fpath):
    #load excel file
    exlFile = openpyxl.load_workbook(fpath)
    sheet = exlFile.worksheets[0]
    #get first row for key of each array
    header = [cell.value for cell in sheet[1]]

    # empty array init
    loopArr = []

    # itrate through the row from second for data
    for row in sheet.iter_rows(min_row=2, values_only=True):
        # create a assciate array using key for header and the data
        itemArr = dict(zip(header, row))
        # add to the main loopArray
        loopArr.append(itemArr)

    # close the open file
    exlFile.close()

    return loopArr

# call the function to get array
# inputfile_path = 'Input/' + arg_param + '.xlsx'
inputfile_path = f"Input/{arg_param}.xlsx"

result = read_xl_write_arr(inputfile_path)
# print(json.dumps(result, indent=2))
# exit()


# api call function

def geoc_api_call(url, baseUrl, API_key):
    call_url = f"{baseUrl}{url}{API_key}"

    response = requests.get(call_url)
    if(response.status_code == 200):
        res_data = response.json()
        # print(res_data['results'][0]['geometry']['location']['lat'])
        eliminated_array = {
            'latitude': res_data['results'][0]['geometry']['location']['lat'] if res_data['results'] and res_data['results'][0] and 'geometry' in res_data['results'][0] and 'location' in res_data['results'][0]['geometry'] else None ,
            'longitude': res_data['results'][0]['geometry']['location']['lng'] if res_data['results'] and res_data['results'][0] and 'geometry' in res_data['results'][0] and 'location' in res_data['results'][0]['geometry'] else None
        }

        return eliminated_array

    return None 

# Creating target array

api_call_array = [] # init api_call_array
final_array = [] # init final_array
for item in result:

    id = item['id']
    name = item['name']
    pin_code = item['pin_code']
    city = item['city']
    statename = item['statename']
    
    full_address = item['name']+ ' ' + item['city'] + ' ' + item['address_line1'] # Address concating 
    # print( quote(full_address) )
    url = quote(full_address) # encode the url

    latlong_info = geoc_api_call(url, baseUrl, API_key)
    # if latlong_info is not None:
    #     item.update(latlong_info)
    item.update(latlong_info)
    final_array.append(item)

    
# print(json.dumps(final_array, indent=2))
# exit()

# generate csv output
    
print('generating output')


csv_header = list(final_array[0].keys())
outputPath = 'csv/output.csv'

# Define csvGenerate function

def generate_Csv(outputPath, csv_header, body_data):
    if not os.path.isdir('csv'):
        os.mkdir('csv')
    
    with open(outputPath, 'w', newline='', encoding='utf-8-sig') as file:
        csv_writer = csv.writer(file, delimiter=';')

        # write header
        csv_writer.writerow(csv_header)

        # write body data
        for b_item in body_data:
            csv_body = [
                str(b_item['id']),
                b_item['name'],
                b_item['pin_code'],
                b_item['city'],
                b_item['statename'],
                str(b_item['latitude']),
                str(b_item['longitude'])

            ]

            csv_writer.writerow(csv_body)



generate_Csv(outputPath, csv_header, final_array) # Call ccsvGenerate function

print('output generated')

