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

# defining a funtion for read excel file and write to an array
def read_xl_write_arr(fpath):
    
    exlFile = openpyxl.load_workbook(fpath) #load excel file
    sheet = exlFile.worksheets[0]
    header = [cell.value for cell in sheet[1]]  #get first row for key of each array
    
    loopArr = [] # empty array init

    # itrate through the row from second for data
    for row in sheet.iter_rows(min_row=2, values_only=True):
        # create a assciate array using key for header and the data
        itemArr = dict(zip(header, row))
        # add to the main loopArray
        loopArr.append(itemArr)

    exlFile.close() # close the open file

    return loopArr


# inputfile_path = 'Input/' + arg_param + '.xlsx'
inputfile_path = f"Input/{arg_param}.xlsx"

result = read_xl_write_arr(inputfile_path) # call the function to get input array
# print(json.dumps(result, indent=2))
# exit()


# api call function
def geoc_api_call(url, baseUrl, API_key):
    call_url = f"{baseUrl}{url}{API_key}"

    response = requests.get(call_url)
    if(response.status_code == 200):
        res_data = response.json()
        eliminated_array = {
            'latitude': res_data['results'][0]['geometry']['location']['lat'] if res_data['results'] and res_data['results'][0] and 'geometry' in res_data['results'][0] and 'location' in res_data['results'][0]['geometry'] else None ,
            'longitude': res_data['results'][0]['geometry']['location']['lng'] if res_data['results'] and res_data['results'][0] and 'geometry' in res_data['results'][0] and 'location' in res_data['results'][0]['geometry'] else None
        }

        return eliminated_array

    return None 

# Creating target array. this array is the final array [id, name, pin_code, statename, lat, long]
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

    latlong_info = geoc_api_call(url, baseUrl, API_key) # call api function 
    # if latlong_info is not None:
    #     item.update(latlong_info)
    item.update(latlong_info) # add to final array
    final_array.append(item) # append each data set

    
# print(json.dumps(final_array, indent=2))
# exit()

# generate csv output
    
print('generating output')


csv_header = list(final_array[0].keys())
outputPath = 'csv/output.csv'
outputPath = f"csv/{arg_param}.csv"

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

