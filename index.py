import sys
# excel reader
import openpyxl
# url encoder
from urllib.parse import quote
# request module
import requests

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
inputfile_path = 'Input/onedata.xlsx'
result = read_xl_write_arr(inputfile_path)
# print(result[0])
# exit
# print(result[0]["id"])
# for data_set in result:
#     print(data_set)


# api call function

def geoc_api_call(url, baseUrl, API_key):
    call_url = f"{baseUrl}{url}{API_key}"

    response = requests.get(call_url)
    if(response.status_code == 200):
        res_data = response.json()
        # print(res_data['results'][0]['geometry']['location']['lat'])
        eliminated_array = {
            'lat': res_data['results'][0]['geometry']['location']['lat'] if res_data['results'] and res_data['results'][0] and 'geometry' in res_data['results'][0] and 'location' in res_data['results'][0]['geometry'] else None ,
            'long': res_data['results'][0]['geometry']['location']['lng'] if res_data['results'] and res_data['results'][0] and 'geometry' in res_data['results'][0] and 'location' in res_data['results'][0]['geometry'] else None
        }

        return eliminated_array

    return None 

# Creating target array
i=0

api_call_array = [] # init api_call_array
final_array = [] # init final_array
for item in result:
    
    full_address = item['name']+ ' ' + item['city'] + ' ' + item['address_line1'] # Address concating 
    # print( quote(full_address) )
    url = quote(full_address) # encode the url
    api_call_array.append(geoc_api_call(url, baseUrl, API_key))

    for r in api_call_array:
        if( isinstance(r, dict) ):
            final_array.append({
                'id': item['id'],
                'name': item['name'],
                'pin_code': item['pin_code'],
                'city': item['city'],
                'statename': item['statename'],
                'latitude': r.get('lat', None),
                'longitude': r.get('long', None)
            })
            print('true')
        # print(type(r))


    i +=1

#print(api_call_array)
    
print(final_array)







