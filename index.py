import sys
# excel reader
import openpyxl


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
inputfile_path = 'Input/aishe_10.xlsx'
result = read_xl_write_arr(inputfile_path)
print(result)

# for data_set in result:
#     print(data_set)