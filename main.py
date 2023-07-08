import openpyxl
import sys
import requests
import json

filename = sys.argv[1]
g_key = sys.argv[2]
wb_obj = openpyxl.load_workbook(filename)

sheet_obj = wb_obj.active
max_row = sheet_obj.max_row

already_writted = []


def write_address_in_file(address, row):
    mycell = sheet_obj['F' + str(row)]
    mycell.value = address


def get_address_from_data(json_data, row):
    for data in json_data:
        types = data['types']

        if row in already_writted:
            continue

        if 'political' in types:
            write_address_in_file(data['formatted_address'], row)
            already_writted.append(row)


def get_info_from_gapi(latitud, longitud, row):
    api_result = requests.get(
        "https://maps.google.com/maps/api/geocode/json?latlng=" + latitud + "," + longitud + "&key=" + g_key,
        stream=True).raw.data
    json_object = json.loads(api_result)

    get_address_from_data(json_object["results"], row)


def read_excel():
    for i in range(2, max_row + 1):
        latitud = sheet_obj.cell(row=i, column=1)
        longitud = sheet_obj.cell(row=i, column=2)

        if longitud.value and latitud.value:
            latitud = str(latitud.value)
            longitud = str(longitud.value)

            get_info_from_gapi(latitud, longitud, i)

    wb_obj.save(filename)


read_excel()
