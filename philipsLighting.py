import requests   # Use to pull data from server
import pyodbc     # Use to connect to a database
import xlrd       # Get excel Info
import xlwt       # Write to excel file
import json       # 

# Dictionary to orgranize  data
EnvisLightingData ={
    'areaNumber': [],
    'areaName' : [],
    'luminaireLevels' : [],
    'getCurrentStatus' : []
}

# Access server
username = '[REDACT]'
password = '[REDACT]'

# Open Workbook
wb = xlrd.open_workbook('WattCenterAreaNumbers.xlsx')

# Excel Column Info
areaNumberColumn = 0
areaNameColumn = 1
areaLuminaireColumn = 2

# Read each sheet in excel file
for sheet in wb.sheets():
    for row in range(1, sheet.nrows):
        # Area number was read as float, convert it to an int
        AreaNumber_f = sheet.cell_value(row,areaNumberColumn)
        AreaNuber_int = int(AreaNumber_f)

        # Store the Area Number, Area Name, and its data
        EnvisLightingData['areaNumber'].append(AreaNuber_int)
        EnvisLightingData['areaName'].append(sheet.cell_value(row,areaNameColumn))

        # gain access to page using credentials for each Area Number
        url = 'https://envision.clemson.edu/services/rest/control_restservice/getCurrentStatus/' + str(EnvisLightingData['areaNumber'][row - 1])
        
        # Get data from server
        r = requests.get(url=url, auth=(username,password), verify=False)
        data = json.loads(r.text)
        dataString = r.text

        # Store data from server
        if dataString.find("luminaireLevels") > 0:
            EnvisLightingData['luminaireLevels'].append(data["luminaireLevels"][0]["luminaireID"])
        else:
            EnvisLightingData['luminaireLevels'].append("No Data")


        sheet.write(row,areaLuminaireColumn,EnvisLightingData['luminaireLevels'][row - 1])


# Last part is to store in database
serverConnect = pyodbc.connect('DRIVER= {ODBC Driver 13 for SQL Server};' 'SERVER = [REDACT]];' 'DATABASE = [REDACT];' 'UID = [REDACT];' 'PWD = [REDACT]')
cursor = serverConnect.cursor()

#SQL Commit
cursor.execute("INSERT INTO (TABLE_NAME: CASE SERSATIVE) (COLUMN NAME1 - SEPERATED BY COMMAS) VALUES(?, ?, ?, ?)", areaLevel, ID, name, type)
cursor.commit()
cursor.close()