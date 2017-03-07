import time

from openpyxl import load_workbook
from requests import Session
from requests.auth import HTTPBasicAuth
from zeep import Client
from zeep.transports import Transport

# session for the webservice
session = Session()
user = 'jreyesc@pucp.edu.pe'
password = 'Jrc=021289'
session.auth = HTTPBasicAuth(user, password)
url = 'https://pucp-csm.symplicity.com/ws/student_api.php?wsdl'
client = Client(url, transport=Transport(session=session))

# open the excel file with pucp codes
wb = load_workbook('codigos.xlsx')
sheet = wb['Hoja1']
for row in sheet.rows:
  code = row[0].value
  if isinstance(code, int):
    result = client.service.getAsObjectBySchoolID(code)
    row[6].value = result['resume_book_flag']
    #if result['resume_book_flag'] == False:
    #  print(result)
    #time.sleep(1)
    # degree_level
    #print(result['degree_level'])
    #print(result['degree_level'])
    #print(result['student_profile__custom_field_93']) # alumno
    #print(result['student_profile__custom_field_94']) # egresado
    #print(result['student_profile__degree_level_according_syllab2'])
    #print(result['student_profile__academic_level_according_sylla'])

    # phone
    #print(str(result['phone']) + ',' + str(result['student__custom_field_12']) + ',' + str(result['student__custom_field_13']))
    #print(result['phone'])
    #print(result['student__custom_field_12'])
    #print(result['student__custom_field_13'])

    # email
    #print(str(result['email']) + ',' + str(result['student__custom_field_16']) + ',' + str(result['student__main_email']))
    #print(result['email']) # cato
    #print(result['student__custom_field_16'])
    #print(result['student__main_email'])

    # craest
    #print(result['student_profile__craest'])

    wb.save(filename = 'codigos.xlsx')

# result = client.service.getAsObjectBySchoolID('20092199')
# print(result)