from googleapiclient.discovery import build
from google.oauth2 import service_account
import os
from create_xlsx import Create_xlsx
from tabel_statistics import Statistics

from config import tabel_folder

class App:
    

    def __init__(self, link, list, range, list2):
        self.link = link
        self.list = list
        self.range = range
        self.list2 = list2

        


    def start(self):
        print("Application starting...")
        print("Trying to read data from googlesheets...")

        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, 'credentials.json')

        credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=credentials).spreadsheets().values()



        #id of document what`s splited for his link`
        spreadsheet_id = self.link.split("/")[5]

        #range - concatenate from list and range by symbol "!"
        range = self.list+ "!" + self.range
        range2 = self.list2+ "!" + self.range
        #print(range2)


        result = service.get(spreadsheetId=spreadsheet_id,
                             range=range, valueRenderOption='UNFORMATTED_VALUE').execute()

        data_from_sheet = result.get('values', [])

        result2 = service.get(spreadsheetId=spreadsheet_id,
                             range=range2, valueRenderOption='UNFORMATTED_VALUE').execute()

        data_from_sheet2 = result2.get('values', [])
        print("Data successfully readed.")
        #print(data_from_sheet2[0][9])
        #print(data_from_sheet2)
        #print(data_from_sheet[3])
        klas = data_from_sheet[0][8]
        first_semstr = data_from_sheet[0][9]
        second_semstr = data_from_sheet2[0][9]
        year = data_from_sheet[0][15]
        f_teacher = data_from_sheet[0][30]
        s_teacher = data_from_sheet2[0][26]
        
        #getting list subjects and remove unnesesery items(first 2 and last 2)
        #subjects = data_from_sheet[3]
        # del subjects[0]
        # del subjects[0]
        # del subjects[len(subjects)-1]
        # del subjects[len(subjects)-1]

        #creating list with dictionary of students and grades for each other 
        students=[]
       
        #length = len(subjects)
        test = []
        each_student=1
        while each_student<=34:
            student = {}
            if data_from_sheet[each_student+3][1] != "":
                student["name"] = data_from_sheet[each_student+3][1]
                student["year"] = ""
                student["years"] = data_from_sheet[0][15]
                student["class"] = data_from_sheet[0][8]
                student["semester"] = data_from_sheet[0][9]

                student["subjects"] = {}
                student["second_semester"] = {}
                i=0

                for item in data_from_sheet[3]:
                    if item !="":
                        student["subjects"][item] = data_from_sheet[each_student+3][i]
                        try:
                            student["second_semester"][item] = data_from_sheet2[each_student+3][i]
                        except IndexError:
                            pass
                    i+=1
                
                students.append(student)
            each_student+=1
        
        # print(students[0]["name"])
        # print(students[0]["second_semester"])

        xlsx_adapter = Create_xlsx(output_folder=tabel_folder + "/")
        xlsx_adapter.create(students=students)
        stats = Statistics(document_id = spreadsheet_id, first_semester=self.list, second_semester = self.list2)
        try:
            stats.generate_statictics(klas=klas, f_semester=first_semstr, s_semester=second_semstr, year=year, f_teacher=f_teacher, s_teacher=s_teacher)
        except ZeroDivisionError:
            pass

       
        print("Application succsesfully completed =)")

        

        



        