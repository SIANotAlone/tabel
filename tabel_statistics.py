
import pandas as pd

import os
from googleapiclient.discovery import build
from google.oauth2 import service_account


class Statistics:
    
    def __init__(self, *, document_id, first_semester, second_semester):
        # self.data_from_first_semester = data_from_first_semester
        # self.data_from_second_semester = data_from_second_semester
        self.document_id = document_id
        self.first_semester = first_semester
        self.second_semester = second_semester

    def generate_statictics(self):
        #print(self.data_from_first_semester[3])

        # grades = []
        # grade = {}
        # j=4
        # for subject in self.data_from_first_semester[3]:
        #     if subject != "":
                
        #         i=2
        #         grade[subject] = []
               
        # data1  = self.data_from_first_semester
        # data1.pop(0)
        # data1.pop(0)
        # data1.pop(0)
        # for item in data1:
        #     item.pop(0)
        #     item.pop(0)
        # data1.pop(0)
        # data1.pop(len(data1)-1)
        
        
        
        
        # i=0
        # print(data1[0])
        # for key,value in grade.items():
        #     if i < len(data1)+1:
        #         for item in data1:
        #             #grade["Українська мова"].append(item[0])
        #             grade[key].append(item[i])
        #         i+=1

        
                
                

        # print(grade)
        SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        BASE_DIR = os.path.dirname(os.path.abspath(__file__))
        SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, 'credentials.json')

        credentials = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=credentials).spreadsheets().values()



        #id of document what`s splited for his link`
        spreadsheet_id = self.document_id
        #range - concatenate from list and range by symbol "!"
        range = self.first_semester+ "!" + "C4:AN38"
        
        #print(range2)


        result = service.get(spreadsheetId=spreadsheet_id,
                             range=range, majorDimension ="COLUMNS", valueRenderOption='UNFORMATTED_VALUE').execute()

        data_from_sheet = result.get('values', [])
        
        #df2 = pd.DataFrame(data_from_sheet)
        #print(df2.mean())

        #print(f"average: {str(avg)}")

        subjetcs = {}

        for subj in data_from_sheet:
            #subjetcs[subj[0]] = {}
            grd = {}
            
            count1 =0 
            count2 =0 
            count3 =0 
            count4 =0 
            count5 =0 
            count6 =0 
            count7 =0 
            count8 =0 
            count9 =0 
            count10 =0 
            count11 =0
            count12 =0  
            for grade in subj:
                
                if grade == 1:   
                    count1+=1
                if grade == 2:   
                    count2+=1
                if grade == 3:   
                    count3+=1
                if grade == 4:   
                    count4+=1
                if grade == 5:   
                    count5+=1
                if grade == 6:   
                    count6+=1
                if grade == 7:   
                    count7+=1
                if grade == 8:   
                    count8+=1
                if grade == 9:   
                    count9+=1
                if grade == 10:   
                    count10+=1
                if grade == 11:   
                    count11+=1
                if grade == 12:   
                    count12+=1
                #print(grade)
            grd["1"] = count1
            grd["2"] = count2
            grd["3"] = count3
            grd["4"] = count4
            grd["5"] = count5
            grd["6"] = count6
            grd["7"] = count7
            grd["8"] = count8
            grd["9"] = count9
            grd["10"] = count10
            grd["11"] = count11
            grd["12"] = count12
            sum = count1 + count2 + count3 + count4 + count5 + count6 + count7 + count8 + count9 + count10 + count11 + count12
            level1_count = count1 + count2 + count3
            level2_count = count4 + count5 + count6 
            level3_count = count7 + count8 + count9
            level4_count = count10 + count11 + count12
            if sum != 0:
                level1_percent = level1_count/sum*100
                level2_percent = level2_count/sum*100
                level3_percent = level3_count/sum*100
                level4_percent = level4_count/sum*100
                grd["sum"] = sum
                grd["level1_count"] = level1_count
                grd["level2_count"] = level2_count
                grd["level3_count"] = level3_count
                grd["level4_count"] = level4_count
                grd["level1_percent"] = level1_percent
                grd["level2_percent"] = level2_percent
                grd["level3_percent"] = level3_percent
                grd["level4_percent"] = level4_percent
            children_count = len(subj)-1
            if children_count != 0:
                average = (1*count1+2*count2+3*count3+4*count1+5*count5+6*count6+7*count7+8*count8+9*count9+10*count10+11*count11+12*count12)/children_count
            grd["children_count"] = children_count
            grd["average"] = average

            subjetcs[subj[0]] = grd
            subj.pop(0)
        


        # print(data_from_sheet)
        print(subjetcs)




