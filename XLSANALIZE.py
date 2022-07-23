import xlsxwriter
import tkinter as tk
from tkinter import filedialog
import xlrd2 as xlrd

class XLSRead():
    def __init__(self):
        self.book2open = xlrd.open_workbook(self.find_file())
        self.SheetsFormat,self.SheetsLabel = self.Sheets()
        self.book2save = xlsxwriter.Workbook(self.path2save())
        self.sheet2save = self.book2save.add_worksheet()        
        self.Operator(self.getValues(self.SheetsFormat[2]), float(input('K1 = ')),float(input('K2 = ')),'-',int(input('column = ')),int(input('SplitPoint = ')))
        self.book2save.close()
        self.book2open.release_resources()
        del self.book2open
  
        
        


    def find_file(self, *args):
        pathw = tk.Tk()
        pathw.withdraw()
        self.file_path = filedialog.askopenfilename(title = 'Open file', defaultextension='.xlsx', filetypes = (('XLSX file','*.xlsx'),('XLS file','*.xls')))
        return self.file_path
    
    def path2save(self):
        E,path2save,flag = 0,'',0
        for i in range(len(self.file_path)):
            if(self.file_path[i] != '.' and E == 0): path2save += self.file_path[i]; flag += 1
            else: E = 1
        path2save += 'Modified'+self.file_path[flag:]
        return path2save

    def getSheets(self):
        i,self.sheets2open = 0,[]
        while True:
            try: self.sheets2open.append(self.book2open.sheet_by_index(i))
            except: break
            i += 1
        return self.sheets2open

    def Sheets(self):
        SheetsFormat,SheetsLabel = self.getSheets(),[]
        for i in range(len(SheetsFormat)):
            if(len(SheetsFormat) < 11): SheetsLabel.append(str(SheetsFormat[i])[10:len(str(SheetsFormat[i]))-1])
            else: SheetsLabel.append(str(SheetsFormat[i])[11:len(str(SheetsFormat[i]))-1])
        return SheetsFormat,SheetsLabel

    def getValues(self, sheet_id):
        i,k,row,column = 0,0,[],[]
        while True:
            try:
                column.append(sheet_id.cell_value(0, i))
            except: break
            i += 1
        while True:
            try:
                row.append(sheet_id.cell_value(k, 0))
            except: break
            k += 1
        del i,k
        values = [[] for i in range(len(column))]
        for i in range(len(values)):
            for k in range(len(row)):
                values[i].append(sheet_id.cell_value(k, i))
        for i in range(len(values)):
            for k in range(len(row)):
                if(str(type(values[i][k])) != "<class 'str'>"):
                    if(values[i][k]%2 == 0 or values[i][k]%3 == 0 or values[i][k]%int(values[i][k]) == 0): values[i][k] = int(values[i][k])
        del row,column
        return values

    def Operator(self, values, K1, K2, OP, column, SplitPoint):
        ans = []
        if(OP == '-'):
            for i in range(len(values[column])):
                if(i < SplitPoint-1): ans.append(values[column][i]-K1)
                if(i >= SplitPoint-1): ans.append(values[column][i]-K2); print(values[column][i],K2,values[column][i]-K2)

        for i in range(len(ans)):
            if(str(type(ans[i])) != "<class 'str'>"):
                try:
                    if(ans[i]%2 == 0 or ans[i]%3 == 0 or ans[i]%int(ans[i]) == 0): ans[i] = int(ans[i])
                except: pass

        self.WriteFile(values,ans)

        #print(ans)

    def WriteFile(self, org, values):
        for i in range(len(values)):
            self.sheet2save.write(i,0,org[0][i])
            self.sheet2save.write(i,1,values[i])
        

XLSRead()



"""
 
# For row 0 and column 0
print(sheet.cell_value(0, 0))
"""
"""
hoja.write(0,0,'PM2.5')
hoja.write(0,1,'PM10')
hoja.write(0,2,'Temperature')
hoja.write(0,3,'Pressure')
hoja.write(0,4,'Humidity')
hoja.write(0,5,'Date/Time')

for i in range(0,len(data)):
    for k in range(0,len(data[i])):
        hoja.write(i+1,k,data[i][k])

libro.close()

print('Done')
"""
