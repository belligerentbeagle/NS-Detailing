import xlsxwriter
import os
from openpyxl import Workbook
import random
from openpyxl.styles import NamedStyle, Font, Border, Side, PatternFill, colors, Fill
from openpyxl.styles.differential import DifferentialStyle
from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule


workbook = Workbook()
sheet = workbook.active

sheet["A1"] = "TIME/NAME"

def assigning(row, duty):
    randomperson = random.randint(2, peoplepresent)
    if row == 2: #if first row just put only
        if sheet.cell(row = row, column = randomperson).value == None:
            sheet.cell(row = row, column = randomperson).value = duty
            sheet.cell(row = row + 1, column = randomperson).value = duty #second half of duty
        else:
            print("Replanning...")
            assigning(row, duty)
    else: #ensures they have rest before duty.
        if sheet.cell(row = row, column = randomperson).value == None and sheet.cell(row = row-1, column = randomperson).value == None:
            sheet.cell(row = row, column = randomperson).value = duty
            sheet.cell(row = row + 1, column = randomperson).value = duty #second half of duty
        else:
            print("Replanning...")
            assigning(row, duty)

def assigningpeak(row, duty): #4hourblock 1 peak 1 non-peak
    randomperson = random.randint(2, peoplepresent)
    if sheet.cell(row = row, column = randomperson).value == None and sheet.cell(row = row-1, column = randomperson).value == None:
        sheet.cell(row = row, column = randomperson).value = duty
    else:
        print("Replanning...")
        assigningpeak(row, duty)

def assigningafterpeak(counter,duty):
    if sheet.cell(row = i, column = counter).value == duty:
        sheet.cell(row = i+1, column = counter).value = duty
    else:
        counter += 1
        assigningafterpeak(counter,duty)

def cellstoleftfilled(row): #counter
    answer = 0
    for i in range(2, peoplepresent+1):
        if sheet.cell(row= row, column= i).value != None:
            answer += 1
    return answer


#set up time
noofdays = int(input("Number of days mounting: "))
row = 2
timedefault = ["1100-1300", "1300-1500","1500-1700", "1700-1900","1900-2100", "2100-2300", "2300-0100","0100-0300", "0300-0500","0500-0700","0700-0900","0900-1100"]
times = noofdays * timedefault
for timeblock in times:
    sheet.cell(row = row, column = 1).value = timeblock
    row += 1
totalrows = row

#set up humans
batch0 = ["Weijie","Maxx"]
batch1 = ["Jack", "Ivan"]
batch2 = ["Junyang", "Alvin", "Yicong", "Jowell", "Jonathan"]
batch3 = ["Bala", "Jinming", "Eugene", "Jian Yong"]
acf = ["Luke", "Ryan", "Stanley", "Yash"]
batch4 = ["Rayshawn", "Kaijie"]
batch5 = ["Denver", "Praveen"]
team = batch0 + batch1 + batch2 + batch3 + batch4 + batch5 + ["COUNTER"]
peoplepresent = len(team)
column = 2
for name in team:
    sheet.cell(row = 1, column = column).value = name
    column += 1

#colour coding
def colourthisrow(row):
    col = sheet['A1']
    col.fill = PatternFill(bgColor="00FF0000")
    row = sheet.row_dimensions[1]
    row.font = Font(underline="single")


#initialise duties
non_peak = ["TG", "18", "XSVC", "XCBT", "CCTV", "CCTV2", "VACS"]
peak = ["TG", "18", "XSVC", "XCBT", "CCTV", "CCTV2", "VACS","TG2", "XSVC2", "CHKR"]
silent = [e for e in non_peak if e not in ('XSVC', 'XCBT')]

#assign dutytypes to hours
#nonpeak = 7, peak = 10, silent = 5
row = 2 #reset row again
status = "weekday" #to be variable

for i in range(2, totalrows):
    if i%2 == 0: #iterates across even rows only so that we assign duty every 4 hours
        if status == "weekday" and (sheet.cell(row= i, column = 1).value in ["1100-1300", "1300-1500","1500-1700","1700-1900","0900-1100"]): #if non_peak on normal hours
            for duty in non_peak:
                assigning(i, duty)
            #if cell is empty (leave, off, MA etc) then put into random
        if status == "weekday" and (sheet.cell(row= i, column = 1).value in ["0700-0900"]):
            print("planning peak hours...")
            colourthisrow(i)
            for duty in peak:
                assigningpeak(i,duty)
            counter = 1 #function below for adding non-peak for 0900-1100
            for duty in non_peak:
                assigningafterpeak(counter,duty)
        if status == "weekend" or (sheet.cell(row= i, column = 1).value in ["1900-2100","2100-2300","2300-0100","0100-0300","0300-0500","0500-0700"]):
            for duty in silent:
                assigning(i,duty)
    sheet.cell(row=i, column= peoplepresent+1 ).value = cellstoleftfilled(i)
            
    








workbook.save(filename="/Users/weiyushit/OneDrive/Github stuff/NS Detailing/hello_world.xlsx")





'''
whereami = os.getcwd()
print(whereami)

#changes directory for my excel
os.chdir("/Users/weiyushit/OneDrive/Github stuff/NS Detailing")

# List all files and directories in current directory -> print(os.listdir('.'))

# Specify a writer
writer = pd.ExcelWriter('example.xlsx', engine='openpyxl')
file = 'example.xlsx'

# Load spreadsheet
xl = pd.ExcelFile(file)
'''
