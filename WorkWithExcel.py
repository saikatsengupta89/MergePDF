from xlrd import open_workbook
import os, string

datadriven=""

def setGlobal():
    global datadriven
    datadriven="Saikat"
    print(datadriven)

available_drives = ['%s:' % d for d in string.ascii_uppercase if os.path.exists('%s:' % d)]
print(available_drives)
print(os.path.realpath('c://').replace('\\','').upper() in available_drives)
print(os.path.isdir("C://MergePDF/Input"))

# move to the input directory and create a list of all pdfs needs to be merged
inputLocation= "C:/MergePDF/Input"
os.chdir(inputLocation)

file_location=inputLocation+"/"+'InputOrder.xlsx'
workbook= open_workbook(file_location)
worksheet= workbook.sheet_by_index(0)
column_order= worksheet.col(0)
column= worksheet.col(1)
for row in range(1, len(column_order)):
    print(column[row].value)

# wb= load_workbook(inputLocation+"/"+'InputOrder.xlsx')
# ws= wb.get_sheet_by_name('InputOrder')
# column_order= ws['A']
# column= ws['B']
list_master=[]
list_present=[]
list_pdf=[]
for row in range(1, len(column_order)):
    list_master.append(column[row].value+".pdf")
print(list_master)

for filename in os.listdir('.'):
    if filename.endswith('.pdf'):
        list_present.append(filename)
print(list_present)

for data in list_master:
    if data in list_present:
        list_pdf.append(data)
print(list_pdf)

setGlobal()
print(datadriven)