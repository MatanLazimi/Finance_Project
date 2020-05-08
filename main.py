import openpyxl
from datetime import date

def main():

    path = "data8.xlsx"
    wb_obj = openpyxl.load_workbook(path)
    sheet_obj = wb_obj.active
#---------------Salaryyyyyyyyyyyyyyyyyy--------------------------------#
#    for k in range(3, 152):
#        cell_obj = sheet_obj.cell(row=k, column=7)
#        print(float(cell_obj.value))
######################################################################


#---------------CalcSenyority--------------------------------#
    for k in range(3, 152):
        cell_obj1 = sheet_obj.cell(row=k, column=6)
        cell_obj2 = sheet_obj.cell(row=k, column=12)
        if cell_obj2.value is None or cell_obj2.value == "-" :
            print(int((date.today() - cell_obj1.value.date()).days/365.25),end=" ")
        else:
            print(int((cell_obj2.value.date() - cell_obj1.value.date()).days/365.25),end=" ")
        print()
######################################################################


main()