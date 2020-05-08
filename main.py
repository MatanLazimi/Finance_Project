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
        cell_obj0 = sheet_obj.cell(row=k, column=2)
        cell_obj1 = sheet_obj.cell(row=k, column=6)
        cell_obj2 = sheet_obj.cell(row=k, column=8)
        cell_obj3 = sheet_obj.cell(row=k, column=4)
        cell_obj4 = sheet_obj.cell(row=k, column=1)
        cell_obj5 = sheet_obj.cell(row=k, column=5)
        cell_obj6 = sheet_obj.cell(row=k, column=7)
        if cell_obj2.value is None or cell_obj2.value == "-" :
            seniority = (date.today() - cell_obj1.value.date()).days/365.25
            if cell_obj2.value is None:
                non_article14 = seniority
            else:
                non_article14 = ((cell_obj2.value.date()).days - (cell_obj1.value.date()).days)/365.25
            """Time for years of article 14"""
            article14 = seniority - non_article14
            """man is 0, female is 1."""
            if cell_obj3.value == "M":
                gender = 0
            else:
                gender = 1
            id = int(cell_obj4.value)
            if int(cell_obj4.value) % 2 == 0:
                salary_growth_rate = 0.04
            else:
                salary_growth_rate = 0.02
            age = (date.today() - cell_obj5.value.date()).days/365.25
            last_salary = int(cell_obj6.value)
            name = cell_obj0.value
            print("name - ",name," ID - ",id," gender - ",gender," age - ",age," salary - ",last_salary," seniority - ",seniority," non_article14 - ",non_article14," article14 - ",article14," rate - ",salary_growth_rate)
        else:
            pass
            """
            print('left: ',(cell_obj2.value.date() - cell_obj1.value.date()).days/365.25,end=" ")
            """
######################################################################


main()