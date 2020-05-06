import openpyxl

def main():

    path = "data8.xlsx"
    wb_obj = openpyxl.load_workbook(path)
    sheet_obj = wb_obj.active
#---------------Salaryyyyyyyyyyyyyyyyyy--------------------------------#
    for k in range(3, 152):
        cell_obj = sheet_obj.cell(row=k, column=7)
        #print(float(cell_obj.value))
######################################################################


#---------------Salaryyyyyyyyyyyyyyyyyy--------------------------------#
    for k in range(3, 152):
        cell_obj = sheet_obj.cell(row=k, column=7)
        print(float(cell_obj.value))
######################################################################


main()