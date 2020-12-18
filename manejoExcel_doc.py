import openpyxl
from datetime import datetime, date, time, timedelta
import calendar
from openpyxl.formatting.rule import ColorScaleRule

list_DaysTecnico1 = []
list_DaysTecnico2 = []
list_car1 = []
list_car2 = []
count_tecnico1 = []
count_tecnico2 = []
count_car1 = []
count_car2 = []

def run():
    name_file = "Control_Documentos.xlsx"
    file = openpyxl.load_workbook(name_file)
    sheet1 = file.get_sheet_by_name('Erick Pasache')
    sheet2 = file.get_sheet_by_name('Johnnel Panduro')
    sheet3 = file.get_sheet_by_name('A1I827')
    sheet4 = file.get_sheet_by_name('C9L911')
    cell_tecnico1 = sheet1['D4':'D14']
    cell_tecnico2 = sheet2['D4':'D14']
    cell_vehiculo1 = sheet3['C4':'C5']
    cell_vehiculo2 = sheet4['C4':'C5']
    dateNow = datetime.now()
    lookCell(cell_tecnico1, list_DaysTecnico1, count_tecnico1, dateNow)
    lookCell(cell_tecnico2, list_DaysTecnico2, count_tecnico2, dateNow)
    lookCell(cell_vehiculo1, list_car1, count_car1, dateNow)
    lookCell(cell_vehiculo2, list_car2, count_car2, dateNow)
    #salta a las funciones para llenar las celdas con datos nuevos
    llenar_cell(sheet1, list_DaysTecnico1, count_tecnico1, 'F4', 'F14')
    llenar_cell(sheet2, list_DaysTecnico2, count_tecnico2, 'F4', 'F14')
    llenar_cell_Estado(sheet1, list_DaysTecnico1, 'E4', 'E14')
    llenar_cell_Estado(sheet2, list_DaysTecnico2, 'E4', 'E14')
    llenar_cell(sheet3, list_car1, count_car1, 'E4', 'E5')
    llenar_cell(sheet4, list_car2, count_car2, 'E4', 'E5')
    llenar_cell_Estado(sheet3, list_car1, 'D4', 'D5')
    llenar_cell_Estado(sheet4, list_car2, 'D4', 'D5')
    #guarda archivo
    file.save(name_file)

    #print(count_Erick)
    #print(count_Johnnel)

def lookCell(celdas, list_DaysTemp, count_temp, dateNow):
    expirados = 0
    porVencer = 0
    vigentes = 0

    for row in celdas:
        for cell in row:
            dateCell= cell.value
            days = dateCell - dateNow
            days = days.days
            list_DaysTemp.append(days)

            if days < 0:
                expirados += 1
            if days > 0 and days < 30:
                porVencer += 1
            if days > 30:
                vigentes += 1

    count_temp.append(porVencer)
    count_temp.append(expirados)
    count_temp.append(vigentes)


def llenar_cell(sheet, list_Days, countList, cellInit, cellEnd):
    i = 0
    cell_list = sheet[cellInit : cellEnd]
    for row in cell_list:
        for cell in row:
            cell.value = list_Days[i]
            i += 1

    sheet['C21'] = countList[0] #porVencer
    sheet['C22'] = countList[1] #expirados
    sheet['C23'] = countList[2] #vigentes
    print(list_Days)

def llenar_cell_Estado(sheet, list_Days, cellInit, cellEnd):
    i = 0
    cell_list = sheet[cellInit : cellEnd]
    for row in cell_list:
        for cell in row:
            if list_Days[i] > 0:
                cell.value = "VIGENTE"
            if list_Days[i] < 0:
                cell.value = "EXPIRADO"
            i += 1

if __name__ == "__main__":
    run()
