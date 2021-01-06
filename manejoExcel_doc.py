import openpyxl
from datetime import datetime, date, time, timedelta
import calendar
from openpyxl.formatting.rule import ColorScaleRule

list_sheetbyName = []
list_diasVigencia = []
list_contador = []

def run():
    name_file = "Control_Documentos.xlsx"
    file = openpyxl.load_workbook(name_file)
    #<file.sheetnames> devuelve los nombres de las hojas en una lista
    list_sheetbyName = file.sheetnames

    for name in list_sheetbyName:
        sheet = file.get_sheet_by_name(name)
        #devuelve el numero maximo de filas por hoja (sheet)
        num_row = sheet.max_row
        cell_total = sheet['D4':'D' + str(num_row)]
        #devuelve la fecha actual del computador
        dateNow = datetime.now()
        lookCell(cell_total, dateNow)
        #salta a las funciones para llenar las celdas con datos nuevos
        llenar_cell(sheet, 'F4', 'F' + str(num_row))
        llenar_cell_Estado(sheet, 'E4', 'E' + str(num_row) )
        #vaciar listas
        list_diasVigencia.clear()
        list_contador.clear()
    
    #guarda archivo
    file.save(name_file)

def lookCell(cell_total, dateNow):
    expirados = 0
    porVencer = 0
    vigentes = 0

    for row in cell_total:
        for cell in row:
            dateCell= cell.value
            dias = dateCell - dateNow
            dias = dias.days
            list_diasVigencia.append(dias)

            if dias < 0:
                expirados += 1
            if dias > 0 and dias < 30:
                porVencer += 1
            if dias > 30:
                vigentes += 1

    list_contador.append(porVencer)
    list_contador.append(expirados)
    list_contador.append(vigentes)


def llenar_cell(sheet, cellInit, cellEnd):
    i = 0
    cell_list = sheet[cellInit : cellEnd]
    for row in cell_list:
        for cell in row:
            cell.value = list_diasVigencia[i]
            i += 1

    sheet['L3'] = list_contador[0] #porVencer
    sheet['L4'] = list_contador[1] #expirados
    sheet['L5'] = list_contador[2] #vigentes

    print(list_diasVigencia)

def llenar_cell_Estado(sheet, cellInit, cellEnd):
    i = 0
    cell_list = sheet[cellInit : cellEnd]
    for row in cell_list:
        for cell in row:
            if list_diasVigencia[i] > 0:
                cell.value = "VIGENTE"
            if list_diasVigencia[i] < 0:
                cell.value = "EXPIRADO"
            i += 1

if __name__ == "__main__":
    run()
