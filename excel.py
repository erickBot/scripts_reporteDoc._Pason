import openpyxl
from datetime import datetime, date, time, timedelta
import calendar
from os import remove
import ssl
import sys

list_sheetbyName = []
list_texto = []
nameSheet = ['erick', 'Jose']
Item = ['doc1' 'doc2']
dateExpiracion = ['date1', 'date2']
diasPorExpirar = ['10', '30']

def run():
    
    file = open('reporte.txt', 'r')
    lineas = len(file.readlines())

    file = open('reporte.txt', 'w')

    for name in nameSheet:
        for item in Item:
            



if __name__ == "__main__":
    run()