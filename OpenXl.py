import openpyxl as xl
import os
import logging as log

log.basicConfig()

workbookName = input('Please enter the workbook name:\n')
workbook = xl.open_workbook(f'{workbookName}.xlsx')