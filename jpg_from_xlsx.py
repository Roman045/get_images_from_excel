from os import getcwd
from win32com.client import Dispatch

n_z = ['0200', '0440', '0650', '0800', '0900', '0950', '0990']
excel_file = 'result.xlsx'
for i in range(0, 2 * len(n_z)):
    app = Dispatch('Excel.Application')
    workbook = app.Workbooks.Open(Filename=getcwd() + '\\' + excel_file)
    app.DisplayAlerts = False
    sheet = workbook.Sheets(1)
    chart = sheet.ChartObjects(1).Chart
    if(i % 2 == 0):
        chart.Export(getcwd() + f'\\excel_pressure{i + 1}.jpg')
    else:
        chart.Export(getcwd() + f'\\excel_shear{i + 1}.jpg')
    workbook.Close(False)
    app.Quit()
