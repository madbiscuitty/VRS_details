from openpyxl import *


wb_form = load_workbook(filename=r'resourses/vrs_details.xlsx', read_only=True)
wb = Workbook()
ws = wb.active
ws.title = 'Project'




    vrs_num = ['019']
    print(vrs_num)
    sheet_form = wb_form['Filter-01']
    # Присваиваем диапазон ячеек согласно номеру VRS
    if vrs_num == '019':
        cells = sheet_form['A5':'D13']
        print(cells)
    elif vrs_num == '034':
        cells = sheet_form['A18':'D26']
    elif vrs_num == '039':
        cells = sheet_form['A31':'D39']
    elif vrs_num == '054':
        cells = sheet_form['A44':'D53']
    elif vrs_num == '058':
        cells = sheet_form['A58':'D67']
    elif vrs_num == '078':
        cells = sheet_form['A72':'D81']
    elif vrs_num == '086':
        cells = sheet_form['A86':'D97']
    elif vrs_num == '097':
        cells = sheet_form['A102':'D111']
    elif vrs_num == '115':
        cells = sheet_form['A116':'D127']
    elif vrs_num == '116':
        cells = sheet_form['A132':'D140']
    elif vrs_num == '138':
        cells = sheet_form['A145':'D154']
    elif vrs_num == '156':
        cells = sheet_form['A159':'D168']
    elif vrs_num == '173':
        cells = sheet_form['A173':'D184']
    elif vrs_num == '193':
        cells = sheet_form['A189':'D198']
    elif vrs_num == '194':
        cells = sheet_form['A203':'D214']
    elif vrs_num == '215':
        cells = sheet_form['A219':'D228']
    elif vrs_num == '234':
        cells = sheet_form['A233':'D242']
    elif vrs_num == '240':
        cells = sheet_form['A247':'D258']
    elif vrs_num == '271':
        cells = sheet_form['A263':'D272']
    elif vrs_num == '289':
        cells = sheet_form['A277':'D288']
    elif vrs_num == '290':
        cells = sheet_form['A293':'D304']
    elif vrs_num == '333':
        cells = sheet_form['A309':'D318']
    elif vrs_num == '337':
        cells = sheet_form['A323':'D334']
    elif vrs_num == '350':
        cells = sheet_form['A339':'D350']
    elif vrs_num == '407':
        cells = sheet_form['A355':'D366']
    elif vrs_num == '414':
        cells = sheet_form['A371':'D382']
    elif vrs_num == '473':
        cells = sheet_form['A387':'D398']
    elif vrs_num == '500':
        cells = sheet_form['A403':'D414']

if __name__ == "__main__":
    table_var()