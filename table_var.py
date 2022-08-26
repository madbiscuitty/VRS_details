from openpyxl import *

wb_form = load_workbook(filename=r'resourses/vrs_details.xlsx', read_only=True)
sheet_form = wb_form['Filter-01']

vrs_num = '019'

filter_block = {
    '019': sheet_form['A5':'D13'],
    '034': sheet_form['A18':'D26'],
    '039': sheet_form['A31':'D39'],
    '058': sheet_form['A58':'D67'],
    '078': sheet_form['A72':'D81'],
    '086': sheet_form['A86':'D97'],
    '097': sheet_form['A102':'D111'],
    '115': sheet_form['A116':'D127'],
    '116': sheet_form['A132':'D140'],
    '138': sheet_form['A145':'D154'],
    '156': sheet_form['A159':'D168'],
    '173': sheet_form['A173':'D184'],
    '193': sheet_form['A189':'D198'],
    '194': sheet_form['A203':'D214'],
    '215': sheet_form['A219':'D228'],
    '234': sheet_form['A233':'D242'],
    '240': sheet_form['A247':'D258'],
    '271': sheet_form['A263':'D272'],
    '289': sheet_form['A277':'D288'],
    '290': sheet_form['A293':'D304'],
    '333': sheet_form['A309':'D318'],
    '337': sheet_form['A323':'D334'],
    '350': sheet_form['A339':'D350'],
    '407': sheet_form['A355':'D366'],
    '414': sheet_form['A371':'D382'],
    '473': sheet_form['A387':'D398'],
    '500': sheet_form['A403':'D414']
}

vent_block = {
    '019': sheet_form['A5':'D13'],
    '034': sheet_form['A18':'D26'],
    '039': sheet_form['A31':'D39'],
    '058': sheet_form['A58':'D67'],
    '078': sheet_form['A72':'D81'],
    '086': sheet_form['A86':'D97'],
    '097': sheet_form['A102':'D111'],
    '115': sheet_form['A116':'D127'],
    '116': sheet_form['A132':'D140'],
    '138': sheet_form['A145':'D154'],
    '156': sheet_form['A159':'D168'],
    '173': sheet_form['A173':'D184'],
    '193': sheet_form['A189':'D198'],
    '194': sheet_form['A203':'D214'],
    '215': sheet_form['A219':'D228'],
    '234': sheet_form['A233':'D242'],
    '240': sheet_form['A247':'D258'],
    '271': sheet_form['A263':'D272'],
    '289': sheet_form['A277':'D288'],
    '290': sheet_form['A293':'D304'],
    '333': sheet_form['A309':'D318'],
    '337': sheet_form['A323':'D334'],
    '350': sheet_form['A339':'D350'],
    '407': sheet_form['A355':'D366'],
    '414': sheet_form['A371':'D382'],
    '473': sheet_form['A387':'D398'],
    '500': sheet_form['A403':'D414']
}

def table(vrs_num, table):
    for key in table.keys():
        if key == vrs_num:
            cells = table[key]
    return cells


