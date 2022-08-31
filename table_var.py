from openpyxl import *
from tkinter import messagebox

wb_form = load_workbook(filename=r'resourses/vrs_details.xlsx', read_only=True)


def fb_table(vrs_num, vrs_perf):

    if vrs_perf == '01':
        sheet_form = wb_form['Filter-01']  # переходим на вкладку 01 исполнения
    elif vrs_perf == '02':
        sheet_form = wb_form['Filter-02']  # переходим на вкладку 02 исполнения
    elif vrs_perf == '03' or vrs_perf == '04':
        sheet_form = wb_form['Filter-03(04)']  # переходим на вкладку 03 исполнения

    filter_table = {
        '019': sheet_form['A5':'D13'],
        '034': sheet_form['A18':'D26'],
        '039': sheet_form['A31':'D39'],
        '054': sheet_form['A44':'D53'],
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

    for key in filter_table.keys():
        if key == vrs_num:
            cells = filter_table[key]
            break
        else:
            messagebox.showerror('Error', 'Что-то пошло не так')
            break
    return cells


def vent_table(vrs_num, vrs_perf, vosk):
    if vrs_perf == '01':
        sheet_form = wb_form['Vent-01']  # переходим на вкладку 01 исполнения
    elif vrs_perf == '02':
        sheet_form = wb_form['Vent-02']  # переходим на вкладку 02 исполнения
    elif vrs_perf == '03' or vrs_perf == '04':
        sheet_form = wb_form['Vent-03(04)']  # переходим на вкладку 03 исполнения

    vent_table = {
        '019' + '2.5': sheet_form['A13':'D15'],
        '019' + '2.8': sheet_form['F13':'I15'],
        '019' + '3.15': sheet_form['K13':'N15']
    }

    for key in vent_table.keys():
        if key == vrs_num + vosk:
            vent = vent_table[key]
            break
        else:
            messagebox.showerror('Error', 'Что-то пошло не так')
            break
    return vent

def vent2_table(vrs_num, vrs_perf, vosk):
    if vrs_perf == '01':
        sheet_form = wb_form['Vent-01']  # переходим на вкладку 01 исполнения
    elif vrs_perf == '02':
        sheet_form = wb_form['Vent-02']  # переходим на вкладку 02 исполнения
    elif vrs_perf == '03' or vrs_perf == '04':
        sheet_form = wb_form['Vent-03(04)']  # переходим на вкладку 03 исполнения

    vent_table = {
        '019' + '2.5': sheet_form['A13':'D15'],
        '019' + '2.8': sheet_form['F13':'I15'],
        '019' + '3.15': sheet_form['K13':'N15']
    }

    for key in vent_table.keys():
        if key == vrs_num + vosk:
            vent2 = vent_table[key]
            break
        else:
            messagebox.showerror('Error', 'Что-то пошло не так')
            break
    return vent2

def air_table(vrs_perf, air):
    if vrs_perf == '01':
        sheet_form = wb_form['Vent-01']  # переходим на вкладку 01 исполнения
    elif vrs_perf == '02':
        sheet_form = wb_form['Vent-02']  # переходим на вкладку 02 исполнения
    elif vrs_perf == '03' or vrs_perf == '04':
        sheet_form = wb_form['Vent-03(04)']  # переходим на вкладку 03 исполнения

    air_table = {
        'АИР56': sheet_form['A4':'D9'],
        'АИР63': sheet_form['F4':'I9'],
        'АИР71': sheet_form['K4':'N9'],
        'АИР80': sheet_form['P4':'S9']
    }

    for key in air_table.keys():
        if key == air:
            air = air_table[key]
            break
        else:
            messagebox.showerror('Error', 'Что-то пошло не так')
            break
    return air

def air2_table(vrs_perf, air2):
    if vrs_perf == '01':
        sheet_form = wb_form['Vent-01']  # переходим на вкладку 01 исполнения
    elif vrs_perf == '02':
        sheet_form = wb_form['Vent-02']  # переходим на вкладку 02 исполнения
    elif vrs_perf == '03' or vrs_perf == '04':
        sheet_form = wb_form['Vent-03(04)']  # переходим на вкладку 03 исполнения

    air_table = {
        'АИР56': sheet_form['A4':'D9'],
        'АИР63': sheet_form['F4':'I9'],
        'АИР71': sheet_form['K4':'N9'],
        'АИР80': sheet_form['P4':'S9']
    }

    for key in air_table.keys():
        if key == air2:
            air2 = air_table[key]
            break
        else:
            messagebox.showerror('Error', 'Что-то пошло не так')
            break
    return air2
def vnv5012_table():
    pass

def v0v5012_table():
    pass

def eko_table():
    pass

def vertklap_table():
    pass

def pp_table():
    pass

def promkam_table():
    pass