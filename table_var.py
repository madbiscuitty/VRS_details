from openpyxl import *
from tkinter import messagebox

wb_form = load_workbook(filename=r'resourses/vrs_details.xlsx', read_only=True)


def filter_table(vrs_num, vrs_perf):
    if vrs_perf == '01':
        sheet_form = wb_form['Filter-01']  # переходим на вкладку 01 исполнения
    elif vrs_perf == '02':
        sheet_form = wb_form['Filter-02']  # переходим на вкладку 02 исполнения
    elif vrs_perf == '03' or vrs_perf == '04':
        sheet_form = wb_form['Filter-03(04)']  # переходим на вкладку 03 исполнения

    filter_list = {
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
    try:
        for key in filter_list.keys():
            if key == vrs_num:
                cells = filter_list[key]
                break
        return cells
    except UnboundLocalError:
        messagebox.showerror('Error', 'Нет данных в таблице для блока\nфильтров VRS ' + vrs_num)


def vent_table(vrs_num, vrs_perf, vosk):
    if vrs_perf == '01':
        sheet_form = wb_form['Vent-01']  # переходим на вкладку 01 исполнения
    elif vrs_perf == '02':
        sheet_form = wb_form['Vent-02']  # переходим на вкладку 02 исполнения
    elif vrs_perf == '03' or vrs_perf == '04':
        sheet_form = wb_form['Vent-03(04)']  # переходим на вкладку 03 исполнения

    vent_list = {
        '019' + '2.5': sheet_form['A13':'D15'],
        '019' + '2.8': sheet_form['F13':'I15'],
        '019' + '3.15': sheet_form['K13':'N15']
    }
    try:
        for key in vent_list.keys():
            if key == vrs_num + vosk:
                vent = vent_list[key]
                break
        return vent
    except UnboundLocalError:
        messagebox.showerror('Error', 'Нет данных в таблице для ВНК блока ВОСК ' + vosk)

def vent_table2(vrs_num, vrs_perf, vosk2):
    if vrs_perf == '01':
        sheet_form = wb_form['Vent-01']  # переходим на вкладку 01 исполнения
    elif vrs_perf == '02':
        sheet_form = wb_form['Vent-02']  # переходим на вкладку 02 исполнения
    elif vrs_perf == '03' or vrs_perf == '04':
        sheet_form = wb_form['Vent-03(04)']  # переходим на вкладку 03 исполнения

    vent_list = {
        '019' + '2.5': sheet_form['A13':'D15'],
        '019' + '2.8': sheet_form['F13':'I15'],
        '019' + '3.15': sheet_form['K13':'N15']
    }
    try:
        for key in vent_list.keys():
            if key == vrs_num + vosk2:
                vent2 = vent_list[key]
                break
        return vent2
    except UnboundLocalError:
        messagebox.showerror('Error', 'Нет данных в таблице для ВНК блока ВОСК ' + vosk2)

def air_table(vrs_perf, air, vosk):
    if vrs_perf == '01':
        sheet_form = wb_form['Vent-01']  # переходим на вкладку 01 исполнения
    elif vrs_perf == '02':
        sheet_form = wb_form['Vent-02']  # переходим на вкладку 02 исполнения
    elif vrs_perf == '03' or vrs_perf == '04':
        sheet_form = wb_form['Vent-03(04)']  # переходим на вкладку 03 исполнения

    air_list = {
        '2.5' + 'АИР56': sheet_form['A4':'D9'],
        '2.5' + 'АИР63': sheet_form['F4':'I9'],
        '2.5' + 'АИР71': sheet_form['K4':'N9'],
        '2.5' + 'АИР80': sheet_form['P4':'S9']
    }
    try:
        for key in air_list.keys():
            if key == vosk + air:
               var_air = air_list[key]
        return var_air
    except UnboundLocalError:
        messagebox.showerror('Error', 'Нет данных в таблице для вентагрегата ВОСК ' + vosk + ' ' + air)

def air_table2(vrs_perf, air2, vosk2):
    if vrs_perf == '01':
        sheet_form = wb_form['Vent-01']  # переходим на вкладку 01 исполнения
    elif vrs_perf == '02':
        sheet_form = wb_form['Vent-02']  # переходим на вкладку 02 исполнения
    elif vrs_perf == '03' or vrs_perf == '04':
        sheet_form = wb_form['Vent-03(04)']  # переходим на вкладку 03 исполнения

    air_list = {
        '2.5' + 'АИР56': sheet_form['A4':'D9'],
        '2.5' + 'АИР63': sheet_form['F4':'I9'],
        '2.5' + 'АИР71': sheet_form['K4':'N9'],
        '2.5' + 'АИР80': sheet_form['P4':'S9']
    }
    try:
        for key in air_list.keys():
            if key == vosk2 + air2:
                var_air2 = air_list[key]
                break
        return var_air2
    except UnboundLocalError:
        messagebox.showerror('Error', 'Нет данных в таблице для вентагрегата ВОСК ' + vosk2 + ' ' + air)

def vnv5012_table(vrs_num, vrs_perf, vnv5012):
    if vrs_perf == '01':
        sheet_form = wb_form['Vnv5012-01']  # переходим на вкладку 01 исполнения
    elif vrs_perf == '02':
        sheet_form = wb_form['Vnv5012-02']  # переходим на вкладку 02 исполнения
    elif vrs_perf == '03' or vrs_perf == '04':
        sheet_form = wb_form['Vnv5012-03(04)']  # переходим на вкладку 03 исполнения

    vnv5012_list = {
        '019' + '160': sheet_form['A4':'D7'],
        '034' + '160': sheet_form['A11':'D14'],
        '039' + '160': sheet_form['A18':'D21'],
        '054' + '160': sheet_form['A25':'D28'],
        '058' + '160': sheet_form['A32':'D35'],
        '078' + '160': sheet_form['A39':'D42'],
        '086' + '160': sheet_form['A46':'D49'],
        '097' + '160': sheet_form['A53':'D56'],
        '115' + '160': sheet_form['A60':'D63'],
        '116' + '160': sheet_form['A67':'D70'], '116' + '180': sheet_form['F67':'I70'],
        '138' + '160': sheet_form['A74':'D77'],
        '156' + '160': sheet_form['A81':'D84'], '156' + '180': sheet_form['F81':'I84'],
        '173' + '160': sheet_form['A88':'D91'],
        '193' + '160': sheet_form['A95':'D98'], '193' + '180': sheet_form['F95':'I98'],
        '194' + '160': sheet_form['A102':'D105'], '194' + '180': sheet_form['F102':'I105'],
        '215' + '160': sheet_form['A109':'D112'], '215' + '180': sheet_form['F109':'I112'],
        '234' + '160': sheet_form['A116':'D119'], '234' + '180': sheet_form['F116':'I119'],
        '240' + '160': sheet_form['A124':'D127'], '240' + '180': sheet_form['F124':'I127'],
        '271' + '160': sheet_form['A131':'D134'], '271' + '180': sheet_form['F131':'I134'],
        '289' + '160': sheet_form['A138':'D141'], '289' + '180': sheet_form['F138':'I141'],
        '290' + '160': sheet_form['A145':'D148'], '290' + '180': sheet_form['F145':'I148'],
        '333' + '160': sheet_form['A152':'D155'], '333' + '180': sheet_form['F152':'I155'],
        '337' + '160': sheet_form['A159':'D162'], '337' + '180': sheet_form['F159':'I162'],
        '350' + '160': sheet_form['A166':'D169'], '350' + '180': sheet_form['F166':'I169'],
        '407' + '160': sheet_form['A173':'D176'], '407' + '180': sheet_form['F173':'I176'],
        '414' + '160': sheet_form['A180':'D183'], '414' + '180': sheet_form['F180':'I183'],
        '473' + '160': sheet_form['A187':'D190'], '473' + '180': sheet_form['F187':'I190'],
        '500' + '160': sheet_form['A194':'D197'], '500' + '180': sheet_form['F194':'I197']
    }
    try:
        for key in vnv5012_list.keys():
            if key == vrs_num + vnv5012:
                var_vnv5012 = vnv5012_list[key]
        return var_vnv5012
    except UnboundLocalError:
        messagebox.showerror('Error', 'Нет данных в таблице для\nблока ВНВ 5012 VRS ' + vrs_num)

def vov5012_table(vrs_num, vrs_perf, vov5012):
    if vrs_perf == '01':
        sheet_form = wb_form['Vov5012-01']  # переходим на вкладку 01 исполнения
    elif vrs_perf == '02':
        sheet_form = wb_form['Vov5012-02']  # переходим на вкладку 02 исполнения
    elif vrs_perf == '03' or vrs_perf == '04':
        sheet_form = wb_form['Vov5012-03(04)']  # переходим на вкладку 03 исполнения

    vov5012_list = {
        '019' + '180': sheet_form['A4':'D13'], '019' + '220': sheet_form['F4':'I13'],
        '019' + '280': sheet_form['K4':'N13'], '019' + '360': sheet_form['P4':'S13'],
        '034' + '180': sheet_form['A4':'D7'], '034' + '220': sheet_form['A4':'D7'],
        '034' + '280': sheet_form['A4':'D7'], '034' + '360': sheet_form['A4':'D7'],
        '039' + '180': sheet_form['A4':'D7'], '039' + '220': sheet_form['A4':'D7'],
        '039' + '280': sheet_form['A4':'D7'], '039' + '360': sheet_form['A4':'D7'],
        '054' + '180': sheet_form['A4':'D7'], '054' + '220': sheet_form['A4':'D7'],
        '054' + '280': sheet_form['A4':'D7'], '054' + '360': sheet_form['A4':'D7'],
        '058' + '180': sheet_form['A4':'D7'], '058' + '220': sheet_form['A4':'D7'],
        '058' + '280': sheet_form['A4':'D7'], '058' + '360': sheet_form['A4':'D7'],
        '078' + '180': sheet_form['A4':'D7'], '078' + '220': sheet_form['A4':'D7'],
        '078' + '280': sheet_form['A4':'D7'], '078' + '360': sheet_form['A4':'D7'],
        '086' + '180': sheet_form['A4':'D7'], '086' + '220': sheet_form['A4':'D7'],
        '086' + '280': sheet_form['A4':'D7'], '086' + '360': sheet_form['A4':'D7'],
        '097' + '180': sheet_form['A4':'D7'], '097' + '220': sheet_form['A4':'D7'],
        '097' + '280': sheet_form['A4':'D7'], '097' + '360': sheet_form['A4':'D7'],
        '115' + '180': sheet_form['A4':'D7'], '115' + '220': sheet_form['A4':'D7'],
        '115' + '280': sheet_form['A4':'D7'], '115' + '360': sheet_form['A4':'D7'],
        '116' + '180': sheet_form['A4':'D7'], '116' + '220': sheet_form['A4':'D7'],
        '116' + '280': sheet_form['A4':'D7'], '116' + '360': sheet_form['A4':'D7'],
        '138' + '180': sheet_form['A4':'D7'], '138' + '220': sheet_form['A4':'D7'],
        '138' + '280': sheet_form['A4':'D7'], '138' + '360': sheet_form['A4':'D7'],
        '156' + '180': sheet_form['A4':'D7'], '156' + '220': sheet_form['A4':'D7'],
        '156' + '280': sheet_form['A4':'D7'], '156' + '360': sheet_form['A4':'D7'],
        '173' + '180': sheet_form['A4':'D7'], '173' + '220': sheet_form['A4':'D7'],
        '173' + '280': sheet_form['A4':'D7'], '173' + '360': sheet_form['A4':'D7'],
        '193' + '180': sheet_form['A4':'D7'], '193' + '220': sheet_form['A4':'D7'],
        '193' + '280': sheet_form['A4':'D7'], '193' + '360': sheet_form['A4':'D7'],
        '194' + '180': sheet_form['A4':'D7'], '194' + '220': sheet_form['A4':'D7'],
        '194' + '280': sheet_form['A4':'D7'], '194' + '360': sheet_form['A4':'D7'],
        '215' + '180': sheet_form['A4':'D7'], '215' + '220': sheet_form['A4':'D7'],
        '215' + '280': sheet_form['A4':'D7'], '215' + '360': sheet_form['A4':'D7'],
        '234' + '180': sheet_form['A4':'D7'], '234' + '220': sheet_form['A4':'D7'],
        '234' + '280': sheet_form['A4':'D7'], '234' + '360': sheet_form['A4':'D7'],
        '240' + '180': sheet_form['A4':'D7'], '240' + '220': sheet_form['A4':'D7'],
        '240' + '280': sheet_form['A4':'D7'], '240' + '360': sheet_form['A4':'D7'],
        '271' + '180': sheet_form['A4':'D7'], '271' + '220': sheet_form['A4':'D7'],
        '271' + '280': sheet_form['A4':'D7'], '271' + '360': sheet_form['A4':'D7'],
        '289' + '180': sheet_form['A4':'D7'], '289' + '220': sheet_form['A4':'D7'],
        '289' + '280': sheet_form['A4':'D7'], '289' + '360': sheet_form['A4':'D7'],
        '290' + '180': sheet_form['A4':'D7'], '290' + '220': sheet_form['A4':'D7'],
        '290' + '280': sheet_form['A4':'D7'], '290' + '360': sheet_form['A4':'D7'],
        '333' + '180': sheet_form['A4':'D7'], '333' + '220': sheet_form['A4':'D7'],
        '333' + '280': sheet_form['A4':'D7'], '333' + '360': sheet_form['A4':'D7'],
        '337' + '180': sheet_form['A4':'D7'], '337' + '220': sheet_form['A4':'D7'],
        '337' + '280': sheet_form['A4':'D7'], '337' + '360': sheet_form['A4':'D7'],
        '350' + '180': sheet_form['A4':'D7'], '350' + '220': sheet_form['A4':'D7'],
        '350' + '280': sheet_form['A4':'D7'], '350' + '360': sheet_form['A4':'D7'],
        '407' + '180': sheet_form['A4':'D7'], '407' + '220': sheet_form['A4':'D7'],
        '407' + '280': sheet_form['A4':'D7'], '407' + '360': sheet_form['A4':'D7'],
        '414' + '180': sheet_form['A4':'D7'], '414' + '220': sheet_form['A4':'D7'],
        '414' + '280': sheet_form['A4':'D7'], '414' + '360': sheet_form['A4':'D7'],
        '473' + '180': sheet_form['A4':'D7'], '473' + '220': sheet_form['A4':'D7'],
        '473' + '280': sheet_form['A4':'D7'], '473' + '360': sheet_form['A4':'D7'],
        '500' + '180': sheet_form['A4':'D7'], '500' + '220': sheet_form['A4':'D7'],
        '500' + '280': sheet_form['A4':'D7'], '500' + '360': sheet_form['A4':'D7']
    }
    try:
        for key in vov5012_list.keys():
            if key == vrs_num + vov5012:
                var_vov5012 = vov5012_list[key]
        return var_vov5012
    except UnboundLocalError:
        messagebox.showerror('Error', 'Нет данных в таблице для\nблока ВОВ 5012 VRS ' + vrs_num)

def eko_table(vrs_num, vrs_perf):
    if vrs_perf == '01':
        sheet_form = wb_form['Eko-01']  # переходим на вкладку 01 исполнения
    elif vrs_perf == '02':
        sheet_form = wb_form['Eko-02']  # переходим на вкладку 02 исполнения
    elif vrs_perf == '03' or vrs_perf == '04':
        sheet_form = wb_form['Eko-03(04)']  # переходим на вкладку 03 исполнения

    eko_list = {
        '019': sheet_form['A4':'D6'],
        '034': sheet_form['A10':'D12'],
        '039': sheet_form['A16':'D18'],
        '054': sheet_form['A22':'D24'],
        '058': sheet_form['A28':'D30'],
        '078': sheet_form['A34':'D36'],
        '086': sheet_form['A40':'D42'],
        '097': sheet_form['A46':'D48'],
        '115': sheet_form['A52':'D54'],
        '116': sheet_form['A58':'D60'],
        '138': sheet_form['A64':'D66'],
        '156': sheet_form['A70':'D72'],
        '173': sheet_form['A76':'D78'],
        '193': sheet_form['A82':'D84'],
        '194': sheet_form['A88':'D90'],
        '215': sheet_form['A94':'D96'],
        '234': sheet_form['A100':'D102'],
        '240': sheet_form['A106':'D108'],
        '271': sheet_form['A112':'D114'],
        '289': sheet_form['A118':'D120'],
        '290': sheet_form['A124':'D126'],
        '333': sheet_form['A130':'D132'],
        '337': sheet_form['A136':'D138'],
        '350': sheet_form['A142':'D144'],
        '407': sheet_form['A148':'D150'],
        '414': sheet_form['A154':'D156'],
        '473': sheet_form['A160':'D162'],
        '500': sheet_form['A166':'D168']
    }
    try:
        for key in eko_list.keys():
            if key == vrs_num:
                var_eko = eko_list[key]
                break
        return var_eko
    except UnboundLocalError:
        messagebox.showerror('Error', 'Нет данных в таблице  для\nблока ЭКО VRS ' + vrs_num)

def vertklap_table(vrs_num, vrs_perf):
    if vrs_perf == '01':
        sheet_form = wb_form['Vertklap-01']  # переходим на вкладку 01 исполнения
    elif vrs_perf == '02':
        sheet_form = wb_form['Vertklap-02']  # переходим на вкладку 02 исполнения
    elif vrs_perf == '03' or vrs_perf == '04':
        sheet_form = wb_form['Vertklap-03(04)']  # переходим на вкладку 03 исполнения

    vertklap_list = {
        '019': sheet_form['A4':'D12'],
        '034': sheet_form['A16':'D24'],
        '039': sheet_form['A28':'D36'],
        '054': sheet_form['A40':'D48'],
        '058': sheet_form['A52':'D60'],
        '078': sheet_form['A64':'D72'],
        '086': sheet_form['A76':'D84'],
        '097': sheet_form['A88':'D96'],
        '115': sheet_form['A100':'D108'],
        '116': sheet_form['A112':'D120'],
        '138': sheet_form['A124':'D132'],
        '156': sheet_form['A136':'D144'],
        '173': sheet_form['A148':'D156'],
        '193': sheet_form['A160':'D168'],
        '194': sheet_form['A172':'D180'],
        '215': sheet_form['A184':'D192'],
        '234': sheet_form['A196':'D204'],
        '240': sheet_form['A208':'D216'],
        '271': sheet_form['A220':'D228'],
        '289': sheet_form['A232':'D240'],
        '290': sheet_form['A244':'D252'],
        '333': sheet_form['A256':'D264'],
        '337': sheet_form['A268':'D276'],
        '350': sheet_form['A280':'D288'],
        '407': sheet_form['A292':'D300'],
        '414': sheet_form['A304':'D312'],
        '473': sheet_form['A316':'D324'],
        '500': sheet_form['A328':'D336']
    }
    try:
        for key in vertklap_list.keys():
            if key == vrs_num:
                var_vertklap = vertklap_list[key]
                break
        return var_vertklap
    except UnboundLocalError:
        messagebox.showerror('Error', 'Нет данных в таблице для\nблока ВНВ 5012 VRS ' + vrs_num)

def pp_table(vrs_num, vrs_perf):
    if vrs_perf == '01':
        sheet_form = wb_form['PP-01']  # переходим на вкладку 01 исполнения
    elif vrs_perf == '02':
        sheet_form = wb_form['PP-02']  # переходим на вкладку 02 исполнения
    elif vrs_perf == '03' or vrs_perf == '04':
        sheet_form = wb_form['PP-03(04)']  # переходим на вкладку 03 исполнения

    pp_list = {
        '019': sheet_form['A4':'D10'],
        '034': sheet_form['A14':'D20'],
        '039': sheet_form['A24':'D30'],
        '054': sheet_form['A34':'D40'],
        '058': sheet_form['A44':'D50'],
        '078': sheet_form['A54':'D60'],
        '086': sheet_form['A64':'D70'],
        '097': sheet_form['A74':'D80'],
        '115': sheet_form['A84':'D90'],
        '116': sheet_form['A94':'D100'],
        '138': sheet_form['A104':'D110'],
        '156': sheet_form['A114':'D120'],
        '173': sheet_form['A124':'D130'],
        '193': sheet_form['A134':'D140'],
        '194': sheet_form['A144':'D150'],
        '215': sheet_form['A154':'D160'],
        '234': sheet_form['A164':'D170'],
        '240': sheet_form['A174':'D180'],
        '271': sheet_form['A184':'D190'],
        '289': sheet_form['A194':'D200'],
        '290': sheet_form['A204':'D210'],
        '333': sheet_form['A214':'D220'],
        '337': sheet_form['A224':'D230'],
        '350': sheet_form['A234':'D240'],
        '407': sheet_form['A244':'D250'],
        '414': sheet_form['A254':'D260'],
        '473': sheet_form['A264':'D270'],
        '500': sheet_form['A274':'D280']
    }
    try:
        for key in pp_list.keys():
            if key == vrs_num:
                var_pp = pp_list[key]
                break
        return var_pp
    except UnboundLocalError:
        messagebox.showerror('Error', 'Нет данных в таблице для\nблока клапана вертикального VRS ' + vrs_num)

def promkam_table(vrs_num, vrs_perf):
    if vrs_perf == '01':
        sheet_form = wb_form['Promkam-01']  # переходим на вкладку 01 исполнения
    elif vrs_perf == '02':
        sheet_form = wb_form['Promkam-02']  # переходим на вкладку 02 исполнения
    elif vrs_perf == '03' or vrs_perf == '04':
        sheet_form = wb_form['Promkam-03(04)']  # переходим на вкладку 03 исполнения

    promkam_list = {
        '019': sheet_form['B3':'E3'],
        '034': sheet_form['B4':'E4'],
        '039': sheet_form['B5':'E5'],
        '054': sheet_form['B6':'E6'],
        '058': sheet_form['B7':'E7'],
        '078': sheet_form['B8':'E8'],
        '086': sheet_form['B9':'E9'],
        '097': sheet_form['B10':'E10'],
        '115': sheet_form['B11':'E12'],
        '116': sheet_form['B13':'E13'],
        '138': sheet_form['B14':'E14'],
        '156': sheet_form['B15':'E15'],
        '173': sheet_form['B16':'E16'],
        '193': sheet_form['B17':'E17'],
        '194': sheet_form['B18':'E18'],
        '215': sheet_form['B19':'E19'],
        '234': sheet_form['B20':'E20'],
        '240': sheet_form['B21':'E21'],
        '271': sheet_form['B22':'E22'],
        '289': sheet_form['B23':'E23'],
        '290': sheet_form['B24':'E24'],
        '333': sheet_form['B25':'E25'],
        '337': sheet_form['B26':'E26'],
        '350': sheet_form['B27':'E27'],
        '407': sheet_form['B28':'E28'],
        '414': sheet_form['B29':'E29'],
        '473': sheet_form['B30':'E30'],
        '500': sheet_form['B31':'E31']
    }
    try:
        for key in promkam_list.keys():
            if key == vrs_num:
                var_promkam = promkam_list[key]
                break
        return var_promkam
    except UnboundLocalError:
        messagebox.showerror('Error', 'Нет данных в таблице для\nблока камеры промежуточной VRS ' + vrs_num)