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
        cells = []
        return cells


def vent_table(vrs_num, vrs_perf, vosk):
    if vrs_perf == '01':
        sheet_form = wb_form['Vent-01']  # переходим на вкладку 01 исполнения
    elif vrs_perf == '02':
        sheet_form = wb_form['Vent-02']  # переходим на вкладку 02 исполнения
    elif vrs_perf == '03' or vrs_perf == '04':
        sheet_form = wb_form['Vent-03(04)']  # переходим на вкладку 03 исполнения

    vent_list = {
        '019' + '2.5': sheet_form['A4':'D6'],
        '019' + '2.8': sheet_form['F4':'I6'],
        '019' + '3.15': sheet_form['K4':'N6'],
        '034' + '2.5': sheet_form['A10':'D12'],
        '034' + '2.8': sheet_form['F10':'I12'],
        '034' + '3.15': sheet_form['K10':'N12'],
        '039' + '2.8': sheet_form['A16':'D18'],
        '039' + '3.15': sheet_form['F16':'I18'],
        '039' + '3.55': sheet_form['K16':'N17'],
        '039' + '4.0': sheet_form['P16':'S17'],
        '054' + '3.15': sheet_form['A22':'D24'],
        '054' + '3.55': sheet_form['F22':'I23'],
        '054' + '4.0': sheet_form['K22':'N23'],
        '058' + '3.55': sheet_form['A28':'D29'],
        '058' + '4.0': sheet_form['F28':'I29'],
        '058' + '4.5': sheet_form['K28':'N29'],
        '058' + '5.0': sheet_form['P28':'S29'],
        '078' + '4.0': sheet_form['A33':'D34'],
        '078' + '4.5': sheet_form['F33':'I34'],
        '078' + '5.0': sheet_form['K33':'N34'],
        '078' + '5.6': sheet_form['P33':'S34'],
        '086' + '4.0': sheet_form['A38':'D39'],
        '086' + '4.5': sheet_form['F38':'I39'],
        '086' + '5.0': sheet_form['K38':'N39'],
        '086' + '5.6': sheet_form['P38':'S39'],
        '097' + '4.0': sheet_form['A43':'D44'],
        '097' + '4.5': sheet_form['F43':'I44'],
        '097' + '5.0': sheet_form['K43':'N44'],
        '097' + '5.6': sheet_form['P43':'S44']
    }
    try:
        for key in vent_list.keys():
            if key == vrs_num + vosk:
                vent = vent_list[key]
                break
        return vent
    except UnboundLocalError:
        messagebox.showerror('Error', 'Нет данных в таблице для ВНК блока ВОСК ' + vosk)
        vent = []
        return vent

def vent_table2(vrs_num, vrs_perf, vosk2):
    if vrs_perf == '01':
        sheet_form = wb_form['Vent-01']  # переходим на вкладку 01 исполнения
    elif vrs_perf == '02':
        sheet_form = wb_form['Vent-02']  # переходим на вкладку 02 исполнения
    elif vrs_perf == '03' or vrs_perf == '04':
        sheet_form = wb_form['Vent-03(04)']  # переходим на вкладку 03 исполнения

    vent_list = {
        '019' + '2.5': sheet_form['A4':'D6'],
        '019' + '2.8': sheet_form['F4':'I6'],
        '019' + '3.15': sheet_form['K4':'N6'],
        '034' + '2.5': sheet_form['A10':'D12'],
        '034' + '2.8': sheet_form['F10':'I12'],
        '034' + '3.15': sheet_form['K10':'N12'],
        '039' + '2.8': sheet_form['A16':'D18'],
        '039' + '3.15': sheet_form['F16':'I18'],
        '039' + '3.55': sheet_form['K16':'N17'],
        '039' + '4.0': sheet_form['P16':'S17'],
        '054' + '3.15': sheet_form['A22':'D24'],
        '054' + '3.55': sheet_form['F22':'I23'],
        '054' + '4.0': sheet_form['K22':'N23'],
        '058' + '3.55': sheet_form['A28':'D29'],
        '058' + '4.0': sheet_form['F28':'I29'],
        '058' + '4.5': sheet_form['K28':'N29'],
        '058' + '5.0': sheet_form['P28':'S29'],
        '078' + '4.0': sheet_form['A33':'D34'],
        '078' + '4.5': sheet_form['F33':'I34'],
        '078' + '5.0': sheet_form['K33':'N34'],
        '078' + '5.6': sheet_form['P33':'S34'],
        '086' + '4.0': sheet_form['A38':'D39'],
        '086' + '4.5': sheet_form['F38':'I39'],
        '086' + '5.0': sheet_form['K38':'N39'],
        '086' + '5.6': sheet_form['P38':'S39'],
        '097' + '4.0': sheet_form['A43':'D44'],
        '097' + '4.5': sheet_form['F43':'I44'],
        '097' + '5.0': sheet_form['K43':'N44'],
        '097' + '5.6': sheet_form['P43':'S44']
    }
    try:
        for key in vent_list.keys():
            if key == vrs_num + vosk2:
                vent2 = vent_list[key]
                break
        return vent2
    except UnboundLocalError:
        messagebox.showerror('Error', 'Нет данных в таблице для ВНК блока ВОСК ' + vosk2)
        vent2 = []
        return vent2

def air_table(vrs_perf, air, vosk):
    if vrs_perf == '01':
        sheet_form = wb_form['Vosk-01']  # переходим на вкладку 01 исполнения
    elif vrs_perf == '02':
        sheet_form = wb_form['Vosk-02']  # переходим на вкладку 02 исполнения
    elif vrs_perf == '03' or vrs_perf == '04':
        sheet_form = wb_form['Vosk-03(04)']  # переходим на вкладку 03 исполнения

    air_list = {
        '2.5' + 'АИР56': sheet_form['A4':'D9'],
        '2.5' + 'АИР63': sheet_form['F4':'I9'],
        '2.5' + 'АИР71': sheet_form['K4':'N9'],
        '2.5' + 'АИР80': sheet_form['P4':'S9'],
        '2.8' + 'АИР56': sheet_form['A13':'D18'],
        '2.8' + 'АИР63': sheet_form['F13':'I18'],
        '2.8' + 'АИР71': sheet_form['K13':'N18'],
        '2.8' + 'АИР80': sheet_form['P13':'S18'],
        '2.8' + 'АИР90': sheet_form['U13':'X18'],
        '3.15' + 'АИР56': sheet_form['A22':'D27'],
        '3.15' + 'АИР63': sheet_form['F22':'I27'],
        '3.15' + 'АИР71': sheet_form['K22':'N27'],
        '3.15' + 'АИР80': sheet_form['P22':'S27'],
        '3.15' + 'АИР90': sheet_form['U22':'X27'],
        '3.15' + 'АИР100': sheet_form['Z22':'AC27'],
        '3.55' + 'АИР56': sheet_form['A31':'D36'],
        '3.55' + 'АИР63': sheet_form['F31':'I36'],
        '3.55' + 'АИР71': sheet_form['K31':'N36'],
        '3.55' + 'АИР80': sheet_form['P31':'S36'],
        '3.55' + 'АИР90': sheet_form['U31':'X36'],
        '3.55' + 'АИР100': sheet_form['Z31':'AC36'],
        '4.0' + 'АИР63': sheet_form['A40':'D46'],
        '4.0' + 'АИР71': sheet_form['F40':'I46'],
        '4.0' + 'АИР80': sheet_form['K40':'N46'],
        '4.0' + 'АИР90': sheet_form['P40':'S46'],
        '4.0' + 'АИР100': sheet_form['U40':'X46'],
        '4.0' + 'АИР112': sheet_form['Z40':'AC46'],
        '4.0' + 'АИР132': sheet_form['AE40':'AH46'],
        '4.5' + 'АИР71': sheet_form['A50':'D56'],
        '4.5' + 'АИР80': sheet_form['F50':'I56'],
        '4.5' + 'АИР90': sheet_form['K50':'N56'],
        '4.5' + 'АИР100': sheet_form['P50':'S56'],
        '4.5' + 'АИР112': sheet_form['U50':'X56'],
        '4.5' + 'АИР132': sheet_form['Z50':'AC56'],
        '4.5' + 'АИР160': sheet_form['AE50':'AH56'],
        '5.0' + 'АИР71': sheet_form['A60':'D66'],
        '5.0' + 'АИР80': sheet_form['F60':'I66'],
        '5.0' + 'АИР90': sheet_form['K60':'N66'],
        '5.0' + 'АИР100': sheet_form['P60':'S66'],
        '5.0' + 'АИР112': sheet_form['U60':'X66'],
        '5.0' + 'АИР132': sheet_form['Z60':'AC66'],
        '5.0' + 'АИР160': sheet_form['AE60':'AH66'],
        '5.0' + 'АИР180': sheet_form['AJ60':'AM66'],
        '5.6' + 'АИР71': sheet_form['A70':'D76'],
        '5.6' + 'АИР80': sheet_form['F70':'I76'],
        '5.6' + 'АИР90': sheet_form['K70':'N76'],
        '5.6' + 'АИР100': sheet_form['P70':'S76'],
        '5.6' + 'АИР112': sheet_form['U70':'X76'],
        '5.6' + 'АИР132': sheet_form['Z70':'AC76'],
        '5.6' + 'АИР160': sheet_form['AE70':'AH76'],
        '5.6' + 'АИР180': sheet_form['AJ70':'AM76']
    }
    try:
        for key in air_list.keys():
            if key == vosk + air:
               var_air = air_list[key]
        return var_air
    except UnboundLocalError:
        messagebox.showerror('Error', 'Нет данных в таблице для вентагрегата ВОСК ' + vosk + ' ' + air)
        var_air = []
        return var_air

def air_table2(vrs_perf, air2, vosk2):
    if vrs_perf == '01':
        sheet_form = wb_form['Vosk-01']  # переходим на вкладку 01 исполнения
    elif vrs_perf == '02':
        sheet_form = wb_form['Vosk-02']  # переходим на вкладку 02 исполнения
    elif vrs_perf == '03' or vrs_perf == '04':
        sheet_form = wb_form['Vosk-03(04)']  # переходим на вкладку 03 исполнения

    air_list = {
        '2.5' + 'АИР56': sheet_form['A4':'D9'],
        '2.5' + 'АИР63': sheet_form['F4':'I9'],
        '2.5' + 'АИР71': sheet_form['K4':'N9'],
        '2.5' + 'АИР80': sheet_form['P4':'S9'],
        '2.8' + 'АИР56': sheet_form['A13':'D18'],
        '2.8' + 'АИР63': sheet_form['F13':'I18'],
        '2.8' + 'АИР71': sheet_form['K13':'N18'],
        '2.8' + 'АИР80': sheet_form['P13':'S18'],
        '2.8' + 'АИР90': sheet_form['U13':'X18'],
        '3.15' + 'АИР56': sheet_form['A22':'D27'],
        '3.15' + 'АИР63': sheet_form['F22':'I27'],
        '3.15' + 'АИР71': sheet_form['K22':'N27'],
        '3.15' + 'АИР80': sheet_form['P22':'S27'],
        '3.15' + 'АИР90': sheet_form['U22':'X27'],
        '3.15' + 'АИР100': sheet_form['Z22':'AC27'],
        '3.55' + 'АИР56': sheet_form['A31':'D36'],
        '3.55' + 'АИР63': sheet_form['F31':'I36'],
        '3.55' + 'АИР71': sheet_form['K31':'N36'],
        '3.55' + 'АИР80': sheet_form['P31':'S36'],
        '3.55' + 'АИР90': sheet_form['U31':'X36'],
        '3.55' + 'АИР100': sheet_form['Z31':'AC36'],
        '4.0' + 'АИР63': sheet_form['A40':'D46'],
        '4.0' + 'АИР71': sheet_form['F40':'I46'],
        '4.0' + 'АИР80': sheet_form['K40':'N46'],
        '4.0' + 'АИР90': sheet_form['P40':'S46'],
        '4.0' + 'АИР100': sheet_form['U40':'X46'],
        '4.0' + 'АИР112': sheet_form['Z40':'AC46'],
        '4.0' + 'АИР132': sheet_form['AE40':'AH46'],
        '4.5' + 'АИР71': sheet_form['A50':'D56'],
        '4.5' + 'АИР80': sheet_form['F50':'I56'],
        '4.5' + 'АИР90': sheet_form['K50':'N56'],
        '4.5' + 'АИР100': sheet_form['P50':'S56'],
        '4.5' + 'АИР112': sheet_form['U50':'X56'],
        '4.5' + 'АИР132': sheet_form['Z50':'AC56'],
        '4.5' + 'АИР160': sheet_form['AE50':'AH56'],
        '5.0' + 'АИР71': sheet_form['A60':'D66'],
        '5.0' + 'АИР80': sheet_form['F60':'I66'],
        '5.0' + 'АИР90': sheet_form['K60':'N66'],
        '5.0' + 'АИР100': sheet_form['P60':'S66'],
        '5.0' + 'АИР112': sheet_form['U60':'X66'],
        '5.0' + 'АИР132': sheet_form['Z60':'AC66'],
        '5.0' + 'АИР160': sheet_form['AE60':'AH66'],
        '5.0' + 'АИР180': sheet_form['AJ60':'AM66'],
        '5.6' + 'АИР71': sheet_form['A70':'D76'],
        '5.6' + 'АИР80': sheet_form['F70':'I76'],
        '5.6' + 'АИР90': sheet_form['K70':'N76'],
        '5.6' + 'АИР100': sheet_form['P70':'S76'],
        '5.6' + 'АИР112': sheet_form['U70':'X76'],
        '5.6' + 'АИР132': sheet_form['Z70':'AC76'],
        '5.6' + 'АИР160': sheet_form['AE70':'AH76'],
        '5.6' + 'АИР180': sheet_form['AJ70':'AM76']
    }
    try:
        for key in air_list.keys():
            if key == vosk2 + air2:
                var_air2 = air_list[key]
                break
        return var_air2
    except UnboundLocalError:
        messagebox.showerror('Error', 'Нет данных в таблице для вентагрегата ВОСК ' + vosk2 + ' ' + air2)
        var_air2 = []
        return var_air2

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
        var_vnv5012 = []
        return var_vnv5012

def vov5012_table(vrs_num, vrs_perf, vov5012):
    if vrs_perf == '01':
        sheet_form = wb_form['Vov5012-01']  # переходим на вкладку 01 исполнения
    elif vrs_perf == '02':
        sheet_form = wb_form['Vov5012-02']  # переходим на вкладку 02 исполнения
    elif vrs_perf == '03' or vrs_perf == '04':
        sheet_form = wb_form['Vov5012-03(04)']  # переходим на вкладку 03 исполнения

    vov5012_list = {
        '019' + '180': sheet_form['A4':'D13'], '019' + '220': sheet_form['F4':'I13'],
        '019' + '280': sheet_form['K4':'N13'], '019' + '310': sheet_form['P4':'S13'],
        '034' + '180': sheet_form['A17':'D26'], '034' + '220': sheet_form['F4':'I13'],
        '034' + '280': sheet_form['K17':'N26'], '034' + '310': sheet_form['P4':'S13'],
        '039' + '180': sheet_form['A30':'D39'], '039' + '220': sheet_form['F30':'I39'],
        '039' + '280': sheet_form['K30':'N39'], '039' + '310': sheet_form['P30':'S39'],
        '054' + '180': sheet_form['A43':'D52'], '054' + '220': sheet_form['F43':'I52'],
        '054' + '280': sheet_form['K43':'N52'], '054' + '310': sheet_form['P43':'S52'],
        '058' + '180': sheet_form['A56':'D65'], '058' + '220': sheet_form['F56':'I65'],
        '058' + '280': sheet_form['K56':'N65'], '058' + '310': sheet_form['P56':'S65'],
        '078' + '180': sheet_form['A69':'D79'], '078' + '220': sheet_form['F69':'I79'],
        '078' + '280': sheet_form['K69':'N79'], '078' + '310': sheet_form['P69':'S79'],
        '086' + '180': sheet_form['A83':'D92'], '086' + '220': sheet_form['F83':'I92'],
        '086' + '280': sheet_form['K83':'N92'], '086' + '310': sheet_form['P83':'S92'],
        '097' + '180': sheet_form['A96':'D106'], '097' + '220': sheet_form['F96':'I106'],
        '097' + '280': sheet_form['K96':'N106'], '097' + '310': sheet_form['P96':'S106'],
        '115' + '180': sheet_form['A110':'D120'], '115' + '220': sheet_form['F110':'I120'],
        '115' + '280': sheet_form['K110':'N120'], '115' + '310': sheet_form['P110':'S120'],
        '116' + '180': sheet_form['A124':'D134'], '116' + '220': sheet_form['F124':'I134'],
        '116' + '280': sheet_form['K124':'N134'], '116' + '310': sheet_form['P124':'S134'],
        '138' + '180': sheet_form['A138':'D148'], '138' + '220': sheet_form['F138':'I148'],
        '138' + '280': sheet_form['K138':'N148'], '138' + '310': sheet_form['P138':'S148'],
        '156' + '180': sheet_form['A152':'D163'], '156' + '220': sheet_form['F152':'I163'],
        '156' + '280': sheet_form['K152':'N163'], '156' + '310': sheet_form['P152':'S163'],
        '173' + '180': sheet_form['A167':'D176'], '173' + '220': sheet_form['F167':'I176'],
        '173' + '280': sheet_form['K167':'N176'], '173' + '310': sheet_form['P167':'S176'],
        '193' + '180': sheet_form['A180':'D191'], '193' + '220': sheet_form['F180':'I913'],
        '193' + '280': sheet_form['K180':'N191'], '193' + '310': sheet_form['P180':'S191'],
        '194' + '180': sheet_form['A195':'D206'], '194' + '220': sheet_form['F195':'I206'],
        '194' + '280': sheet_form['K195':'N206'], '194' + '310': sheet_form['P195':'S206'],
        '215' + '180': sheet_form['A210':'D221'], '215' + '220': sheet_form['F210':'I221'],
        '215' + '280': sheet_form['K210':'N221'], '215' + '310': sheet_form['P210':'S221'],
        '234' + '180': sheet_form['A225':'D236'], '234' + '220': sheet_form['F225':'I236'],
        '234' + '280': sheet_form['K225':'N236'], '234' + '310': sheet_form['P225':'S236'],
        '240' + '180': sheet_form['A240':'D251'], '240' + '220': sheet_form['F240':'I251'],
        '240' + '280': sheet_form['K240':'N251'], '240' + '310': sheet_form['P240':'S251'],
        '271' + '180': sheet_form['A255':'D266'], '271' + '220': sheet_form['F255':'I266'],
        '271' + '280': sheet_form['K255':'N266'], '271' + '310': sheet_form['P255':'S266'],
        '289' + '180': sheet_form['A270':'D281'], '289' + '220': sheet_form['F270':'I281'],
        '289' + '280': sheet_form['K270':'N281'], '289' + '310': sheet_form['P270':'S281'],
        '290' + '180': sheet_form['A285':'D296'], '290' + '220': sheet_form['F285':'I296'],
        '290' + '280': sheet_form['K285':'N296'], '290' + '310': sheet_form['P285':'S296'],
        '333' + '180': sheet_form['A300':'D311'], '333' + '220': sheet_form['F300':'I311'],
        '333' + '280': sheet_form['K300':'N311'], '333' + '310': sheet_form['P300':'S311'],
        '337' + '180': sheet_form['A315':'D326'], '337' + '220': sheet_form['F315':'I326'],
        '337' + '280': sheet_form['K315':'N326'], '337' + '310': sheet_form['P315':'S326'],
        '350' + '180': sheet_form['A330':'D341'], '350' + '220': sheet_form['F330':'I341'],
        '350' + '280': sheet_form['K330':'N341'], '350' + '310': sheet_form['P330':'S341'],
        '407' + '180': sheet_form['A345':'D356'], '407' + '220': sheet_form['F345':'I356'],
        '407' + '280': sheet_form['K345':'N356'], '407' + '310': sheet_form['P345':'S356'],
        '414' + '180': sheet_form['A360':'D371'], '414' + '220': sheet_form['F360':'I371'],
        '414' + '280': sheet_form['K360':'N371'], '414' + '310': sheet_form['P360':'S371'],
        '473' + '180': sheet_form['A375':'D386'], '473' + '220': sheet_form['F375':'I386'],
        '473' + '280': sheet_form['K375':'N386'], '473' + '310': sheet_form['P375':'S386'],
        '500' + '180': sheet_form['A390':'D401'], '500' + '220': sheet_form['F390':'I401'],
        '500' + '280': sheet_form['K390':'N401'], '500' + '310': sheet_form['P390':'S401']
    }
    try:
        for key in vov5012_list.keys():
            if key == vrs_num + vov5012:
                var_vov5012 = vov5012_list[key]
        return var_vov5012
    except UnboundLocalError:
        messagebox.showerror('Error', 'Нет данных в таблице для\nблока ВОВ 5012 VRS ' + vrs_num)
        var_vov5012 = []
        return var_vov5012

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
        var_eko = []
        return var_eko

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
        var_vertklap = []
        return var_vertklap

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
        '138': sheet_form['A94':'D100'],
        '156': sheet_form['A104':'D111'],
        '173': sheet_form['A115':'D121'],
        '193': sheet_form['A125':'D132'],
        '215': sheet_form['A136':'D142'],
        '234': sheet_form['A146':'D153'],
        '240': sheet_form['A157':'D164'],
        '271': sheet_form['A168':'D175'],
        '289': sheet_form['A179':'D186'],
        '333': sheet_form['A190':'D197'],
        '350': sheet_form['A201':'D208'],
        '473': sheet_form['A212':'D219']
    }
    try:
        for key in pp_list.keys():
            if key == vrs_num:
                var_pp = pp_list[key]
                break
        return var_pp
    except UnboundLocalError:
        messagebox.showerror('Error', 'Нет данных в таблице для\nблока клапана вертикального VRS ' + vrs_num)
        var_pp = []
        return var_pp

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
        var_promkam = []
        return var_promkam