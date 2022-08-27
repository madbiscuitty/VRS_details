from openpyxl import *
from tkinter import *
from tkinter import messagebox
from tkinter.ttk import Combobox
from table_var import *

wb = Workbook()
ws = wb.active
#wb.worksheets[sht].clear()
ws.title = 'Project'

class VRS:
    def __init__(self, width, height, title="MyWindow", resizable=(False, False), icon=r"resourses/icon.ico"):
        self.root = Tk()
        self.root.title(title)
        self.root.resizable(resizable[0], resizable[1])
        if icon:
            self.root.iconbitmap(icon)
        #объекты виджетов окна
        self.vrs_num = Combobox(self.root, values=("019", "034", "039", "054", "058", "078", "086", "097", "115",
                                                   "116", "138", "156", "173", "193", "194", "215", "234", "240",
                                                   "271", "289", "290", "333", "337", "350", "407", "414", "473",
                                                   "500"), width=7, state="readonly")
        self.vrs_perf = Combobox(self.root, values=("01", "02", "03", "04"), width=7, state="readonly")
        self.air = Combobox(self.root, values=("АИР56", "АИР63", "АИР71", "АИР80", "АИР90", "АИР100", "АИР112",
                                               "АИР132", "АИР160", "АИР180", "АИР225", "АИР250"), width=7, state="readonly")
        self.air2 = Combobox(self.root, values=("АИР56", "АИР63", "АИР71", "АИР80", "АИР90", "АИР100", "АИР112",
                                                "АИР132", "АИР160", "АИР180", "АИР225", "АИР250"), width=7, state="readonly")
        self.vosk = Combobox(self.root, values=("2.5", "2.8", "3.15", "3.55", "4.0", "4.5", "5.0", "5.6", "6.3", "7.1",
                                                "8.0", "9.0", "10.0", "11.2", "12.5"), width=7, state="readonly")
        self.vosk2 = Combobox(self.root, values=("2.5", "2.8", "3.15", "3.55", "4.0", "4.5", "5.0", "5.6", "6.3", "7.1",
                                                "8.0", "9.0", "10.0", "11.2", "12.5"), width=7, state="readonly")
        self.vnv5012_widht = Combobox(self.root, values=("160", "180"), width=7, state="readonly")
       # self.vnv4816_widht = Combobox(self.root, values=("240", "320", "400"), width=7, state="readonly")
        self.vov5012_widht = Combobox(self.root, values=("180", "220", "260", "310"), width=7, state="readonly")
       # self.vov4816_widht = Combobox(self.root, values=("240", "320", "400"), width=7, state="readonly")
        self.filterbox = IntVar()   # заводим int переменную
        self.ventbox = IntVar()
        self.vent2box = IntVar()
        self.vnv5012box = IntVar()
        #self.vnv4816box = IntVar()
        self.vov5012box = IntVar()
        #self.vov4816box = IntVar()
        self.ekobox = IntVar()
        self.vertklapbox = IntVar()
        self.ppbox = IntVar()
        self.promkambox = IntVar()
        count = (1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
        self.entryfilter = Spinbox(self.root, values=(count), bd=2)  # заводим спинбокс
        self.entryvent = Spinbox(self.root, values=(count), width=7, bd=2)
        self.entryvent2 = Spinbox(self.root, values=(count), width=7, bd=2)
        self.entryvnv5012 = Spinbox(self.root, values=(count), width=7, bd=2)
       # self.entryvnv4816 = Spinbox(self.root, values=(count), width=7, bd=2)
        self.entryvov5012 = Spinbox(self.root, values=(count), width=7, bd=2)
       # self.entryvov4816 = Spinbox(self.root, values=(count), width=7, bd=2)
        self.entryeko = Spinbox(self.root, values=(count), bd=2)
        self.entryvertklap = Spinbox(self.root, values=(count), bd=2)
        self.entrypp = Spinbox(self.root, values=(count), bd=2)
        self.entrypromkam = Spinbox(self.root, values=(count), bd=2)

    #Запуск окна программы
    def run(self):
        self.draw_widgets()
        self.root.mainloop()
    #Построение окна программы
    def draw_widgets(self):
        Label(self.root, text="Выберите типоразмер VRS:", justify=LEFT, font=("", 9, "bold")).grid(row=0, column=0,
                                                                                                    sticky=W)
        self.vrs_num.grid(row=0, column=1, columnspan=3, sticky=W + E, padx=5, pady=5)
        self.vrs_num.current(0)                                                                                                     # устанавливаем дефолтную позицию выпадающего меню
        Label(self.root, text="Выберите исполнение VRS:", justify=LEFT, font=("", 10, "bold")).grid(row=1, column=0, sticky=W)
        self.vrs_perf.grid(row=1, column=1, columnspan=3, sticky=W + E, padx=5, pady=5)       # вывод выпадающего меню
        self.vrs_perf.current(0)                                                                                                    # устанавливаем дефолтную позицию выпадающего меню
        Label(self.root, text="Выберите параметры блоков:", justify=LEFT, font=("", 12, "bold")).grid(
            row=2, column=0, sticky=W) # вывод строки с текстом
        Button(self.root, width=2, height=1, text="?", command=self.info).grid(row=2, column=1, sticky=W)
        Checkbutton(self.root, text="Блок фильтра", justify=LEFT, font=("", 9, "bold"), variable=self.filterbox).grid(
            row=3, column=0,
                                                                                                 sticky=W)          # строка чекбокса
        self.entryfilter.grid(row=3, column=1, columnspan=3, sticky=W+E, padx=5, pady=5)  # рисуем спинбокс
        Checkbutton(self.root, text="Блок вентилятора ВОСК", font=("", 9, "bold"), justify=LEFT,
                    variable=self.ventbox).grid(row=4,
                                                                                                       column=0, sticky=W)   # строка чекбокса
        self.vosk.grid(row=4, column=1, sticky=W, padx=5, pady=5)
        self.vosk.current(0)
        self.air.grid(row=4, column=2, sticky=W, padx=5, pady=5)           # выпадающая менюшка аиров
        self.air.current(0)  # устанавливаем дефолтную позицию выпадающего меню
        self.entryvent.grid(row=4, column=3, sticky=W, padx=5, pady=5)
        Checkbutton(self.root, text="Блок вентилятора ВОСК 2", justify=LEFT, font=("", 9, "bold"), variable=self.vent2box).grid(row=5,
                                                                                                          column=0, sticky=W)  # строка чекбокса
        self.vosk2.grid(row=5, column=1, sticky=W, padx=5, pady=5)
        self.vosk2.current(0)
        self.air2.grid(row=5, column=2, sticky=W, padx=5, pady=5)  # выпадающая менюшка аиров
        self.air2.current(0)  # устанавливаем дефолтную позицию выпадающего меню
        self.entryvent2.grid(row=5, column=3, sticky=W, padx=5, pady=5)
        Checkbutton(self.root, text="Блок ВНВ 5012", justify=LEFT, font=("", 9, "bold"), variable=self.vnv5012box).grid(row=6, column=0,
                                                                                                  sticky=W)        # строка чекбокса
        self.vnv5012_widht.grid(row=6, column=1, sticky=W, padx=5, pady=5)  # выпадающая менюшка внв
        self.vnv5012_widht.current(0)  # устанавливаем дефолтную позицию выпадающего меню
        self.entryvnv5012.grid(row=6, column=2, columnspan=3, sticky=W+E, padx=5, pady=5)
        # Checkbutton(self.root, text="Блок ВНВ 4816", justify=LEFT, variable=self.vnv4816box).grid(row=6, column=0, sticky=W)      # строка чекбокса
        # self.vnv4816_widht.grid(row=6, column=1, sticky=W)  # выпадающая менюшка внв
        # self.vnv4816_widht.current(0)  # устанавливаем дефолтную позицию выпадающего меню
        # self.entryvnv4816.grid(row=6, column=2, columnspan=2, sticky=W)
        Checkbutton(self.root, text="Блок ВОВ 5012", justify=LEFT, font=("", 9, "bold"), variable=self.vov5012box).grid(row=7, column=0,
                                                                                                  sticky=W)     # строка чекбокса
        self.vov5012_widht.grid(row=7, column=1, sticky=W, padx=5, pady=5)  # выпадающая менюшка внв
        self.vov5012_widht.current(0)  # устанавливаем дефолтную позицию выпадающего меню
        self.entryvov5012.grid(row=7, column=2, columnspan=3, sticky=W+E, padx=5, pady=5)
        # Checkbutton(self.root, text="Блок ВОВ 4816", justify=LEFT, variable=self.vov4816box).grid(row=8, column=0, sticky=W)       # строка чекбокса
        # self.vov4816_widht.grid(row=8, column=1, sticky=W)  # выпадающая менюшка внв
        # self.vov4816_widht.current(0)  # устанавливаем дефолтную позицию выпадающего меню
        # self.entryvov4816.grid(row=8, column=2, columnspan=2, sticky=W)
        Checkbutton(self.root, text="Блок ЭКО", justify=LEFT, font=("", 9, "bold"), variable=self.ekobox).grid(row=9, column=0,
                                                                                          sticky=W)  # строка чекбокса
        self.entryeko.grid(row=9, column=1, columnspan=3, sticky=W+E, padx=5, pady=5)
        Checkbutton(self.root, text="Блок вертикального клапана", justify=LEFT, font=("", 9, "bold"), variable=self.vertklapbox).grid(
            row=10, column=0, sticky=W)  # строка чекбокса
        self.entryvertklap.grid(row=10, column=1, columnspan=3, sticky=W+E, padx=5, pady=5)
        Checkbutton(self.root, text="Блок пластинчатого утилизатора", justify=LEFT, font=("", 9, "bold"), variable=self.ppbox).grid(
            row=11, column=0, sticky=W)  # строка чекбокса
        self.entrypp.grid(row=11, column=1, columnspan=3, sticky=W+E, padx=5, pady=5)
        Checkbutton(self.root, text="Блок камеры промежуточной",justify=LEFT, font=("", 9, "bold"),
                    variable=self.promkambox).grid(row=12,
                                                                                                            column=0, sticky=W)  # строка чекбокса
        self.entrypromkam.grid(row=12, column=1, columnspan=3, sticky=W+E, padx=5, pady=5) # рисуем спинбокс

        Button(self.root, width=6, height=2, text="Выгрузить", font=("Ghost type A", 9, "bold"),
               command=self.action).grid(row=13, column=0, sticky=W+E)                        # Кнопка выгрузки
        Button(self.root, width=6, height=2, text="Закрыть", command=self.root.destroy).grid(row=13, column=1,
                                                                                             columnspan=3,
                                                                                             sticky=W+E)  # Кнопка закрытия
        Label(self.root, text="v0.0.5 alfa", justify=LEFT).grid(row=14, column=3, sticky=E)


    #Конопка INFO
    def info(self, title="INFO", resizable=(False, False), icon=r"resourses/info.ico"):
        draw = Tk()
        draw.title(title)
        draw.resizable(resizable[0], resizable[1])
        if icon:
            draw.iconbitmap(icon)

        Label(draw, text="- Для блоков ВОСК выбираем типоразмер ВОСК, типоразмер двигателя и "
                                             "количество.\n- Для блока ВНВ и ВОВ "
                                             "выбираем ширину ТО и количество.\n- Для всего остального на выбор только "
                                             "количество.", justify=LEFT, padx=15, pady=15, font=("", 10, "bold")).grid(
                                                                                              row=0, column=0, sticky=W)
        Button(draw, width=10, height=2, text="Ок", command=draw.destroy).grid(row=1, column=0)
    #вывод данных в excel
    def action(self):
        global sheet_form
        vrs_num = self.vrs_num.get()
        vrs_perf = self.vrs_perf.get()
        vosk = self.vosk.get()
        vosk2 = self.vosk2.get()
        air = self.air.get()
        air2 = self.air2.get()
        vnv = self.vnv5012_widht.get()
        vov = self.vov5012_widht.get()
        #форматирование заголовка в выводимой таблице
        ws.cell(row=1, column=1).value = 'Перечень VRS-500-' + vrs_num + '-' + vrs_perf
        ws.cell(row=2, column=1).value = 'Обозначение'
        ws.cell(row=2, column=2).value = 'Наименование'
        ws.cell(row=2, column=3).value = 'Кол-во'
        ws.cell(row=2, column=4).value = 'Материал'
        ws.merge_cells('A1:D1')
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 35
        ws.column_dimensions['C'].width = 7
        ws.column_dimensions['D'].width = 12
        if self.filterbox.get():
            filteramount = int(self.entryfilter.get())
            ws.append(['Блок фильтра', '', filteramount])

            cells = fb_table(vrs_num, vrs_perf)
            for sign, name, amount, material in cells:
                #print(sign.value, name.value, amount.value, material.value) # выводим нужные строки
                ws.append([sign.value, name.value, int(amount.value*filteramount), material.value])

        if self.ventbox.get():
            ventamount = int(self.entryvent.get())
            ws.append(['Блок вентилятора ВОСК ' + self.vosk.get() + ' ' + self.air.get(), '', ventamount])
            #ws.append(['Разблюдовка:'])

            vent = vent_table(vrs_num, vrs_perf, vosk)
            air = air_table(vrs_perf, air)
            for sign, name, amount, material in vent:
                #print(sign.value, name.value, amount.value, material.value) #проверяем в консоли на вывод нужных строк
                ws.append([sign.value, name.value, int(amount.value * ventamount), material.value])
            for sign, name, amount, material in air:
                #print(sign.value, name.value, amount.value, material.value) #проверяем в консоли на вывод нужных строк
                ws.append([sign.value, name.value, int(amount.value * ventamount), material.value])

        if self.vent2box.get():
            ventamount = int(self.entryvent2.get())
            ws.append(['Блок вентилятора ВОСК ' + self.vosk2.get() + ' ' + self.air2.get(), '', ventamount])
            #ws.append(['Разблюдовка:'])

            vent2 = vent2_table(vrs_num, vrs_perf, vosk2)
            air2 = air2_table(vrs_perf, air2)
            for sign, name, amount, material in vent2:
                #print(sign.value, name.value, amount.value, material.value) #проверяем в консоли на вывод нужных строк
                ws.append([sign.value, name.value, int(amount.value * ventamount), material.value])
            for sign, name, amount, material in air2:
                #print(sign.value, name.value, amount.value, material.value) #проверяем в консоли на вывод нужных строк
                ws.append([sign.value, name.value, int(amount.value * ventamount), material.value])


        if self.filterbox.get() == 0 and self.ventbox.get() == 0 and self.vent2box.get() == 0 and \
                self.vnv5012box.get()  == 0 and self.vov5012box.get() == 0 and self.ekobox.get() == 0 and \
                self.vertklapbox.get() == 0 and self.ppbox.get() == 0 and self.promkambox.get() == 0:
            messagebox.showerror("VRS", "Выберите хотя бы 1 блок")

        if self.filterbox.get() == 1 or self.ventbox.get() == 1 or self.vent2box.get() == 1 or  self.vnv5012box.get()\
                == 1 or self.vov5012box.get() == 1 or self.ekobox.get() == 1 or self.vertklapbox.get() == 1 or  \
                self.ppbox.get() == 1 or self.promkambox.get() == 1:
            wb.save('Перечень VRS-500-' + vrs_num + '-' + vrs_perf + '.xlsx')
            messagebox.showinfo("VRS", 'Перечень VRS-500-' + vrs_num + '-' + vrs_perf + '\nвыгружен в корень папки')

if __name__ == "__main__":
    program = VRS(460, 200, "Перечень деталей")
    program.run()