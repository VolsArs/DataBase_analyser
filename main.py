import operator
import sys
from tkinter import *
import tkinter.filedialog as fd
from tkinter import scrolledtext
import pyodbc
import datetime
import webbrowser
from time import *

root = Tk()


def choosing_language_params_2(select: IntVar) -> list:
   # print('selet val' + str(select.get()))
    list_of_params_2_ru = ['АЛАРМ', 'АКТ/НЕПДТВ', '']
    list_of_params_2_eng = ['ALARM', 'ACT/UNACK', '']
    if select.get() == 1:
        return list_of_params_2_ru
    if select.get() == 0:
        return list_of_params_2_eng


def choosing_language_params_3(select: IntVar) -> list:
    alarm_prioreties_3_ru = ['15-КРИТИЧЕСКИЙ', '11-ПРЕДУПРЕДИТЕЛЬН', '07-СИГНАЛЬНЫЙ']
    alarm_prioreties_3_eng = ['15-CRITICAL', '11-WARNING', '07-ADVISORY']
    if select.get() == 1:
        return alarm_prioreties_3_ru
    else:
        return alarm_prioreties_3_eng


def setWindow(root):
    root.title('Анализ журнала алармов и событий')
    root.resizable(False, False)

    w = 820
    h = 700
    ws = root.winfo_screenwidth()
    wh = root.winfo_screenheight()
    x = int(ws / 2 - w / 2)
    y = int(wh / 2 - h / 2)
    root.geometry("{0}x{1}+{2}+{3}".format(w, h, x, y))


setWindow(root)


def choose_file() -> str:
    global path_text
    filetypes = ("База данных", "*.mdb"),

    filename = fd.askopenfilename(title="Открыть файл", initialdir="/",
                                  filetypes=filetypes)
    if filename:
        # path_text.configure(text='')
        path_text.delete(0, 30)
        path_text.insert(0, filename)
        print(filename)


def connection(path: str) -> pyodbc:
    global text
    try:
        dr = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        con_string = dr + \
                     r'DBQ=' + path
        conn = pyodbc.connect(con_string)
        return conn

        # print(row)
    except pyodbc.Error as e:
        print("Error in Connection")
        text.insert(END, 'Ошибка подключения к базе данных, убедитесь что файл выбран верно.' + '\n')


def query(conn: pyodbc, list_of_params: list, type_of_query: int) -> list:
    curs = conn.cursor()
    returned_list = []
    if type_of_query == 2:
        sql_alarms = """
            SELECT [Module] FROM JOURNAL WHERE [Event Type] LIKE ? AND [State] LIKE ? 
            """
        params = (list_of_params[0], list_of_params[1])
    if type_of_query == 3:
        sql_alarms = """
                SELECT [Module] FROM JOURNAL WHERE [Event Type] LIKE ? AND [State] LIKE ? AND [Level] LIKE ?
                """
        params = (list_of_params[0], list_of_params[1], list_of_params[2])

    curs.execute(sql_alarms, params)
    for row in curs.fetchall():
        returned_list.insert(0, row[0])
    return returned_list


def find_alarms():  # общее представление об алармах в базе
    alarms = []
    critical = []
    predupred = []
    signaln = []

    alarms = query(connection(path_text.get()), choosing_language_params_2(choice_lang), 2)

    for i in range(0, 3):
        if i == 0:
            params = choosing_language_params_2(choice_lang)
            params[2] = choosing_language_params_3(choice_lang)[0]
            critical = query(connection(path_text.get()), params, 3)
        if i == 1:
            params1 = choosing_language_params_2(choice_lang)
            params1[2] = choosing_language_params_3(choice_lang)[1]
            predupred = query(connection(path_text.get()), params1, 3)
        if i == 2:
            params2 = choosing_language_params_2(choice_lang)
            params2[2] = choosing_language_params_3(choice_lang)[2]
            signaln = query(connection(path_text.get()), params2, 3)

    hour = time_difference(connection(path_text.get()))
    alarm_per_hour = (len(alarms) / int(hour))
    print(alarm_per_hour)
    print(str(len(alarms)) +' список алармов')
    alarm_per_minute = (len(alarms) / (int(hour) * 60)) // 1
    text.insert(END, 'Общее колличество алармов ' + str(len(alarms)) + '\n')
    text.insert(END, 'Колличество критических алармов ' + str(len(critical)) + '\n')
    text.insert(END, 'Колличество предупредительных алармов ' + str(len(predupred)) + '\n')
    text.insert(END, 'Колличество сигнальных алармов ' + str(len(signaln)) + '\n')
    text.insert(END, ''  '\n')
    text.insert(END, 'Исследуемый период времени ' + str(hour) + ' часов' + '\n')
    text.insert(END, 'Колличество алармов в час ' + str(alarm_per_hour) + '\n')
    text.insert(END, 'Колличество алармов в минуту ' + str(alarm_per_minute) + '\n')
    text.insert(END,
                'Согласно ISA 18.2 - Управление  системами сигнализации для обрабатывающих отраслей промышленности' + '\n'
                + ' низкий допустимы уровнеь <6  и высокий допустимый < 12  в час')


def time_difference(conn: pyodbc) -> float:
    curs = conn.cursor()
    date_time = []
    sql_time = """
           SELECT [Date/Time] FROM JOURNAL 
           """
    params = ()
    curs.execute(sql_time, params)
    for row in curs.fetchall():
        date_time.insert(0, row[0])
        # text.insert(END,row[0]+'\n') #
    # date_time.reverse()
    print(date_time[0])
    print(date_time[len(date_time) - 1])
    time_start = date_time[0]
    time_finish = date_time[len(date_time) - 1]
    time_dif = time_start - time_finish
    print(time_dif)
    time_dif_hours = (time_dif.days * 24 + time_dif.seconds / 3600) // 1
    print(time_dif_hours)

    return time_dif_hours


def counting_alarms_of_modules():
    sorted_dict = define_modules()
    index = 0
    for key, value in sorted_dict.items():
        index += 1
        text.insert(END, str(index) + '. Модуль ' + str(key) + ' сгенерировал ' + str(value) + ' алармов' + '\n')
    printing_module_names(sorted_dict)


def define_modules() -> dict:
    alarms_dict = dict()
    params = choosing_language_params_2(choice_lang)
    print(params)
    alarms = query(connection(path_text.get()), params, 2)

    alarms_set = set(alarms)

    for item in alarms_set:
        count = alarms.count(item)
        alarms_dict[item] = count

    sorted_dict = {}
    sorted_keys = sorted(alarms_dict, key=alarms_dict.get, reverse=True)
    for x in sorted_keys:
        sorted_dict[x] = alarms_dict[x]
    return sorted_dict


def printing_module_names(alarms: dict):
    index = 0
    for key, value in alarms.items():
        index += 1
        if index > 10:
            break
        modules_text.insert(END,str(index)+ ". " + str(key) + '\n')


def attributes_of_exact_alarm(conn: pyodbc):
    curs = conn.cursor()
    attribute = []
    attribute_dict = dict()
    module_name = module_name_text.get()
    sql_alarms = """
            SELECT [Attribute] FROM JOURNAL WHERE [Event Type] LIKE ? AND [State] LIKE ? AND [Module] LIKE ?
            """
    if choice_lang.get() == 1:
        params = ('АЛАРМ', 'АКТ/НЕПДТВ', module_name)
    else:
        params = ('ALARM', 'ACT/UNACK', module_name)
    curs.execute(sql_alarms, params)

    for row in curs.fetchall():
        attribute.insert(0, row[0])

    alarms_set = set(attribute)
    for item in alarms_set:
        count = attribute.count(item)
        attribute_dict[item] = count

    sorted_dict = {}
    sorted_keys = sorted(attribute_dict, key=attribute_dict.get, reverse=True)
    for x in sorted_keys:
        sorted_dict[x] = attribute_dict[x]

    index = 0
    for key, value in sorted_dict.items():
        index += 1
        text.insert(END, str(index) + '. Алармы модуля ' + module_name_text.get() + ' имеют атрибуты ' + str(key)
                    + ' в колличестве ' + str(value) + '\n')


def description_of_exact_alarm(conn: pyodbc):
    curs = conn.cursor()
    descriptions = []
    description_dict = dict()

    module_name = module_name_text.get()
    sql_alarms = """
                SELECT [Desc2] FROM JOURNAL WHERE [Event Type] LIKE ? AND [State] LIKE ? AND [Module] LIKE ?
                """
    if choice_lang.get() == 1:
        params = ('АЛАРМ', 'АКТ/НЕПДТВ', module_name)
    else:
        params = ('ALARM', 'ACT/UNACK', module_name)
    curs.execute(sql_alarms, params)

    for row in curs.fetchall():
        descriptions.insert(0, row[0])

    alarms_set = set(descriptions)
    for item in alarms_set:
        count = descriptions.count(item)
        description_dict[item] = count

    sorted_dict = {}
    sorted_keys = sorted(description_dict, key=description_dict.get, reverse=True)
    for x in sorted_keys:
        sorted_dict[x] = description_dict[x]

    index = 0
    for key, value in sorted_dict.items():
        index += 1
        text.insert(END, str(index) + ' Алармы модуля ' + module_name_text.get() + ' ' + str(key)
                    + ' в колличестве ' + str(value) + '\n')


def alarm_flood(conn: pyodbc):
    curs = conn.cursor()
    global start_flow
    alarms_time_stamp = []

    sql_alarms = """
           SELECT [date/time] FROM JOURNAL WHERE [Event Type] LIKE ? AND [State] LIKE ? 
           """
    if choice_lang.get() == 1:
        params = ('АЛАРМ', 'АКТ/НЕПДТВ')
    else:
        params = ('ALARM', 'ACT/UNACK')
    curs.execute(sql_alarms, params)
    for row in curs.fetchall():
        alarms_time_stamp.insert(0, row[0])

    alarms_time_stamp.reverse()

    cycle(alarms_time_stamp)
    sorted_tuple = sorted(start_flow.items(), key=lambda x: x[1], reverse=True)
    start_flow = dict(sorted_tuple)
    text.insert(END, '10 наиболее интенсивных периода:' + '\n')
    index = 0
    for key, value in start_flow.items():
        index += 1
        if index > 10:
            break
        text.insert(END,
                    str(index) + ' Начало периода ' + str(key) + ' получено алармов за период ' + str(value) + '\n')
    text.insert(END,
                'Каждый период равняется 10 минутам, согласно стандарту ISA 18.2 если оператору приходит более 10 ' + '\n'
                + ' алармов за 10 минут он перестает на них реагировать.')


count = 0
start_flow = {}
def cycle(alarms: list) -> dict:
    global count
    global start_flow
    inside_alarms = alarms
    print(str(len(inside_alarms)) + ' Размер списка  в начале')
    start_time = alarms[0]

    while len(inside_alarms) > 2:
        for i in range(len(alarms)):
            if len(inside_alarms) < 2:
                break
            diff = inside_alarms[1].timestamp() - inside_alarms[0].timestamp()

            if diff > 600:
                # print(str(diff) + '  разница больше 600  из первого блока')
                inside_alarms.pop(0)
                start_flow.setdefault(start_time, count)
                count = 0
            #  print(str(len(inside_alarms)) + 'размер первый блок')

            if diff < 600:
                #  print(str(diff) + ' разница  из второго блока')
                start_time = inside_alarms[0]
                count += 1
                inside_alarms.pop(1)
            #  print(str(len(inside_alarms)) + 'размер второй блок')

    return start_flow


def clear_screen():
    text.delete(1.0, END)


def clear_module_names():
    modules_text.delete(1.0, END)



def help_file_open():
    webbrowser.open('Help.txt', 'r')


def exit_programm():
    root.destroy()


# разметка рабочего окна
top_frame = Frame(root, bg='white', bd=1)
middle_frame = Frame(root, bg='black', bd=1)
bottom_frame = Frame(root, bg='white', bd=1)
top_frame.place(relx=0, rely=0, relwidth=1)
middle_frame.place(relx=0, rely=0.04)
bottom_frame.place(relx=0, rely=0.08, relheight=1, relwidth=1)

main_menu = Menu(root)
root.config(menu=main_menu)

file_menu = Menu(main_menu, tearoff=0)
edit_menu = Menu(main_menu, tearoff=0)
help_menu = Menu(main_menu, tearoff=0)
file_menu.add_command(label="Открыть", command=choose_file)
file_menu.add_command(label="Выход", command=exit_programm)

edit_menu.add_command(label="Подсчет", command=lambda: counting_alarms_of_modules())
edit_menu.add_command(label='Алармы', command=lambda: find_alarms())
edit_menu.add_command(label="Поток алармов", command=lambda: alarm_flood(connection(path_text.get())))
edit_menu.add_command(label="Атрибуты", command=lambda: attributes_of_exact_alarm(connection(path_text.get())))
edit_menu.add_command(label="Описание", command=lambda: description_of_exact_alarm(connection(path_text.get())))
edit_menu.add_command(label="Сменить язык", command=lambda: print_laguage(choice_lang.get()))
edit_menu.add_command(label="Очистить", command=clear_screen)
help_menu.add_command(label="Помощь", command=help_file_open)

main_menu.add_cascade(label="Файл", menu=file_menu)
main_menu.add_cascade(label="Редактировать", menu=edit_menu)
main_menu.add_cascade(label="Помощь", menu=help_menu)
# ввод пути к файлу
path_text = Entry(top_frame, font='Tahoma 12', width=62, bd=4)
path_text.pack(side='left')
path_text.insert(END, 'Введите путь к файлу *.mdb')
# кнопка открыть файл
btn_file = Button(top_frame, text="Выбрать файл", command=choose_file)
btn_file.pack(side='left')
# текстовое поля для имени модуля
module_name_text = Entry(top_frame, font='Tahoma 11', width=15, bd=4)
module_name_text.pack(side='left')
module_name_text.insert(END, 'Название модуля')

# текстовое поле

text = scrolledtext.ScrolledText(bottom_frame,
                                 width=81, height=40,
                                 font=("Times New Roman", 12))
text.pack(side='left')
# текстовое поле для модулей
modules_text = scrolledtext.ScrolledText(bottom_frame,
                                         width=20, height=40,
                                         font=("Times New Roman", 12))
modules_text.pack(side='left')

# функциаональные кнопки
prnt = Button(middle_frame, text='Подсчет', command=lambda: counting_alarms_of_modules())
prnt.pack(side='left')

alarm_btn = Button(middle_frame, text='Алармы', command=lambda: find_alarms())
alarm_btn.pack(side='left')

alarm_btn = Button(middle_frame, text='Поток_алармов', command=lambda: alarm_flood(connection(path_text.get())))
alarm_btn.pack(side='left')

lang = Label(middle_frame, text = 'Работа с модулем')
lang.pack(side='left')

alarm_btn = Button(middle_frame, text='Атрибуты', command=lambda: attributes_of_exact_alarm(connection(path_text.get())))
alarm_btn.pack(side='left')

alarm_btn = Button(middle_frame, text='Описание', command=lambda: description_of_exact_alarm(connection(path_text.get())))
alarm_btn.pack(side='left')

alarm_btn = Button(middle_frame, text='Очистить', command=lambda: clear_screen())
alarm_btn.pack(side='left')

choice_lang = IntVar()
check = Checkbutton(middle_frame, bd=0, variable=choice_lang, onvalue=1, offvalue=0)
check.pack(side='left', anchor='e')
choice_lang.set(1)

alarm_btn = Button(middle_frame, text='Очистить_модули', command=lambda: clear_module_names())
alarm_btn.pack(side='right')


def print_laguage(choice: int):
    if choice == 1:
        text.insert(END, 'Выбранный язык БД - Русский' + '\n')
    else:
        text.insert(END, 'Выбраный язык БД - Английский' + '\n')


# метка выбранного языка
lang = Label(middle_frame, text = 'Язык БД')
lang.pack(side='right', anchor='e')

root.mainloop()
