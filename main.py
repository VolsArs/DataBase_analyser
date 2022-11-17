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


def setWindow(root):
    root.title('Анализ журнала алармов и событий')
    root.resizable(False, False)

    w = 800
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


def find_alarms(conn: pyodbc):  # общее представление об алармах в базе
    curs = conn.cursor()
    alarms = []
    critical = []
    predupred = []
    signaln = []

    sql_alarms = """
    SELECT [Module] FROM JOURNAL WHERE [Event Type] LIKE ? AND [State] LIKE ? 
    """
    params = ('АЛАРМ', 'АКТ/НЕПДТВ')
    curs.execute(sql_alarms, params)  ##SELECT [Event Type] FROM JOURNAL WHERE  ?
    for row in curs.fetchall():
        alarms.insert(0, row[0])
        # text.insert(END,row[0]+'\n') #
        # print(row)

    sql_critical = """
        SELECT [Module] FROM JOURNAL WHERE [Event Type] LIKE ? AND [State] LIKE ? AND [Level] LIKE ?
        """
    params_critical = ('АЛАРМ', 'АКТ/НЕПДТВ', '15-КРИТИЧЕСКИЙ')
    curs.execute(sql_critical, params_critical)
    for row in curs.fetchall():
        critical.insert(0, row[0])
        # print(row)

    sql_alarms = """
        SELECT [Module] FROM JOURNAL WHERE [Event Type] LIKE ? AND [State] LIKE ? AND [Level] LIKE ?
        """
    params = ('АЛАРМ', 'АКТ/НЕПДТВ', '11-ПРЕДУПРЕДИТЕЛЬН')
    curs.execute(sql_alarms, params)
    for row in curs.fetchall():
        predupred.insert(0, row[0])
    # print(row)

    sql_alarms = """
        SELECT [Module] FROM JOURNAL WHERE [Event Type] LIKE ? AND [State] LIKE ? AND [Level] LIKE ?
        """
    params = ('АЛАРМ', 'АКТ/НЕПДТВ', '07-СИГНАЛЬНЫЙ')
    curs.execute(sql_alarms, params)
    for row in curs.fetchall():
        signaln.insert(0, row[0])
        # print(row)
    hour = time_difference(connection(path_text.get()))
    alarm_per_hour = (len(alarms) / int(hour)) // 1
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
    curs.execute(sql_time, params)  ##SELECT [Event Type] FROM JOURNAL WHERE  ?
    for row in curs.fetchall():
        date_time.insert(0, row[0])
        # text.insert(END,row[0]+'\n') #

    time_start = date_time[0]
    time_finish = date_time[len(date_time) - 1]
    time_dif = time_start - time_finish
    time_dif_hours = (time_dif.days * 24 + time_dif.seconds / 3600) // 1

    return time_dif_hours


def counting_alarms_of_modules(conn: pyodbc):
    curs = conn.cursor()
    alarms = []
    # alarms_set = set()
    alarms_dict = dict()

    sql_alarms = """
        SELECT [Module] FROM JOURNAL WHERE [Event Type] LIKE ? AND [State] LIKE ? 
        """
    params = ('АЛАРМ', 'АКТ/НЕПДТВ')
    curs.execute(sql_alarms, params)  ##SELECT [Event Type] FROM JOURNAL WHERE  ?
    for row in curs.fetchall():
        alarms.insert(0, row[0])
    # for x in alarms:
    #  alarms_set.add(alarms[x])
    #  text.insert(END, 'Общее колличество алармов set ' + str(alarms_set.) + '\n')
    alarms_set = set(alarms)

    for item in alarms_set:
        count = alarms.count(item)
        alarms_dict[item] = count

    sorted_dict = {}
    sorted_keys = sorted(alarms_dict, key=alarms_dict.get, reverse=True)
    for x in sorted_keys:
        sorted_dict[x] = alarms_dict[x]

    #text.insert(END, 'Размер ' + str(len(alarms_dict)) + '\n')
    index = 0
    for key, value in sorted_dict.items():
        index += 1
        text.insert(END, str(index) + ' Модуль ' + str(key) + ' сгенерировал ' + str(value) + ' алармов' + '\n')


count = 0
start_flow = {}


def alarm_flood(conn: pyodbc):
    curs = conn.cursor()
    global start_flow
    alarms_time_stamp = []

    sql_alarms = """
           SELECT [date/time] FROM JOURNAL WHERE [Event Type] LIKE ? AND [State] LIKE ? 
           """
    params = ('АЛАРМ', 'АКТ/НЕПДТВ')
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
                print(str(diff) + '  разница больше 600  из первого блока')
                inside_alarms.pop(0)
                start_flow.setdefault(start_time, count)
                count = 0
                print(str(len(inside_alarms)) + 'размер первый блок')

            if diff < 600:
                print(str(diff) + ' разница  из второго блока')
                start_time = inside_alarms[0]
                count += 1
                inside_alarms.pop(1)
                print(str(len(inside_alarms)) + 'размер второй блок')

    return start_flow


def clear_screen():
    print("Очистка")
    text.delete(1.0, END)


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
edit_menu.add_command(label="Подсчет", command=lambda: counting_alarms_of_modules(connection(path_text.get())))
edit_menu.add_command(label='Алармы', command=lambda: find_alarms(connection(path_text.get())))
edit_menu.add_command(label="Поток алармов", command=lambda: alarm_flood(connection(path_text.get())))
edit_menu.add_command(label="Очистить", command=clear_screen)
help_menu.add_command(label="Помощь", command=help_file_open)

main_menu.add_cascade(label="Файл", menu=file_menu)
main_menu.add_cascade(label="Редактировать", menu=edit_menu)
main_menu.add_cascade(label="Помощь", menu=help_menu)
# ввод пути к файлу
path_text = Entry(top_frame, font='Tahoma 12', width=60, bd=4)
path_text.pack(side='left')
path_text.insert(END, 'Введите путь к файлу *.mdb')
# кнопка открыть файл
btn_file = Button(top_frame, text="Выбрать файл", command=choose_file)
btn_file.pack(side='left')

# текстовое поле

text = scrolledtext.ScrolledText(bottom_frame,
                                 width=40, height=40,
                                 font=("Times New Roman", 12))
text.pack(fill='both')

# функциаональные кнопки
prnt = Button(middle_frame, text='Подсчет', command=lambda: counting_alarms_of_modules(connection(path_text.get())))
prnt.pack(side='left')

alarm_btn = Button(middle_frame, text='Алармы', command=lambda: find_alarms(connection(path_text.get())))
alarm_btn.pack(side='left')

alarm_btn = Button(middle_frame, text='Поток_алармов', command=lambda: alarm_flood(connection(path_text.get())))
alarm_btn.pack(side='left')

alarm_btn = Button(middle_frame, text='Очистить', command=lambda: clear_screen())
alarm_btn.pack(side='left')

# метка временная
# label2 = Label(middle_frame, text='Место для кнопок')

# label2.pack(side='left')

root.mainloop()
