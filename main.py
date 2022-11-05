import operator
from tkinter import *
import tkinter.filedialog as fd
from tkinter import scrolledtext
import pyodbc
import datetime

root = Tk()

def setWindow(root):
    root.title('Анализ журнала алармов и событий')
    root.resizable(False,False)
    w = 800
    h = 700
    ws = root.winfo_screenwidth()
    wh = root.winfo_screenheight()
    x = int(ws / 2 - w / 2)
    y = int(wh / 2 - h / 2)
    root.geometry("{0}x{1}+{2}+{3}".format(w,h,x,y))

setWindow(root)

def choose_file()-> str:
    global path_text
    filetypes = ("База данных", "*.mdb"),

    filename = fd.askopenfilename(title="Открыть файл", initialdir="/",
                                      filetypes=filetypes)
    if filename:
        #path_text.configure(text='')
        path_text.delete(0,30)
        path_text.insert(0,filename)
        print(filename)

def connection(path :str)->pyodbc:
    global text
    try:
        dr = r'DRIVER={Microsoft Access Driver (*.mdb, *.accdb)};'
        con_string = dr + \
                             r'DBQ=' + path
        conn = pyodbc.connect(con_string)
        return conn

        #print(row)
    except pyodbc.Error as e:
        print("Error in Connection")
        text.insert(END, 'Ошибка подключения к базе данных, убедитесь что файл выбран верно.'+'\n')


def find_alarms(conn: pyodbc):# общее представление об алармах в базе
    curs = conn.cursor()
    alarms = []
    critical = []
    predupred =[]
    signaln = []

    sql_alarms = """
    SELECT [Module] FROM JOURNAL WHERE [Event Type] LIKE ? AND [State] LIKE ? 
    """
    params = ('ALARM','ACT/UNACK')
    curs.execute(sql_alarms,params ) ##SELECT [Event Type] FROM JOURNAL WHERE  ?
    for row in curs.fetchall():
        alarms.insert(0,row[0])
        #text.insert(END,row[0]+'\n') #
        #print(row)

    sql_critical = """
        SELECT [Module] FROM JOURNAL WHERE [Event Type] LIKE ? AND [State] LIKE ? AND [Level] LIKE ?
        """
    params_critical = ('ALARM', 'ACT/UNACK','15-CRITICAL')
    curs.execute(sql_critical, params_critical)
    for row in curs.fetchall():
        critical.insert(0,row[0])
        #print(row)

    sql_alarms = """
        SELECT [Module] FROM JOURNAL WHERE [Event Type] LIKE ? AND [State] LIKE ? AND [Level] LIKE ?
        """
    params = ('ALARM', 'ACT/UNACK', '11-WARNING')
    curs.execute(sql_alarms, params)
    for row in curs.fetchall():
        predupred.insert(0, row[0])
       # print(row)

    sql_alarms = """
        SELECT [Module] FROM JOURNAL WHERE [Event Type] LIKE ? AND [State] LIKE ? AND [Level] LIKE ?
        """
    params = ('ALARM', 'ACT/UNACK', '07-ADVISORY')
    curs.execute(sql_alarms, params)
    for row in curs.fetchall():
        signaln.insert(0, row[0])
        #print(row)
    hour = time_difference(connection(path_text.get()))
    alarm_per_hour =  (len(alarms) / int(hour))//1
    alarm_per_minute =  (len(alarms) / (int(hour)*60))//1
    text.insert(END, 'Общее колличество алармов '+ str(len(alarms))+'\n')
    text.insert(END, 'Колличество критических алармов ' + str(len(critical))+'\n')
    text.insert(END, 'Колличество предупредительных алармов ' + str(len(predupred))+'\n')
    text.insert(END, 'Колличество сигнальных алармов ' + str(len(signaln)) + '\n')
    text.insert(END, ''  '\n')
    text.insert(END, 'Исследуемый период времени ' + str(hour)+' часов' + '\n')
    text.insert(END, 'Колличество алармов в час ' + str(alarm_per_hour) + '\n')
    text.insert(END, 'Колличество алармов в минуту ' + str(alarm_per_minute) + '\n')
    text.insert(END, 'Согласно ISA 18.2 - Управление  системами сигнализации для обрабатывающих отраслей промышленности'+'\n'
                + ' низкий допустимы уровнеь <6  и высокий допустимый < 12  в час')


def time_difference(conn: pyodbc)->float:
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
    time_finish = date_time[len(date_time)-1]
    time_dif = time_start - time_finish
    time_dif_hours = (time_dif.days*24 + time_dif.seconds/3600)//1


    return time_dif_hours

def counting_alarms_of_modules(conn: pyodbc):
    curs = conn.cursor()
    alarms = []
    #alarms_set = set()
    alarms_dict = dict()

    sql_alarms = """
        SELECT [Module] FROM JOURNAL WHERE [Event Type] LIKE ? AND [State] LIKE ? 
        """
    params = ('ALARM', 'ACT/UNACK')
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
    sorted_keys = sorted(alarms_dict,key=alarms_dict.get, reverse= True)
    for x in sorted_keys:
        sorted_dict[x] = alarms_dict[x]

    text.insert(END, 'Размер ' + str(len(alarms_dict))   + '\n')
    index = 0
    for key,value in sorted_dict.items():
        index+=1
        text.insert(END,str(index)+ ' Модуле ' + str(key) + ' сгенерировал '+ str(value) + ' алармов'+ '\n')

   # text.insert(END, 'Алармы ' + str(alarms_dict) + '\n')
#def alarm_flood():


def clear_screen():
    print("Очистка")
    text.delete(1.0,END)




#разметка рабочего окна
top_frame = Frame(root, bg='white', bd=1)
middle_frame = Frame(root, bg='black', bd=1)
bottom_frame = Frame(root, bg='white', bd=1)
top_frame.place(relx=0,rely=0,relwidth=1)
middle_frame.place(relx=0,rely=0.04)
bottom_frame.place(relx=0,rely=0.08,relheight=1,relwidth=1)
# ввод пути к файлу
path_text = Entry(top_frame, font='Tahoma 12', width=60,bd=4)
path_text.pack(side='left')
path_text.insert(END,'Введите путь к файлу *.mdb')
#кнопка открыть файл
btn_file = Button(top_frame, text="Выбрать файл", command=choose_file)
btn_file.pack(side='left')

#текстовое поле

text= scrolledtext.ScrolledText(bottom_frame,
                                      width=40, height=40,
                                      font=("Times New Roman", 12))
text.pack(fill='both')

#функциаональные кнопки
prnt = Button(middle_frame, text='Подсчет', command= lambda: counting_alarms_of_modules(connection(path_text.get())))
prnt.pack(side='left')

alarm_btn = Button(middle_frame, text='Алармы', command= lambda:find_alarms(connection(path_text.get())))
alarm_btn.pack(side='left')

alarm_btn = Button(middle_frame, text='Очистить', command=lambda :clear_screen())
alarm_btn.pack(side='left')

#метка временная
label2 = Label(middle_frame, text='Место для кнопок')


label2.pack(side='left')


root.mainloop()
