from tkinter import *  # библиотека для работы с графическим окном
import time  # библиотека для работы с временем
import glob  # чтобы получить список файлов
import datetime  # библиотека для работы с датой и временем
import os.path  # чтобы проверить существование файла
from PIL import Image, ImageTk, ImageEnhance  # библиотека для работы с картинками
import pickle  # загрузка бинарных файлов
from io import BytesIO  # библиотека для байтового ввода-вывода

root = Tk()  # окно в переменную
root.title('Онлайн-контроль качества сборки изделий на ПМИ')  # название окна

# переводим на русский названия дефектов
DefectType = [
    ('COPLANARITY (Not use)', ''),
    ('Wrong Size', 'неверные размеры'),
    ('No Comp', 'отсутствует'),
    ('No comp', 'отсутствует'),
    ('COMPONENT_SHIFT (Not use)', ''),
    ('Head over heels', 'перевернут'),
    ('POSITION', ''),
    ('Bad_Solder', 'непропай'),
    ('Bad Solder', 'непропай'),
    ('Lifted lead', 'приподнят'),
    ('Lifted body', 'приподнят'),
    ('Solder Brige', 'перемычка'),
    ('Solder Bridge', 'перемычка'),
    ('BILL_BOARDING (Not use)', ''),
    ('TOMBSTONE', '«надгробный камень»'),
    ('BODY_DIMENSION (Not use)', ''),
    ('Polarity', 'неверная полярность'),
    ('Wrong marking', 'неверная маркировка'),
    ('Wrong comp', 'неверная маркировка'),
    ('ABSENCE', 'отсутствует'),
    ('Out of position', 'смещение'),
    ('No lead', 'отсутствует вывод'),
    ('PARTICLE (Not use)', ''),
    ('Foreign body', 'инородный корпус'),
    ('Foreign lead', 'инородный вывод'),
    ('WFMI', ''),
    ('CDM', ''),
    ('UNDEFINED', 'нужна проверка'),
    ('PADOVERHANG', 'приподнят вывод'),
    ('COPLANARITY', ''),
    ('DIMENSION', 'неверные размеры'),
    ('MISSING', 'отсутствует'),
    ('COMPONENT_SHIFT', ''),
    ('UPSIDEDOWN', 'перевернут'),
    ('POSITION', ''),
    ('SOLDER_JOINT', 'непропай'),
    ('LIFTED_LEAD', 'приподнят'),
    ('LIFTED_BODY', 'приподнят'),
    ('BRIDGING', 'перемычка'),
    ('BILL_BOARDING', ''),
    ('TOMBSTONE', '«надгробный камень»'),
    ('BODY_DIMENSION', ''),
    ('POLARITY', 'неверная полярность'),
    ('OCR_OCV', 'неверная маркировка'),
    ('ABSENCE', 'отсутствует'),
    ('OVERHANG', 'смещение'),
    ('MISSING_LEAD', 'отсутствует вывод'),
    ('PARTICLE', ''),
    ('FOREIGNMATERIAL_BODY', 'инородный корпус'),
    ('FOREIGNMATERIAL_LEAD', 'инородный вывод'),
    ('WFMI', ''),
    ('CDM', ''),
    ('Other', 'другое')
]

# Линия 1 - шаблон

l1_history = Text(root, bg="grey95", width=27, height=3, borderwidth=0, relief='solid')  # нет обводки
l1_history['font'] = ('Arial', 23)
l1_history['foreground'] = 'brown'
l1_history.grid(row=0, column=0)
l1_history.insert(1.0, 'Собирали\n')
l1_history.insert(2.0, '')

tx_1_2 = Text(root, bg="grey95", width=27, height=3, borderwidth=0, relief='solid')  # нет обводки
tx_1_2['font'] = ('Arial', 23)
tx_1_2['foreground'] = 'brown'
scr_1_2 = Scrollbar(root, command=tx_1_2.yview)
tx_1_2.configure(yscrollcommand=scr_1_2.set)
tx_1_2.grid(row=0, column=1)
scr_1_2.grid(row=0, column=2, ipady=30)
tx_1_2.insert(1.0, 'Дефекты сборки:\n')

canvas_2_2 = Canvas(root, bg="grey95", width=470, height=807, borderwidth=1, relief='solid')  # grey95
canvas_2_2.grid(row=1, column=0)

tx_2_0 = Text(root, bg="grey95", width=27, height=23, borderwidth=1, relief='solid')  # обведено тонкой линией
tx_2_0['font'] = ('Arial', 23)
scr_2_0 = Scrollbar(root, command=tx_2_0.yview)
tx_2_0.configure(yscrollcommand=scr_2_0.set)
tx_2_0.grid(row=1, column=1)
scr_2_0.grid(row=1, column=2, ipady=380)
tx_2_0.insert(1.0, 'Платы на линии 1 (NEVO):\n')

l1_now = Text(root, bg="grey95", width=25, height=3, borderwidth=3, relief='solid')  # обведено толстой линией
l1_now.grid(row=2, column=0)
l1_now['font'] = ('Arial', 25)
l1_now['foreground'] = 'darkgreen'  # 'red'
l1_now.insert(1.0, 'Линия 1: \n')
l1_now.insert(2.0, 'Выход годных плат: ')
l1_now.insert(3.0, '')

tx_0_2 = Text(root, bg="grey95", width=25, height=3, borderwidth=3, relief='solid')  # обведено толстой линией
tx_0_2['font'] = ('Arial', 25)
tx_0_2['foreground'] = 'darkgreen'  # 'red'
scr_0_2 = Scrollbar(root, command=tx_0_2.yview)
tx_0_2.configure(yscrollcommand=scr_0_2.set)
tx_0_2.grid(row=2, column=1)
scr_0_2.grid(row=2, column=2, ipady=30)
tx_0_2.insert(1.0, 'Текущие дефекты:\n')

# Линия 2 - шаблон

tx_1_3 = Text(root, bg="grey95", width=27, height=3, borderwidth=0, relief='solid')  # нет обводки
tx_1_3['font'] = ('Arial', 23)
tx_1_3['foreground'] = 'darkblue'
tx_1_3.grid(row=0, column=3)
tx_1_3.insert(1.0, 'Собирали\n')
tx_1_3.insert(2.0, '')

tx_1_5 = Text(root, bg="grey95", width=27, height=3, borderwidth=0, relief='solid')  # нет обводки
tx_1_5['font'] = ('Arial', 23)
tx_1_5['foreground'] = 'darkblue'
scr_1_5 = Scrollbar(root, command=tx_1_5.yview)
tx_1_5.configure(yscrollcommand=scr_1_5.set)
tx_1_5.grid(row=0, column=4)
scr_1_5.grid(row=0, column=5, ipady=20)
tx_1_5.insert(1.0, 'Дефекты сборки:\n')

canvas_2_5 = Canvas(root, bg="grey95", width=470, height=807, borderwidth=1, relief='solid')  # grey95
canvas_2_5.grid(row=1, column=3)

tx_2_3 = Text(root, bg="grey95", width=27, height=23, borderwidth=1, relief='solid')  # обведено тонкой линией
tx_2_3['font'] = ('Arial', 23)
scr_2_3 = Scrollbar(root, command=tx_2_3.yview)
tx_2_3.configure(yscrollcommand=scr_2_3.set)
tx_2_3.grid(row=1, column=4)
scr_2_3.grid(row=1, column=5, ipady=380)
tx_2_3.insert(1.0, 'Платы на линии 2 (NEVO):\n')

tx_0_3 = Text(root, bg="grey95", width=25, height=3, borderwidth=3, relief='solid')  # обведено толстой линией
tx_0_3.grid(row=2, column=3)
tx_0_3['font'] = ('Arial', 25)
tx_0_3['foreground'] = 'darkblue'  # 'red'
tx_0_3.insert(1.0, 'Линия 2: \n')
tx_0_3.insert(2.0, 'Выход годных плат: ')
tx_0_3.insert(3.0, '')

tx_0_5 = Text(root, bg="grey95", width=25, height=3, borderwidth=3, relief='solid')  # обведено толстой линией
tx_0_5['font'] = ('Arial', 25)
tx_0_5['foreground'] = 'darkblue'  # 'red'
scr_0_5 = Scrollbar(root, command=tx_0_5.yview)
tx_0_5.configure(yscrollcommand=scr_0_5.set)
tx_0_5.grid(row=2, column=4)
scr_0_5.grid(row=2, column=5, ipady=20)
tx_0_5.insert(1.0, 'Текущие дефекты:\n')

# переменные для размещения картинок
step_x = 155
step_y = 162
folder1 = '//tk22001s/DATA/Documents/Цех №1/Сетевая/Отчеты оптика AOI/QX/XML_PM'  # здесь отчеты в формате xml
result = []  # список для хранения результатов


def f_teko():  # find information about product in TEKO.csv
    list_teko = []
    with open('\\\\tk22001s\\DATA\\Documents\\Цех №1\\Сетевая\\Отчеты оптика AOI\\TEKO\\TEKO.csv') as file1:
        for line in file1:
            column = 0
            teko1 = ['', '', '']
            for char in line:
                if char == ',':
                    column += 1
                if char != '\n' and char != ',':
                    teko1[column] += char
            list_teko.append(teko1)
    teko_dash = []
    for line in list_teko:
        if '-' not in line[0]:
            continue
        else:
            teko_dash.append(line[0])
    return teko_dash


teko_dash = f_teko()
last_5_images_kz_previous = []
num_of_boards_1 = 0
last_5_images_qx_previous = []
num_of_boards_2 = 0


def ask_info():  # функция - обработчик
    global last_5_images_kz_previous, num_of_boards_1, last_5_images_qx_previous, num_of_boards_2, last_5_images_kz, last_5_images_qx
    today_now = datetime.datetime.now()  # сегодня
    today_minus_1 = today_now - datetime.timedelta(days=1)  # вчера
    today1 = today_now.strftime('%Y_%m_%d')
    today_1 = today_minus_1.strftime('%Y_%m_%d')
    # print(today1, today_1)

    # получаем данные со второй линии
    result_file2 = []
    file_name = '\\\\tk22001s\\DATA\\Documents\\Цех №1\\Сетевая\\Отчеты оптика AOI\\Data_AOI_QX_1d_30s.csv'
    with open(file_name) as file:
        for line in file:
            result_file2.append(line.split(';'))
    # print(result_file2)

    # запрашиваем данные о предыдущей сборке по линии 2
    num = len(result_file2) - 1
    file_result2 = '\\\\tk22001s\\DATA\\Documents\\Цех №1\\Сетевая\\Отчеты оптика AOI\\QX\\file_result2.txt'
    temp2_file_qx = []
    if len(result_file2) > 1:
        with open(file_result2) as file:
            for line in file:
                line = line.replace('\n', '')
                if 'Дата' in line:
                    continue
                # эта сборка была не сегодня или вчера
                temp3_file = line.split(';')
                if result_file2[num][2] == temp3_file[2] and today1 not in line and today_1 not in line:
                    if int(temp3_file[4]) > 3:
                        temp2_file_qx = temp3_file
    # print(temp2_file_qx)
    # input()

    # получаем данные с первой линии
    result_file = []
    file_name = '\\\\tk22001s\\DATA\\Documents\\Цех №1\\Сетевая\\Отчеты оптика AOI\\Data_AOI_KZ_5d_10min.csv'
    with open(file_name) as file:
        for line in file:
            result_file.append(line.split(';'))
    num = len(result_file) - 1
    if len(result_file) != 0:  # переводим названия дефектов на русский язык (текущая сборка)
        for line in DefectType:
            if line[0] in result_file[num][9]:
                temp = result_file[num][9].split(' ')
                result_file[num][9] = temp[0]+' '+line[1]
            if line[0] in result_file[num][11]:
                temp1 = result_file[num][11].split(' ')
                result_file[num][11] = temp1[0]+' '+line[1]
            if line[0] in result_file[num][13]:
                temp2 = result_file[num][13].split(' ')
                result_file[num][13] = temp2[0]+' '+line[1]
    # print(result_file)

    # запрашиваем данные о предыдущей сборке по линии 1
    today1 = today_now.strftime('%d.%m.%Y')
    today_1 = today_minus_1.strftime('%d.%m.%Y')
    file_name2 = '\\\\tk22001s\\DATA\\Documents\\Цех №1\\Сетевая\\Отчеты оптика AOI\\Stat_data_AOI.csv'
    temp2_file = []
    with open(file_name2) as file:
        for line in file:
            line = line.replace('\n', '')
            if 'Дата' in line:  # пропускаем заголовок
                continue
            # эта сборка была не сегодня или вчера
            temp3_file = line.split(';')
            if result_file[num][2] == temp3_file[2] and today1 not in line and today_1 not in line:
                if int(temp3_file[4]) > 5:  # рассматриваем результаты с количеством плат больше 5
                    temp2_file = temp3_file
    if len(temp2_file) != 0:  # переводим названия дефектов на русский язык (предыдущая сборка)
        for line in DefectType:
            if line[0] in temp2_file[9]:
                temp2_file[9] = temp2_file[9].replace(line[0], line[1])
            if line[0] in temp2_file[11]:
                temp2_file[11] = temp2_file[11].replace(line[0], line[1])
            if line[0] in temp2_file[13]:
                temp2_file[13] = temp2_file[13].replace(line[0], line[1])
        good = ((int(temp2_file[3]) - int(temp2_file[7])) / int(temp2_file[3])) * 100
        good = float('{:.1f}'.format(good))  # выход годных (предыдущая сборка)
        temp4 = temp2_file[0]+' '+str(good)+'%'
        if temp2_file[10] != '0':
            temp4 += ' '+temp2_file[9]
        if temp2_file[12] != '0':
            temp4 += ' '+temp2_file[11]
        if temp2_file[14] != '0':
            temp4 += ' '+temp2_file[13]  # собрали информацию об дефектах предыдущей сборки
    # print(temp2_file)

    file_name_3 = '\\\\tk22001s\\DATA\\Documents\\Цех №1\\Сетевая\\Отчеты оптика AOI\\last_ml_qx.csv'
    last_ml_qx = []
    with open(file_name_3) as file:  # получили данные об последних мультилистах с оптики QX500
        for line in file:
            line = line.replace('\n', '')
            line = line.replace('Нужна проверка', 'проверка')
            last_ml_qx.append(line)
    last_ml_qx_list = []
    for i in range(len(last_ml_qx)):
        temp5 = ''
        if i + 1 < 10 and len(last_ml_qx) > i:
            temp5 = '  ' + str(i + 1) + ':' + ' ' + last_ml_qx[i]
        elif i + 1 >= 10 and len(last_ml_qx) > i:
            temp5 = str(i + 1) + ':' + ' ' + last_ml_qx[i]
        last_ml_qx_list.append(temp5)  # преобразовали в вид, пригодный для вывода на экран
    # print(last_ml_qx_list)
    # input()

    file_last_ml = '\\\\tk22001s\\DATA\\Documents\\Цех №1\\Сетевая\\Отчеты оптика AOI\\last_ml.csv'
    last_ml = []
    with open(file_last_ml) as file:  # получили данные об последних мультилистах с оптики KohYoung
        for line in file:
            line = line.replace('\n', '')
            temp3 = line.split(';')
            last_ml.append(temp3)
    last_ml_list = []

    # определяем только результат
    for i in range(len(last_ml)):
        temp5 = ''
        if i + 1 < 10 and len(last_ml) > i:
            temp5 = '  ' + str(i + 1) + ':' + ' ' + last_ml[i][1]
        elif i + 1 >= 10 and len(last_ml) > i:
            temp5 = str(i + 1) + ':' + ' ' + last_ml[i][1]
        if last_ml[i][3] == 'Good':
            temp5 += ' ' + '+'
        elif last_ml[i][3] == 'Pass':
            temp5 += ' ' + '+\''
        elif last_ml[i][3] == 'NG':
            temp5 += ' ' + 'Ремонт'
        elif last_ml[i][3] == 'Нужна проверка':
            temp5 += ' ' + last_ml[i][4]
        temp51 = temp5
        last_ml_list.append(temp51)  # преобразовали в вид, пригодный для вывода на экран
    # print(last_ml_list)

    # изменяем данные по линии 1 на экране
    # print(result_file[1])
    l1_now.delete(1.0, END)
    if len(result_file) > 1:
        temp_v = result_file[1][15].replace('\n', '')
    else:
        temp_v = 0
    if float(temp_v) >= 97.5:
        l1_now['foreground'] = 'darkgreen'
    else:
        l1_now['foreground'] = 'darkred'
    if len(result_file) > 1:
        l1_now.insert(1.0, f'Л1(KZ):{result_file[1][2]}:{result_file[1][3]}\n')
    else:
        l1_now.insert(1.0, f'Л1(KZ):\n')
    l1_now.insert(2.0, f'Выход годных плат: {temp_v}%\n')
    num = len(result_file) - 1
    if len(result_file) > 1:
        view_1 = (int(result_file[num][4]) - int(result_file[num][5])) / int(result_file[num][4]) * 100
        view_1 = float('{:.1f}'.format(view_1))
    else:
        view_1 = 0
    l1_now.insert(3.0, f'Выход без проверки: {view_1}%')
    # print(result_file[1][2], temp_v, view_1)

    tx_0_2.delete(1.0, END)
    if result_file[num][9] != '':
        tx_0_2.insert(2.0, f'{result_file[num][9]} - {result_file[num][10]}\n')
    if result_file[num][11] != '':
        tx_0_2.insert(3.0, f'{result_file[num][11]} - {result_file[num][12]}\n')
    if result_file[num][13] != '':
        tx_0_2.insert(4.0, f'{result_file[num][13]} - {result_file[num][14]}\n')
    # print(result_file[num][9], result_file[num][11], result_file[num][13])

    if len(temp2_file) != 0:
        l1_history.delete(1.0, END)
        l1_history.insert(1.0, f'Собирали {temp2_file[0]}:{temp2_file[3]}\n')
        l1_history.insert(2.0, 'Выход годных плат: ' + str(good) + '%\n')
        view_2 = (int(temp2_file[4]) - int(temp2_file[5])) / int(temp2_file[4]) * 100
        view_2 = float('{:.1f}'.format(view_2))
        l1_history.insert(3.0, f'Выход без проверки: {view_2}%')

        tx_1_2.delete(1.0, END)
        # tx_1_2.insert(1.0, 'Дефекты сборки:\n')
        if temp2_file[10] != '0':
            tx_1_2.insert(2.0, f'{temp2_file[9]} - {temp2_file[10]}\n')
        if temp2_file[12] != '0':
            tx_1_2.insert(3.0, f'{temp2_file[11]} - {temp2_file[12]}\n')
        if temp2_file[14] != '0':
            tx_1_2.insert(4.0, f'{temp2_file[13]} - {temp2_file[14]}')
    else:
        l1_history.delete(1.0, END)
        tx_1_2.delete(1.0, END)
    # print(temp2_file[0], good, view_2, temp2_file[9], temp2_file[11], temp2_file[13])

    my_time = today_now.strftime(' %H:%M')

    tx_2_0.delete(1.0, END)
    tx_2_0.insert(1.0, f'Выход Zenith (NEVO {my_time}):\n')
    position = 2.0
    for i in range(len(last_ml_list)):
        tx_2_0.insert(position + i, f'{last_ml_list[i]}\n')
        # print(ml)

    #  картинки с 1-ой линии

    last_5_images_kz = []
    i = 1
    for ml in last_ml:
        if len(last_5_images_kz) == 5:
            break
        if ml[3] == 'NG' or ml[3] == 'Нужна проверка':
            last_5_images_kz.append([i, ml[0]])
        i += 1
    # print(last_5_images_kz)  # заготовка для сбора картинок

    folder_images = '\\\\tk22001s\\DATA\\Documents\\Цех №1\\Сетевая\\Отчеты оптика AOI\\images'  # собираем картинки
    files = glob.glob(folder_images + '\\*')
    files_list = []
    for file in files:
        files_list.append(file)
    for i in range(len(last_5_images_kz)):
        last_5_images_kz[i].append([])
        for file in files_list:
            if last_5_images_kz[i][1] in file:
                try:
                    with open(file, 'rb') as f:  # читаем в бинарном виде
                        image_file_s = pickle.load(f)
                    pil_image_1 = Image.open(BytesIO(image_file_s))  # загружаем в виде картинки
                    if float(pil_image_1.size[0]) > float(pil_image_1.size[1]):  # если ширина картинки больше высоты
                        base_width = 140  # задаем значение 140
                        ratio = (base_width / float(pil_image_1.size[0]))  # вычисляем коэффициент маштабирования
                        height = int(
                            (float(pil_image_1.size[1])) * float(ratio))  # высоту картинки вычисляем через коэффициент
                        pil_image = pil_image_1.resize((base_width, height), Image.ANTIALIAS)
                    else:  # если высота картинки больше ширины
                        base_height = 140  # задаем значение 140
                        ratio = (base_height / float(pil_image_1.size[1]))  # вычисляем коэффициент маштабирования
                        width = int(
                            (float(pil_image_1.size[0])) * float(ratio))  # ширину картинки вычисляем через коэффициент
                        pil_image = pil_image_1.resize((width, base_height), Image.ANTIALIAS)
                    pil_image = ImageEnhance.Brightness(pil_image).enhance(2.5)  # настраиваем яркость
                    pil_image = ImageEnhance.Contrast(pil_image).enhance(1.5)  # настраиваем контраст
                    pil_image = ImageEnhance.Sharpness(pil_image).enhance(5.5)  # настраиваем резкость
                    image = ImageTk.PhotoImage(pil_image)
                    last_5_images_kz[i][2].append(image)
                except PermissionError:
                    pass
                except FileNotFoundError:
                    pass

    # print('Before:', last_5_images_kz)  # собрали картинки
    if len(last_5_images_kz_previous) == 0 or len(last_ml) != num_of_boards_1:
        last_5_images_kz_previous = last_5_images_kz
        num_of_boards_1 = len(last_ml)
    elif last_5_images_kz != last_5_images_kz_previous:  # если появилась новая NG плата
        last_5_images_kz_previous = last_5_images_kz
        num_of_boards_1 = len(last_ml)
    else:  # иначе запускаем карусель из картинок
        for line in last_5_images_kz_previous:
            if len(line[2]) > 3:
                # print(line[2])
                line2 = []
                for i in range(len(line[2])):
                    if i + 1 < len(line[2]):
                        line2.append(line[2][i + 1])
                    else:
                        line2.append(line[2][0])
                line[2] = line2
                # print('Modify:', line[2])
    # print('After:', last_5_images_kz_previous)

    # выводим картинки и подписи к ним
    # вывести название компонента с дефектом (4 - BAR64) - перспективно!!!
    canvas_2_2.delete('all')
    if len(last_5_images_kz_previous) > 0:
        for y in range(len(last_5_images_kz_previous)):
            for x in range(len(last_5_images_kz_previous[y][2])):
                canvas_2_2.create_text(75 + step_x * x, 10 + step_y * y, text=last_5_images_kz_previous[y][0], justify=CENTER, font="Verdana 12")
                canvas_2_2.create_image(80 + step_x * x, 90 + step_y * y, image=last_5_images_kz_previous[y][2][x])

    # изменяем данные по линии 2
    # print(result_file2[1])
    if len(result_file2) > 1:
        if float(result_file2[1][5]) >= 97.5:
            tx_0_3['foreground'] = 'darkblue'
        else:
            tx_0_3['foreground'] = 'darkred'
        # print(result_file2)
        if len(result_file2) != 1:
            tx_0_3.delete(1.0, END)
            tx_0_3.insert(1.0, f'Л2(QX):{result_file2[1][2]}:{result_file2[1][3]}\n')
            if float(result_file2[1][5]) > 100:
                result_file2[1][5] = str(100)
            tx_0_3.insert(2.0, f'Выход годных плат: {result_file2[1][5]}%\n')
            tx_0_3.insert(3.0, f'Выход без проверки: {float(result_file2[1][6])}%')

        # temp2_file_qx = ['2021_03_25', '2', 'R-RPUMbv15_B', '176', '11', '0', '0', '100.0', '6', '45.5', 'DD1_PIC16F1503_SO-14_Polarity', '1', 'DD1_PIC16F1503_SO-14_Out of position', '1', 'C2_0603_0_1mkF_10%_No Comp', '1']

        if len(temp2_file_qx) != 0:  # переводим названия дефектов на русский язык (предыдущая сборка)
            for line in DefectType:
                if line[0] in temp2_file_qx[10]:
                    temp2_file_qx[10] = temp2_file_qx[10].replace(line[0], line[1])
                    len_qx = len(temp2_file_qx[10].split('_'))
                    temp2_file_qx[10] = temp2_file_qx[10].split('_')[0] + ' ' + temp2_file_qx[10].split('_')[len_qx-1]
                if line[0] in temp2_file_qx[12]:
                    temp2_file_qx[12] = temp2_file_qx[12].replace(line[0], line[1])
                    len_qx = len(temp2_file_qx[12].split('_'))
                    temp2_file_qx[12] = temp2_file_qx[12].split('_')[0] + ' ' + temp2_file_qx[12].split('_')[len_qx-1]
                if line[0] in temp2_file_qx[14]:
                    temp2_file_qx[14] = temp2_file_qx[14].replace(line[0], line[1])
                    len_qx = len(temp2_file_qx[14].split('_'))
                    temp2_file_qx[14] = temp2_file_qx[14].split('_')[0] + ' ' + temp2_file_qx[14].split('_')[len_qx-1]

        # print(temp2_file_qx)
        # input()

        if len(temp2_file_qx) != 0:
            tx_1_3.delete(1.0, END)
            date = datetime.datetime.strptime(temp2_file_qx[0], '%Y_%m_%d')
            date = date.strftime('%d.%m.%Y')
            tx_1_3.insert(1.0, f'Собирали {date}:{temp2_file_qx[3]}\n')
            tx_1_3.insert(2.0, 'Выход годных плат: ' + str(temp2_file_qx[7]) + '%\n')
            tx_1_3.insert(3.0, f'Выход без проверки: {temp2_file_qx[9]}%')

            tx_1_5.delete(1.0, END)
            if temp2_file_qx[11] != '0':
                tx_1_5.insert(2.0, f'{temp2_file_qx[10]} - {temp2_file_qx[11]}\n')
            if temp2_file_qx[13] != '0':
                tx_1_5.insert(3.0, f'{temp2_file_qx[12]} - {temp2_file_qx[13]}\n')
            if temp2_file_qx[15] != '0':
                tx_1_5.insert(4.0, f'{temp2_file_qx[14]} - {temp2_file_qx[15]}')

        else:
            tx_1_3.delete(1.0, END)  # очищаем историю сборки
            tx_1_5.delete(1.0, END)  # очищаем дефекты

        num = len(result_file2) - 1
        tx_0_5.delete(1.0, END)
        if result_file2[num][7] != '':
            tx_0_5.insert(2.0, f'{result_file2[num][7]} - {result_file2[num][8]}\n')
        if result_file2[num][9] != '':
            tx_0_5.insert(3.0, f'{result_file2[num][9]} - {result_file2[num][10]}\n')
        if result_file2[num][11] != '':
            tx_0_5.insert(4.0, f'{result_file2[num][11]} - {result_file2[num][12]}\n')
        # print(result_file[num][9], result_file[num][11], result_file[num][13])

        tx_2_3.delete(1.0, END)
        tx_2_3.insert(1.0, f'Выход QX500 (NEVO {my_time}):\n')
        position = 2.0
        for i in range(len(last_ml_qx_list)):
            tx_2_3.insert(position+i, f'{last_ml_qx_list[i]}\n')
            # print(ml)
        # print(result_file2[1][2], result_file2[1][5], result_file2[1][6], date, temp2_file_qx[7], temp2_file_qx[9])

    #  картинки со 2-ой линии
    file_last_ml = '\\\\tk22001s\\DATA\\Documents\\Цех №1\\Сетевая\\Отчеты оптика AOI\\last_ml_qx.csv'
    last_ml_qx = []
    with open(file_last_ml) as file:
        for line in file:
            line = line.replace('\n', '')
            temp3 = line.split(' ')[0]
            if len(line.split(' ')) > 2:
                temp7 = line.split(' ')[1] + ' ' + line.split(' ')[2]
            else:
                temp7 = line.split(' ')[1]
            last_ml_qx.append([temp3, temp7])
    # print(last_ml_qx)
    last_5_images_qx = []
    i = 1
    for ml in last_ml_qx:
        if len(last_5_images_qx) == 5:
            break
        if ml[1] == 'Ремонт' or ml[1] == 'Нужна проверка':
            last_5_images_qx.append([i, ml[0]])
        i += 1
    # print(last_5_images_qx)  # собираем картинки - шаблон
    folder_images = '\\\\tk22001s\\DATA\\Documents\\Цех №1\\Сетевая\\Отчеты оптика AOI\\QX\\ExportedImages'
    files = glob.glob(folder_images + '\\*')
    files_list = []
    today_date = datetime.datetime.now()  # смотрим данные за сегодня
    today_date_mdY = today_date.strftime('%Y_%m_%d')

    for file in files:
        if today_date_mdY in file:
            files_list.append(file)
    for i in range(len(last_5_images_qx)):
        last_5_images_qx[i].append([])
        for file in files_list:
            run = 0
            # print(teko_dash)
            for dash in teko_dash:
                if dash in file:
                    run = 1
            if run == 1:
                my_var = file.split('-')[5].replace('.jpg', '').split('.')
                # print(my_var)
            else:
                my_var = file.split('-')[3].replace('.jpg', '').split('.')
                # print(my_var)
                # input()
            if last_5_images_qx[i][1] in file and len(my_var) == 3:
                pil_image_1 = Image.open(file)
                if float(pil_image_1.size[0]) > float(pil_image_1.size[1]):  # если ширина картинки больше высоты
                    base_width = 140  # задаем значение 140
                    ratio = (base_width / float(pil_image_1.size[0]))  # вычисляем коэффициент маштабирования
                    height = int(
                        (float(pil_image_1.size[1])) * float(ratio))  # высоту картинки вычисляем через коэффициент
                    pil_image = pil_image_1.resize((base_width, height), Image.ANTIALIAS)
                else:  # если высота картинки больше ширины
                    base_height = 140  # задаем значение 140
                    ratio = (base_height / float(pil_image_1.size[1]))  # вычисляем коэффициент маштабирования
                    width = int(
                        (float(pil_image_1.size[0])) * float(ratio))  # ширину картинки вычисляем через коэффициент
                    pil_image = pil_image_1.resize((width, base_height), Image.ANTIALIAS)
                pil_image = ImageEnhance.Brightness(pil_image).enhance(3.5)  # настраиваем яркость
                pil_image = ImageEnhance.Contrast(pil_image).enhance(1.5)  # настраиваем контраст
                pil_image = ImageEnhance.Sharpness(pil_image).enhance(5.5)  # настраиваем резкость
                image = ImageTk.PhotoImage(pil_image)
                last_5_images_qx[i][2].append(image)
    # print('Before:', last_5_images_qx)  # собрали картинки
    if len(last_5_images_qx_previous) == 0 or len(last_ml_qx) != num_of_boards_2:
        last_5_images_qx_previous = last_5_images_qx
        num_of_boards_2 = len(last_ml_qx)
    elif last_5_images_qx != last_5_images_qx_previous:  # если появилась новая NG
        last_5_images_qx_previous = last_5_images_qx
        num_of_boards_2 = len(last_ml_qx)
    else:  # иначе запускаем карусель из картинок
        for line in last_5_images_qx_previous:
            if len(line[2]) > 3:
                # print(line[2])
                line2 = []
                for i in range(len(line[2])):
                    if i + 1 < len(line[2]):
                        line2.append(line[2][i + 1])
                    else:
                        line2.append(line[2][0])
                line[2] = line2
                # print('Modify:', line[2])
    # print('After:', last_5_images_qx_previous)

    # выводим картинки и подписи к ним

    canvas_2_5.delete('all')
    if len(last_5_images_qx_previous) > 0:
        for y in range(len(last_5_images_qx_previous)):
            for x in range(len(last_5_images_qx_previous[y][2])):
                canvas_2_5.create_text(75 + step_x * x, 10 + step_y * y, text=last_5_images_qx_previous[y][0], justify=CENTER, font="Verdana 12")
                canvas_2_5.create_image(80 + step_x * x, 90 + step_y * y, image=last_5_images_qx_previous[y][2][x])
    root.after(10*1000, ask_info)


ask_info()
root.mainloop()
