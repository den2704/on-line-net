import time  # библиотека для работы с временем
import datetime  # библиотека для работы с датой и временем
import glob  # чтобы получить список файлов
import os.path  # чтобы проверить существование файла
from operator import itemgetter
import xlwt  # библиотека для записи в файл .xls

my_period = 30  # 30 дней - диапазон просмотра результатов
period = datetime.timedelta(days=my_period)
start_date = datetime.datetime.now() - period
start_date = start_date.strftime('%Y.%m.%d')
end_date = datetime.datetime.now().strftime('%Y.%m.%d')
user_period = start_date+'-'+end_date  # 2020.09.17-2020.09.21 - Sep 17 - Sep 21
file_name_1 = '\\\Tk22001s\\data\\Documents\\Цех №1\\Сетевая\\Отчеты оптика AOI\\QX\\inspection.txt'
file_name_2 = '\\\Tk22001s\\data\\Documents\\Цех №1\\Сетевая\\Отчеты оптика AOI\\QX\\rework.txt'
header = 'Дата;Номер линии;Программа;Печатных плат всего;Количество мультилистов всего;Количество сработок мл;Кол. деф. м/л;Кол. деф. плат;Кол-во дефектов;дефект.1;д.1.шт.;дефект.2;д.2.шт.;дефект.3;д.3.шт.;Выход годных плат,%;Паяльная паста'


def inspection_rework():
    global file_name_1, file_name_2
    files = glob.glob('\\\Tk22001s\data\Documents\\Цех №1\\Сетевая\\Отчеты оптика AOI\\QX\\*.txt')
    date_1 = datetime.datetime.timestamp(datetime.datetime.now()) - my_period * 24 * 60 * 60
    line_f_insp = []; line_f_rework = []
    for i in range(10):
        line_f_insp.append("")
        line_f_rework.append("")
    with open(file_name_1, "w") as result:
        i = 0
        for file in files:
            if 'inspection' in file and file != 'inspection.txt' and os.path.getmtime(file) > date_1:
                if i != 0:
                    result.write('\n')
                i += 1
                for line in open(file, 'r'):
                    result.write(line)  # собрали inspection.txt
    with open(file_name_2, "w") as result:
        i = 0
        for file in files:
            if 'rework' in file and file != 'rework.txt' and os.path.getmtime(file) > date_1:
                if i != 0:
                    result.write('\n')
                i += 1
                for line in open(file, 'r'):
                    result.write(line)  # собрали rework.txt


def input_date(thedate):
    day_weekday = thedate.weekday()
    if day_weekday == 0:
        day_w = 'Mon'
    elif day_weekday == 1:
        day_w = 'Tue'
    elif day_weekday == 2:
        day_w = 'Wed'
    elif day_weekday == 3:
        day_w = 'Thu'
    elif day_weekday == 4:
        day_w = 'Fri'
    elif day_weekday == 5:
        day_w = 'Sat'
    else:
        day_w = 'Sun'

    m = thedate.month
    if m == 1:
        month = 'Jan'
    elif m == 2:
        month = 'Feb'
    elif m == 3:
        month = 'Mar'
    elif m == 4:
        month = 'Apr'
    elif m == 5:
        month = 'May'
    elif m == 6:
        month = 'Jun'
    elif m == 7:
        month = 'Jul'
    elif m == 8:
        month = 'Aug'
    elif m == 9:
        month = 'Sep'
    elif m == 10:
        month = 'Oct'
    elif m == 11:
        month = 'Nov'
    else:
        month = 'Dec'
   
    if len(str(thedate.day)) != 2:
        day = '0'+str(thedate.day)
    else:
        day = str(thedate.day)
    return day_w+' '+month+' '+day


def change_date(date):
    date1 = date.replace(date[:4], '')
    # print(date1[:3])
    if date1[:3] == 'Jan':
        date1 = date1.replace('Jan', '01')
    elif date1[:3] == 'Feb':
        date1 = date1.replace('Feb', '02')
    elif date1[:3] == 'Mar':
        date1 = date1.replace('Mar', '03')
    elif date1[:3] == 'Apr':
        date1 = date1.replace('Apr', '04')
    elif date1[:3] == 'May':
        date1 = date1.replace('May', '05')
    elif date1[:3] == 'Jun':
        date1 = date1.replace('Jun', '06')
    elif date1[:3] == 'Jul':
        date1 = date1.replace('Jul', '07')
    elif date1[:3] == 'Aug':
        date1 = date1.replace('Aug', '08')
    elif date1[:3] == 'Sep':
        date1 = date1.replace('Sep', '09')
    elif date1[:3] == 'Oct':
        date1 = date1.replace('Oct', '10')
    elif date1[:3] == 'Nov':
        date1 = date1.replace('Nov', '11')
    elif date1[:3] == 'Dec':
        date1 = date1.replace('Dec', '12')
    date1 = datetime.datetime.strptime(date1, '%m %d %H:%M:%S %Y')
    return date1


def f_teko():  # ищем информацию о продуктах в TEKO.csv
    global list_teko
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
    # for lt in list_teko:
    #     print(lt)


def solder_paste():  # добавляем паяльную пасту
    global result_file
    for line in result_file:
        result_file[result_file.index(line)].append('NEVO')
    # for line in result_file:
    #     if datetime.datetime.strptime(line[0], '%d.%m.%Y') < datetime.datetime.strptime('25.11.2020', '%d.%m.%Y'):
    #         result_file[result_file.index(line)].append('UNION SOLTEK G4A-SM833')
    #     else:
    #         result_file[result_file.index(line)].append('ALPHA OL107F(A)')


def write_to_data():  # записываем результаты в .xls
    global result_file, header
    line_num = 0
    book = xlwt.Workbook(encoding='utf-8')
    sheet1 = book.add_sheet('AOI_QX')
    header_list = header.split(';')
    col_num = 0
    for h in header_list:
        sheet1.write(line_num, col_num, h)
        col_num += 1
    line_num = 1
    for line_r_f in result_file:
        if int(line_r_f[3]) != 0:
            col_num = 0
            for r_f in line_r_f:
                sheet1.write(line_num, col_num, r_f)
                col_num += 1
            line_num += 1
    try:
        book.save('\\\\tk22001s\\DATA\\Documents\\Цех №1\\Сетевая\\Отчеты оптика AOI\\Data_AOI_QX.xls')
    except PermissionError:
        print("Please, close file!")
    try:
        book.save('\\\\tk22001s\\DATA\\Documents\\Цех №1\\Сетевая\\Отчеты оптика AOI\\Data_AOI_QX_2.xls')
    except PermissionError:
        print("Please, close file 2!")


period = []
if '-' in user_period:
    period = user_period.split('-')
else:
    period.append(user_period)
for i in range(len(period)):
    period[i] = datetime.datetime.strptime(period[i], '%Y.%m.%d')
user_dates = []
current_date = period[0]
if len(period) == 2:
    while current_date != period[1]+datetime.timedelta(days=1):
        user_dates.append(current_date)
        current_date = current_date+datetime.timedelta(days=1)
else:
    user_dates.append(current_date)
for i in range(len(user_dates)):
    user_dates[i] = input_date(user_dates[i])
# print(user_dates)

############ запускам обработку данных ############
result_file = []
list_teko = []
job_barcode = []
line_f_insp = [] #line_f_insp

inspection_rework()  # собрали inspection.txt и rework.txt
for i in range(10):
    line_f_insp.append('')  # заготовка строчки с данными

# 0.Assembly Name/1.Bar Code/2.Batch Code/3.DateTime/4.Index/5.Status/
# 6.Template/7.Component Ref/8.Result Code/9.Operator

split = []; jobs = []; date_job = []
with open(file_name_1) as file:  # анализируем inspection.txt
    for line in file:
        for date in user_dates:
            if date in line:
                column_l = 0
                for i in range(10):
                    line_f_insp[i] = ''
                for char in line:
                    if char == '\t':
                        column_l += 1
                    elif char != '\n':
                        line_f_insp[column_l] += char
                if line_f_insp[0].find('/') != -1:
                    split = line_f_insp[0].split('/')
                    line_f_insp[0] = split[1]
                date_job.append([date, line_f_insp[0], line_f_insp[3]])
for line in date_job:
    line[2] = change_date(line[2])
date_job = sorted(date_job, key=itemgetter(2))
for line in date_job:
    if [line[0], 2, line[1]] not in result_file:
        result_file.append([line[0], 2, line[1]])  # Дата[0];Номер линии[1];Программа[2]
# for line in result_file:
#     print(line)
with open(file_name_1) as file:  # анализируем inspection.txt
    for line in file:
        for rf in result_file:
            job_barcode.append([])
            if rf[0] in line and rf[2] in line:
                column_l = 0
                for i in range(10):
                    line_f_insp[i] = ''
                for char in line:
                    if char == '\t':
                        column_l += 1
                    elif char != '\n':
                        line_f_insp[column_l] += char
                if line_f_insp[0].find('/') != -1:
                    split = line_f_insp[0].split('/')
                    line_f_insp[0] = split[1]
                if line_f_insp[1].find('_') != -1 and line_f_insp[1] != 'NO_BARCODE_READ' and line_f_insp[1] != '0_0':
                    split = line_f_insp[1].split('_')
                    for i in range(len(split)):
                        if split[i].isdigit():
                            line_f_insp[1] = split[i]
                if line_f_insp[1] not in job_barcode[result_file.index(rf)] and line_f_insp[1].isdigit() and int(line_f_insp[1]) != 0 and len(line_f_insp[1]) > 7: #'NO_BARCODE_READ' not in line_f_insp[1]
                    job_barcode[result_file.index(rf)].append(line_f_insp[1])
f_teko()  # ищем информацию о продуктах в TEKO.csv
for line in result_file:
    for lt in list_teko:
        if line[2] == lt[0]:
            line.append(lt[1])
for line in result_file:
    # print(line)
    # print(line[3])
    # print(job_barcode[result_file.index(line)])
    line[3] = int(line[3])*(len(job_barcode[result_file.index(line)]))  # Печатных плат всего[3]
    line.append(len(job_barcode[result_file.index(line)]))  # Количество м/л всего[4]
# for line in result_file:
#     print(line)
# input()
job_number = 0
job_name = ""
line_f_r4 = []
#0.Assembly Name/1.Bar Code/2.Batch Code/3.DateTime/4.Index/5.Status/6.Template/7.Component Ref/8.Result Code/
# 9.Operator
line_f_rework = [] #line_f_rework
defect = ''
num = 0
line_f_r4_job = []
for i in range(10):
    line_f_rework.append('')
job_barcode_rework = []
job_barcode_def = []
job_barcode_def_boards = []
defects = []
defect_name = []
quantity = []
for rf in result_file:
    job_barcode_rework.append([])
    job_barcode_def.append([])
    job_barcode_def_boards.append([])
    defects.append([])
    defect_name.append([])
    quantity.append([])

with open(file_name_2) as file:  # анализируем rework.txt
    for line in file:
        # print(line)
        for rf in result_file:
            if rf[0] in line and rf[2] in line:
                column_l = 0
                for i in range(10):
                    line_f_rework[i] = ''
                for char in line:
                    if char == '\t':
                        column_l += 1
                    elif char != '\n':
                        line_f_rework[column_l] += char
                if line_f_rework[0].find('/') != -1:
                    split = line_f_rework[0].split('/')
                    line_f_rework[0] = split[1]
                if line_f_rework[1].find('_') != -1 and line_f_rework[1] != 'NO_BARCODE_READ' and line_f_rework[1] != '0_0':
                        split = line_f_rework[1].split('_')
                        for i in range(len(split)):
                            if split[i].isdigit():
                                line_f_rework[1] = split[i]
                if line_f_rework[1] not in job_barcode_rework[result_file.index(rf)] and line_f_rework[1].isdigit() and int(line_f_rework[1]) != 0 and len(line_f_rework[1]) > 7:
                    job_barcode_rework[result_file.index(rf)].append(line_f_rework[1])
                if line_f_rework[1] not in job_barcode_def[result_file.index(rf)] and line_f_rework[8] != 'False Fail' and line_f_rework[1].isdigit() and int(line_f_rework[1]) != 0 and len(line_f_rework[1]) > 7:
                    job_barcode_def[result_file.index(rf)].append(line_f_rework[1])
                if 'False Fail' not in line_f_rework[8]:  # если была не ложная сработка
                    split = []
                    split = line_f_rework[7].split('.')
                    split[1] = split[1].replace('Board', '')
                    if len(split) == 2:
                        # print(5)
                        CRD = split.pop(1)
                        split.append(1)
                        split.append(CRD)
                    # print(split)
                    temp = [int(line_f_rework[1]), int(split[1])]
                    if temp not in job_barcode_def_boards[result_file.index(rf)]:
                        job_barcode_def_boards[result_file.index(rf)].append(temp)
                    if temp in job_barcode_def_boards[result_file.index(rf)]:
                        defects[result_file.index(rf)].append(split[2]+line_f_rework[6]+'_'+line_f_rework[8])
            defects[result_file.index(rf)] = sorted(defects[result_file.index(rf)])

for rf in result_file:  # собираем дефекты в список: дефект - количество
    for defect in defects[result_file.index(rf)]:
        if defect in defect_name[result_file.index(rf)]:
            quantity[result_file.index(rf)][defect_name[result_file.index(rf)].index(defect)] += 1
        else:
            defect_name[result_file.index(rf)].append(defect)
            quantity[result_file.index(rf)].append(1)
defects_all = []
for i in range(len(result_file)):  # для каждой строчки в файле результатов - есть  все дефекты
    defects_all.append([])
    for j in range(len(defect_name[i])):
        defects_all[i].append([defect_name[i][j], quantity[i][j]])

for rf in result_file:
    defects_all[result_file.index(rf)] = pcbguid_allbarcode_pcbid_sorted = sorted(defects_all[result_file.index(rf)], key=itemgetter(0), reverse=False)
    defects_all[result_file.index(rf)] = pcbguid_allbarcode_pcbid_sorted = sorted(defects_all[result_file.index(rf)], key=itemgetter(1), reverse=True)
    defects_del = []
    for k in range(len(defects_all[result_file.index(rf)])):
        if k > 2:
            defects_del.append(defects_all[result_file.index(rf)][k])
    for defect in defects_del:
        defects_all[result_file.index(rf)].remove(defect)
    rf.append(len(job_barcode_rework[result_file.index(rf)]))  # Количество сработок м/л[5]
    rf.append(len(job_barcode_def[result_file.index(rf)]))  # Кол. деф. м/л[6]
    rf.append(len(job_barcode_def_boards[result_file.index(rf)]))  # Кол. деф. плат[7]
    rf.append(len(defects[result_file.index(rf)]))  # Кол-во дефектов[8]
    for k in range(len(defects_all[result_file.index(rf)])):
        rf.append(defects_all[result_file.index(rf)][k][0])  # дефект.1[9] дефект.2[11] дефект.3[13]
        rf.append(defects_all[result_file.index(rf)][k][1])  # д.1.шт.[10] д.2.шт.[12] д.3.шт.[14]
    if len(defects_all[result_file.index(rf)]) == 0:
        for j in range(3):
            rf.append('')
            rf.append(0)
    if len(defects_all[result_file.index(rf)]) == 1:
        for j in range(2):
            rf.append('')
            rf.append(0)
    if len(defects_all[result_file.index(rf)]) == 2:
        rf.append('')
        rf.append(0)
# записали дефекты

for rf in result_file:
    try:
        good = ((rf[3]-rf[7])/rf[3])*100
    except ZeroDivisionError:
        good = 0
    good = float('{:.1f}'.format(good))
    rf.append(good)  # Выход годных плат,%[15]
solder_paste()  # Паяльная паста[16]
print('\nCyberOptics QX500')
print('Dates,'+user_period)
print(header)
for rf in result_file:
    print(rf)
now = datetime.datetime.now()
print(now.strftime('%d-%m-%Y %H:%M'))
file_name_3 = '\\\Tk22001s\data\Documents\\Цех №1\\Сетевая\\Отчеты оптика AOI\\Data_AOI_QX_30d_1h.csv'
file_data = []
file_data.append(header+'\n')
for i in range(len(result_file)):
    for j in range(len(result_file[i])):
        result_file[i][j] = str(result_file[i][j])
for rf in result_file:
    file_data.append(';'.join(rf) + '\n')
with open(file_name_3, 'w') as file3:
    file3.write(''.join(file_data))
write_to_data()
