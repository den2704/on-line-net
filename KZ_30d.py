import datetime
import pyodbc
from operator import itemgetter
import time
import xlwt
from PIL import Image, ImageTk
from io import BytesIO

header = 'Дата;Номер линии;Программа;Печатных плат всего;Количество мультилистов всего;Количество сработок мл;Кол. деф. м/л;Кол. деф. плат;Кол-во дефектов;дефект.1;д.1.шт.;дефект.2;д.2.шт.;дефект.3;д.3.шт.;Выход годных плат,%;Паяльная паста'


def job_name_m(name):
    second = 0
    job_name = ''
    if str(name) == 'None':
        return
    for char in str(name):
        if char == '\\':
            second += 1
            continue
        if char == '.':
            break
        if second == 2:
            job_name += char
    return job_name


def all_pcb_ml():
    global result_file, pcbguid, barcode, server, username, password
    database1 = 'KY_AOI'
    for line in result_file:
        date1_all = datetime.datetime.strptime(line[0], '%d.%m.%Y')
        date2_all = datetime.datetime.strptime(line[0], '%d.%m.%Y') + datetime.timedelta(days=1)
        aoi_result = []
        pcbguid_allbarcode_pcbid = []
        conn_result = pyodbc.connect(
            'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server + ';DATABASE=' + database1 + ';UID=' + username + ';PWD=' + password)
        cursor_result = conn_result.cursor()
        cursor_result.execute(
            "SELECT PCBGUID,ALLBarCode,PCBID,ArrayCnt FROM dbo.TB_AOIPCB WHERE ?<EndDateTime and EndDateTime<? and JobFileIDShare=?",
            date1_all, date2_all, line[2])
        for row in cursor_result:
            if row not in aoi_result:
                aoi_result.append(row)
        cursor_result.close()
        conn_result.close()
        if len(aoi_result) == 0:
            continue
        for i in range(len(aoi_result)):
            temp = [aoi_result[i][0], aoi_result[i][1], aoi_result[i][2]]
            if temp not in pcbguid_allbarcode_pcbid and aoi_result[i][1].isdigit() and int(aoi_result[i][1]) != 0 and len(aoi_result[i][1]) > 7 and len(aoi_result[i][1]) < 10:
                pcbguid_allbarcode_pcbid.append(temp)
        if len(pcbguid_allbarcode_pcbid) != 0:
            pcbguid_allbarcode_pcbid_sorted = sorted(pcbguid_allbarcode_pcbid, key=itemgetter(1, 2), reverse=True)
            previous = pcbguid_allbarcode_pcbid_sorted[0]
            pcbguid_allbarcode_pcbid = []
            if pcbguid_allbarcode_pcbid_sorted[0][1] not in barcode:
                pcbguid_allbarcode_pcbid.append(pcbguid_allbarcode_pcbid_sorted[0])
                pcbguid.append(pcbguid_allbarcode_pcbid_sorted[0][0])
                barcode.append(pcbguid_allbarcode_pcbid_sorted[0][1])
            for i in range(1, len(pcbguid_allbarcode_pcbid_sorted)):
                if int(pcbguid_allbarcode_pcbid_sorted[i][1]) == int(previous[1]):
                    continue
                else:
                    if pcbguid_allbarcode_pcbid_sorted[i][1] not in barcode:
                        pcbguid_allbarcode_pcbid.append(pcbguid_allbarcode_pcbid_sorted[i])
                        pcbguid.append(pcbguid_allbarcode_pcbid_sorted[i][0])
                        barcode.append(pcbguid_allbarcode_pcbid_sorted[i][1])
                previous = pcbguid_allbarcode_pcbid_sorted[i]
        result_file[result_file.index(line)].append(len(pcbguid_allbarcode_pcbid) * aoi_result[0][3])  # pp
        result_file[result_file.index(line)].append(len(pcbguid_allbarcode_pcbid))  # ml
        # print(pcbguid)


def defects_ml():
    global result_file, server, username, password, pcbguid, barcode
    database1 = 'KY_AOI'
    for line in result_file:
        date1_def = datetime.datetime.strptime(line[0], '%d.%m.%Y')
        date2_def = datetime.datetime.strptime(line[0], '%d.%m.%Y') + datetime.timedelta(days=1)
        def_result = []
        conn_def = pyodbc.connect(
            'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server + ';DATABASE=' + database1 + ';UID=' + username + ';PWD=' + password)
        cursor_def = conn_def.cursor()
        cursor_def.execute(
            "SELECT PCBGUID,PCBResultBefore,PCBResultAfter FROM dbo.TB_AOIPCB WHERE ?<EndDateTime and EndDateTime<? and JobFileIDShare=?",
            date1_def, date2_def, line[2])
        for row in cursor_def:
            # print(row)
            if row[0] in pcbguid and row[1] != 11000000:
                if row not in def_result:
                    def_result.append(row)
        cursor_def.close()
        conn_def.close()
        def_13_13 = []
        def_13_12 = []
        for def_r in def_result:
            if def_r[1] == 13000000 and def_r[2] == 13000000:
                def_13_13.append(def_r)
            elif def_r[1] == 13000000 and def_r[2] == 12000000:
                def_13_12.append(def_r)
        result_file[result_file.index(line)].append(len(def_result))  # def_result
        result_file[result_file.index(line)].append(len(def_13_13))  # ml 13 13 def


def defects_board():
    global result_file, server, username, password, result_dbname, DefectType, pcbguid, barcode, det_result
    for line in result_file:
        det_result = []; dbn_result = []; dba_result = []; boards = []
        qdc_result = []; def_comp = []; def_comp_q = []; dc_d = []
        database = 'KY_AOI'
        date1_db = datetime.datetime.strptime(line[0], '%d.%m.%Y')
        date2_db = datetime.datetime.strptime(line[0], '%d.%m.%Y') + datetime.timedelta(days=1)
        conn_aoi = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
        cursor_aoi = conn_aoi.cursor()
        cursor_aoi.execute("SELECT ResultDBName FROM dbo.TB_AOIPCB WHERE ?<EndDateTime and EndDateTime<?", date1_db, date2_db)
        result_dbname = []
        for row in cursor_aoi:
            result_dbname.append(row[0])
            break
        cursor_aoi.close()
        conn_aoi.close()
        if len(result_dbname) > 1:
            print('Try to solve the problem with 2 databases')
            print(result_dbname)
        dbname = result_dbname[0]
        conn_db = pyodbc.connect(
            'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server + ';DATABASE=' + dbname + ';UID=' + username + ';PWD=' + password)
        cursor_db = conn_db.cursor()
        cursor_db.execute(
            "SELECT PCBGUID FROM dbo.TB_AOIResult WHERE ?<EndDateTime and EndDateTime<? and JobFileIDShare=?", date1_db, date2_db, line[2])
        for row in cursor_db:
            # print(row)
            if row[0] in pcbguid:
                dbn_result.append(row[0])
                pcbguid.remove(row[0])
        cursor_db.close()
        conn_db.close()
        conn_dba = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server + ';DATABASE=' + dbname + ';UID=' + username + ';PWD=' + password)
        cursor_dba = conn_dba.cursor()
        cursor_dba.execute("SELECT PCBGUID,ArrayIndex,ResultAfter FROM dbo.TB_AOIDefect WHERE ResultAfter=13000000")
        for dba in cursor_dba:
            if dba[0] in dbn_result:
                dba_result.append(dba)
        cursor_dba.close()
        conn_dba.close()
        for dba in dba_result:
            if [dba[0], dba[1]] not in boards:
                boards.append([dba[0], dba[1]])
        result_file[result_file.index(line)].append(len(boards))  # def boards
        conn_det = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server + ';DATABASE=' + dbname + ';UID=' + username + ';PWD=' + password)
        cursor_det = conn_det.cursor()
        cursor_det.execute("SELECT PCBGUID,ComponentGUID,Defect FROM dbo.TB_AOIDefectDetail")
        for det in cursor_det:
            if det[0] in dbn_result and det[2] != 12000000:
                if det not in det_result:
                    det_result.append(det)
                    det_comp.append(det[1])
        cursor_det.close()
        conn_det.close()
        result_file[result_file.index(line)].append(len(det_result))  # defects
        conn_qdc = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server + ';DATABASE=' + dbname + ';UID=' + username + ';PWD=' + password)
        cursor_qdc = conn_qdc.cursor()
        cursor_qdc.execute("SELECT PCBGUID,ComponentGUID,uname,PartNumber,ResultAfter FROM dbo.TB_AOIDefect WHERE ResultBefore>11000000")
        for qdc in cursor_qdc:
            if qdc[0] in dbn_result and qdc[4] != 12000000:
                qdc_result.append(qdc)
        cursor_qdc.close()
        conn_qdc.close()
        for i in range(len(det_result)):
            for j in range(len(qdc_result)):
                if det_result[i][1] == qdc_result[j][1]:
                    def_comp.append(
                        [det_result[i][1], det_result[i][2], qdc_result[j][2], qdc_result[j][3], qdc_result[j][4]])
        for one in def_comp:
            for two in DefectType:
                if one[1] == two[0]:
                    one[1] = two[1]
        for one in def_comp:
            def_comp_q.append([one[2], one[1]])
        for q in def_comp_q:
            if [q[0] + ' ' + q[1], 0] not in dc_d:
                dc_d.append([q[0] + ' ' + q[1], 0])
        for q in def_comp_q:
            for dc in dc_d:
                if q[0] + ' ' + q[1] == dc[0]:
                    dc[1] += 1
        dc_d = sorted(dc_d)
        dc_d = sorted(dc_d, key=itemgetter(1), reverse=True)
        if len(dc_d) >= 3:
            result_file[result_file.index(line)].append(dc_d[0][0])  # top 3 defects
            result_file[result_file.index(line)].append(dc_d[0][1])
            result_file[result_file.index(line)].append(dc_d[1][0])
            result_file[result_file.index(line)].append(dc_d[1][1])
            result_file[result_file.index(line)].append(dc_d[2][0])
            result_file[result_file.index(line)].append(dc_d[2][1])
        elif len(dc_d) == 2:
            result_file[result_file.index(line)].append(dc_d[0][0])  # top 3 defects
            result_file[result_file.index(line)].append(dc_d[0][1])
            result_file[result_file.index(line)].append(dc_d[1][0])
            result_file[result_file.index(line)].append(dc_d[1][1])
            result_file[result_file.index(line)].append('')
            result_file[result_file.index(line)].append(0)
        elif len(dc_d) == 1:
            result_file[result_file.index(line)].append(dc_d[0][0])  # top 3 defects
            result_file[result_file.index(line)].append(dc_d[0][1])
            result_file[result_file.index(line)].append('')
            result_file[result_file.index(line)].append(0)
            result_file[result_file.index(line)].append('')
            result_file[result_file.index(line)].append(0)
        elif len(dc_d) == 0:
            result_file[result_file.index(line)].append('')
            result_file[result_file.index(line)].append(0)
            result_file[result_file.index(line)].append('')
            result_file[result_file.index(line)].append(0)
            result_file[result_file.index(line)].append('')
            result_file[result_file.index(line)].append(0)


def write_to_data():
    global result_file, header
    file_data = []
    file_name = '\\\\tk22001s\\DATA\\Documents\\Цех №1\\Сетевая\\Отчеты оптика AOI\\Data_AOI_KZ_30d_1h.csv'
    # file_name2 = '\\\\tk22001s\\DATA\\Documents\\Цех №1\\Сетевая\\Отчеты оптика AOI\\Stat_data_AOI.csv'
    for i in range(len(result_file)):
        for j in range(len(result_file[i])):
            result_file[i][j] = str(result_file[i][j])
    file_data.append(header+'\n')
    for line_r_f in result_file:  # add new days
        if line_r_f[3] != '0':
                file_data.append(';'.join(line_r_f) + '\n')
    with open(file_name, 'w') as file:
        file.write(''.join(file_data))
    # with open(file_name2, 'w') as file2:
    #     file2.write(''.join(file_data))
    line_num = 0
    book = xlwt.Workbook(encoding='utf-8')
    sheet1 = book.add_sheet('AOI_KZ')
    header_list = header.split(';')
    col_num = 0
    for h in header_list:
        sheet1.write(line_num, col_num, h)
        col_num += 1
    line_num = 1
    for line_r_f in result_file:
        col_num = 0
        for r_f in line_r_f:
            sheet1.write(line_num, col_num, r_f)
            col_num += 1
        line_num += 1
    try:
        book.save('\\\\tk22001s\\DATA\\Documents\\Цех №1\\Сетевая\\Отчеты оптика AOI\\Data_AOI_KZ.xls')
    except PermissionError:
        print("Please, close file!")
    try:
        book.save('\\\\tk22001s\\DATA\\Documents\\Цех №1\\Сетевая\\Отчеты оптика AOI\\Data_AOI_KZ_2.xls')
    except PermissionError:
        print("Please, close file 2!")


def array_total_comp():
    global result_file, server, username, password
    database = 'KY_AOI'
    for line in result_file:
        array = []
        total_comp = []
        date1_all = datetime.datetime.strptime(line[0], '%d.%m.%Y')
        date2_all = datetime.datetime.strptime(line[0], '%d.%m.%Y') + datetime.timedelta(days=1)
        conn_array = pyodbc.connect(
            'DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
        cursor_array = conn_array.cursor()
        cursor_array.execute(
            "SELECT ArrayCnt,PCBTotalComp FROM dbo.TB_AOIPCB WHERE ?<EndDateTime and EndDateTime<? and JobFileIDShare=?",
            date1_all, date2_all, line[2])
        for row in cursor_array:
            if row[0] not in array:
                array.append(row[0])
                total_comp.append(row[1])
        cursor_array.close()
        conn_array.close()
        result_file[result_file.index(line)].append(max(array))
        result_file[result_file.index(line)].append(max(total_comp))


def good():
    global result_file
    for line in result_file:
        if line[3] != 0:
            good = (line[3]-line[7])*100/line[3]
        else:
            good = 0
        good = float('{:.1f}'.format(good))
        result_file[result_file.index(line)].append(good)


def solder_paste():
    global result_file
    date_25_11_2020 = datetime.datetime.strptime('25.11.2020', '%d.%m.%Y')
    date_10_12_2020 = datetime.datetime.strptime('10.12.2020', '%d.%m.%Y')
    date_29_12_2020 = datetime.datetime.strptime('29.12.2020', '%d.%m.%Y')
    for line in result_file:
        if datetime.datetime.strptime(line[0], '%d.%m.%Y') < date_25_11_2020:
            result_file[result_file.index(line)].append('UNION SOLTEK G4A-SM833')
        elif datetime.datetime.strptime(line[0], '%d.%m.%Y') < date_10_12_2020:
            result_file[result_file.index(line)].append('ALPHA OL107F(A)')
        elif datetime.datetime.strptime(line[0], '%d.%m.%Y') < date_29_12_2020:
            result_file[result_file.index(line)].append('AIM REL61M8')
        else:
            result_file[result_file.index(line)].append('NEVO')


def images():
    global det_result, det_comp, ImageDBName
    print(ImageDBName)
    # pil_image = Image.open(BytesIO(Image1[0]))
    # image = ImageTk.PhotoImage(pil_image)
    # label.config(image=image, text='')
    # label.image = image


server = 'AP-SL-01691R\MSSQL2016'
username = 'user3'
password = 'public3'
database = 'KY_AOI'
conn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+';UID='+username+';PWD='+password)
cursor = conn.cursor()
cursor.execute("SELECT DefectType,DefectName,AcceptablePass FROM dbo.TB_DefectType")
DefectType = []
for row in cursor:
    DefectType.append([row[0], row[1]])
cursor.close()
conn.close()
# print(DefectType)
# with open('DefectType.csv', 'w') as file:
#     for line in DefectType:
#         file.write(line[1]+'\n')
# input()
det_result = []
det_comp = []
update_list = []
while True:
    print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'))
    # current_date = datetime.datetime.now().strftime('%Y-%m-%d')
    # current_H = int(datetime.datetime.now().strftime('%H'))
    # if 7 < current_H < 12 or current_H == 7:
    #     check_H = 7
    # elif 12 < current_H < 15 or current_H == 12:
    #     check_H = 12
    # else:
    #     check_H = 0
    # if [current_date, check_H, 1] not in update_list and check_H == 7:
    #     update_list.append([current_date, check_H, 0])
    # if [current_date, check_H, 1] not in update_list and check_H == 12:
    #     update_list.append([current_date, check_H, 0])
    # summa = 0
    # for ul in update_list:
    #     summa += ul[2]
    # if len(update_list) != summa:
    pcbguid = []; barcode = []; result_file = []
    conn_date = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
    cursor_date = conn_date.cursor()
    cursor_date.execute("SELECT EndDateTime FROM dbo.TB_AOIPCB")
    EndDateTime = []
    for row in cursor_date:
        EndDateTime.append(row[0].strftime("%Y-%m-%d"))
    cursor_date.close()
    conn_date.close()
    EndDateTime = sorted(EndDateTime, reverse=True)
    EndDateTime_3 = []
    data_num = 0
    for data in EndDateTime:
        if data not in EndDateTime_3 and data_num < 30:
            EndDateTime_3.append(data)
            data_num += 1
    date1 = datetime.datetime.strptime(EndDateTime_3[len(EndDateTime_3)-1], '%Y-%m-%d')
    date2 = datetime.datetime.strptime(EndDateTime_3[0], '%Y-%m-%d')+datetime.timedelta(days=1)
    print(date1, date2)
    conn_aoi = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
    cursor_aoi = conn_aoi.cursor()
    cursor_aoi.execute("SELECT ResultDBName, ImageDBName FROM dbo.TB_AOIPCB WHERE ?<EndDateTime and EndDateTime<?", date1, date2)
    result_dbname = []
    ImageDBName = []
    for row in cursor_aoi:
        if row[0] not in result_dbname:
            result_dbname.append(row[0])
            ImageDBName.append(row[1])
    cursor_aoi.close()
    conn_aoi.close()
    result_dbname = sorted(result_dbname)
    # print(result_dbname)
    # print(len(result_dbname))
    conn_all_job = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER=' + server + ';DATABASE=' + database + ';UID=' + username + ';PWD=' + password)
    cursor_all_job = conn_all_job.cursor()
    cursor_all_job.execute("SELECT EndDateTime,JobFileIDShare FROM dbo.TB_AOIPCB WHERE ?<EndDateTime and EndDateTime<?", date1, date2)
    for row in cursor_all_job:
        if [row[0], 1, row[1]] not in result_file and row[1] != None:
            result_file.append([row[0], 1, row[1]])
    cursor_all_job.close()
    conn_all_job.close()
    result_file_m = sorted(result_file, key=itemgetter(0))
    for row in result_file_m:
        row[0] = row[0].strftime("%d.%m.%Y %H:%M")
    for line in result_file_m:
        s = line[0][10:]
        line[0] = line[0].replace(s, '')
    result_file = []
    for line in result_file_m:
        if line not in result_file:
            result_file.append(line)
    all_pcb_ml()
    defects_ml()
    defects_board()
    # array_total_comp()
    good()
    solder_paste()
    # images()
    if '\\' in line[2]:
        for line in result_file:
            line[2] = job_name_m(line[2])
    write_to_data()
    print()
    print(header)
    for line in result_file:
        print(line)
    # update_list[len(update_list)-1][2] = 1
    time.sleep(60*60)
