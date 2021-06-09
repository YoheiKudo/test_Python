import pymysql.cursors
import openpyxl as Excel

connection = pymysql.connect(host='192.168.1.231', user='kudo', password='1111',
                             db='sahashinewsystem', charset='utf8mb4',
                             cursorclass=pymysql.cursors.DictCursor)

try:
    wb = Excel.Workbook()
    ws = wb.active
    ws.title = '加硫実績'

    with connection.cursor() as cursor:
        sql = "SELECT * FROM kakouhin.df_karyujisseki"
        cursor.execute(sql)
        result = cursor.fetchall()
        ws['A1'] = 'ID'
        ws['B1'] = '加硫日'
        ws['C1'] = '直'
        ws['D1'] = '製番'
        ws['E1'] = 'ショット数'
        ws['F1'] = '型番'
        ws['G1'] = '不良コード1'
        ws['H1'] = '不良数1'
        ws['I1'] = '不良コード2'
        ws['J1'] = '不良数2'
        ws['K1'] = '不良コード3'
        ws['L1'] = '不良数3'
        ws['M1'] = '不良コード4'
        ws['N1'] = '不良数4'
        ws['O1'] = '停止コード1'
        ws['P1'] = '停止時間1'
        ws['Q1'] = '停止コード2'
        ws['R1'] = '停止時間2'
        ws['S1'] = '停止コード3'
        ws['T1'] = '停止時間3'
        ws['U1'] = '停止コード4'
        ws['V1'] = '停止時間4'
        ws['W1'] = 'ゴム配合1'
        ws['X1'] = 'ゴム配合2'
        ws['Y1'] = 'ゴム配合3'
        ws['Z1'] = 'ゴム配合4'
        ws['AA1'] = '処理日'
        i = 2
        for row in result:
            stri = str(i)
            ws['A' + stri] = row['ID']
            ws['B' + stri] = row['KARYUBI']
            ws['C' + stri] = row['TYOKU']
            ws['D' + stri] = row['SEIBAN']
            ws['E' + stri] = row['KARYU_SHOT']
            ws['F' + stri] = row['KATABAN']
            ws['G' + stri] = row['HURYO_CD1']
            ws['H' + stri] = row['HURYO_SU1']
            ws['I' + stri] = row['HURYO_CD2']
            ws['J' + stri] = row['HURYO_SU2']
            ws['K' + stri] = row['HURYO_CD3']
            ws['L' + stri] = row['HURYO_SU3']
            ws['M' + stri] = row['HURYO_CD4']
            ws['N' + stri] = row['HURYO_SU4']
            ws['O' + stri] = row['TEISI_CD1']
            ws['P' + stri] = row['TEISI_JIKAN1']
            ws['Q' + stri] = row['TEISI_CD2']
            ws['R' + stri] = row['TEISI_JIKAN2']
            ws['S' + stri] = row['TEISI_CD3']
            ws['T' + stri] = row['TEISI_JIKAN3']
            ws['U' + stri] = row['TEISI_CD4']
            ws['V' + stri] = row['TEISI_JIKAN4']
            ws['W' + stri] = row['GM1']
            ws['X' + stri] = row['GM2']
            ws['Y' + stri] = row['GM3']
            ws['Z' + stri] = row['GM4']
            ws['AA' + stri] = row['SYORIBI']
            i += 1

    ws2 = wb.create_sheet(title='検査実績')
    with connection.cursor() as cursor:
        sql = "SELECT * FROM kakouhin.df_kakou_kensa"
        cursor.execute(sql)
        result = cursor.fetchall()
        ws2['A1'] = 'ID'
        ws2['B1'] = '検査部署'
        ws2['C1'] = '作業日'
        ws2['D1'] = '製番'
        ws2['E1'] = '仕入区分'
        ws2['F1'] = '個数'
        ws2['G1'] = '手直し'
        ws2['H1'] = '不良コード1'
        ws2['I1'] = '不良数1'
        ws2['J1'] = '不良コード2'
        ws2['K1'] = '不良数2'
        ws2['L1'] = '不良コード3'
        ws2['M1'] = '不良数3'
        ws2['N1'] = '不良コード4'
        ws2['O1'] = '不良数4'
        ws2['P1'] = '作業開始時間'
        ws2['Q1'] = '作業終了時間'
        ws2['R1'] = '入力日'
        i = 2
        for row in result:
            stri = str(i)
            ws2['A' + stri] = row['ID']
            ws2['B' + stri] = row['KENSABUSYO']
            ws2['C' + stri] = row['SAGYOUBI']
            ws2['D' + stri] = row['SEIBAN']
            ws2['E' + stri] = row['SIIREKUBUN']
            ws2['F' + stri] = row['SAGYOU_KOSU']
            ws2['G' + stri] = row['TENAOSI']
            ws2['H' + stri] = row['HURYO_CD1']
            ws2['I' + stri] = row['HURYO_SU1']
            ws2['J' + stri] = row['HURYO_CD2']
            ws2['K' + stri] = row['HURYO_SU2']
            ws2['L' + stri] = row['HURYO_CD3']
            ws2['M' + stri] = row['HURYO_SU3']
            ws2['N' + stri] = row['HURYO_CD4']
            ws2['O' + stri] = row['HURYO_SU4']
            if len(str(row['S_H_TIME'])) == 1:
                row['S_H_TIME'] = '0' + str(row['S_H_TIME'])
            else:
                row['S_H_TIME'] = str(row['S_H_TIME'])

            if len(str(row['S_M_TIME'])) == 1:
                row['S_M_TIME'] = '0' + str(row['S_M_TIME'])
            else:
                row['S_M_TIME'] = str(row['S_M_TIME'])

            if len(str(row['E_H_TIME'])) == 1:
                row['E_H_TIME'] = '0' + str(row['E_H_TIME'])
            else:
                row['E_H_TIME'] = str(row['E_H_TIME'])

            if len(str(row['E_M_TIME'])) == 1:
                row['E_M_TIME'] = '0' + str(row['E_M_TIME'])
            else:
                row['E_M_TIME'] = str(row['E_M_TIME'])

            ws2['P' + stri] = str(row['S_H_TIME']) + ':' + str(row['S_M_TIME'])
            ws2['Q' + stri] = str(row['E_H_TIME']) + ':' + str(row['E_M_TIME'])
            ws2['R' + stri] = row['INPUT']
            i += 1

    wb.save('kakouhin.xlsx')
    connection.commit()

finally:
    connection.close()
