#!/usr/bin/env python3 
# -*- coding: utf-8 -*-

import openpyxl as xl
import publicFun
import os
import time

def main():
    coon = publicFun.connect_db()
    cursor = coon.cursor()
    sql = 'select id,stu_id,edu_email,edu_pwd from email_detail where tag=0 '
    result = cursor.execute(sql)
    data = cursor.fetchall()

    for detail in data:
        detail = list(detail)
        print(detail)
        write_excel_file('./',detail)
        sql_up = 'update email_detail set tag=1 where id = '+ str(detail[0])
        result1 = cursor.execute(sql_up)
        if result1:
            coon.commit()
        else:
            print("状态修改失败")

    cursor.close()
    coon.close()

def write_excel_file(folder_path,date):
    now = time.strftime("%d_%m_%Y")
    result_path = os.path.join(folder_path, "stu_edu_db_%s.xlsx"%now)
    print(result_path)
    headers = ['stu_id', 'edu_email', 'edu_pwd']
    stu_edu = []
    stu_edu.append(date[1])
    stu_edu.append(date[2])
    stu_edu.append(date[3])
    #print('***** 开始写入excel文件 ' + result_path + ' ***** \n')
    if os.path.exists(result_path):
        #print('***** excel已存在，在表后添加数据 ' + result_path + ' ***** \n')
        workbook = xl.load_workbook(result_path)
        sheet = workbook.active
        sheet.append(stu_edu)
        workbook.save(result_path)
    else:
        #print('***** excel不存在，创建excel ' + result_path + ' ***** \n')
        workbook = xl.Workbook()
        workbook.save(result_path)
        sheet = workbook.active
        sheet.append(headers)
        sheet.append(stu_edu)
        workbook.save(result_path)
    print('***** 写入Excel文件 ' + result_path + ' ***** \n')


if __name__ == '__main__':
    main()

