
import xlrd

def read_excel():
    # 打开文件
    workbook = xlrd.open_workbook('./stu_edu.xlsx')
    sheet = workbook.sheet_by_index(0)
    row_num = sheet.nrows
    col_num = sheet.ncols

    data1 = []
    data2 = []
    for i in range(1,row_num):
        data = []
        for j in range(col_num):
            data.append(sheet.cell(i,j).value)
        data1.append(data)
    data2.append(row_num)
    data2.append(col_num)
    data2.append(data1)
    #print(data1)
    return data2
read_excel()