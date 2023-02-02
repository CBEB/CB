import xlwings as xw
import requests
import akshare as ak
import pandas as pd
import time

data = ak.bond_cov_comparison()
#print(data)
#print (data.iloc[80])

def getValueOnLine(cb_code):
    #getv()
    #stock_zh_a_hist_df = ak.stock_zh_a_hist(symbol="000001", period="daily", start_date="20170301", end_date='20210907',adjust="")
    #print(stock_zh_a_hist_df)

    #ddd = ak.bond_zh_hs_cov_daily(symbol="sz128100")
    #ddd =ak.bond_zh_hs_cov_spot()

    #data.iloc[-1]  # 选取DataFrame最后一行，返回的是Series
    #data.iloc[-1:]  # 选取DataFrame最后一行，返回的是DataFrame
    #print(ddd)
    #print(ddd.iloc[-1].close  )
    # print(ddd.iloc[-1]  )


    val = data[data['转债代码']==str(cb_code)].iloc[-1]
    value=val.转债最新价
    #print(value)
    if(value=='-'):
        return 0.0
    return  float(value)
    #print(type(value))


    # 根据日期定义文件名字
    #current_time = time.strftime('%Y-%m-%d', time.localtime())
    #file_name = current_time + ".xlsx"

    #writer = pd.ExcelWriter(file_name)
    #data.to_excel(writer, "sheet1")
    #writer.save()
    #print("数据保存成功")

def xls_func():
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    path = r"D:\CB\2022"
    wb = app.books.open(path + r'\CB-M-week-lowPrice.xlsx')
    """工作表"""
    # sht = wb.sheets.active                 # 获取当前活动的工作表
    # sht = wb.sheets[0]                     # 按索引获取工作表
    # sht = wb.sheets['Sheet1']              # 按表名获取工作表
    # sht1 = wb.sheets.add()                 # 新建工作表，默认新建的放在最前面。
    # sht1 = wb.sheets.add('新建工作表', after=sht)   # 新建工作表，放在sht工作表后面。

    sheet = wb.sheets['cb']
    #sheet.range('A5').value = ['小兰', 23, '女']
    #sheet.range('B3').value = ['小兰b5', 12, '女']


    """ 读取单元格 """
    #b3 = sheet.range('d3')
    # 获取 b3 中的值
    #v = b3.value
    #print("v=",v)
    print('\n===========update value===============')
    # 也可以根据行列号读取
    list_value = sheet.range('D3:L300').value
    for i in range(len(list_value)):
        start_value = list_value[i][8]
        #print(type(start_value))
        if type(start_value)==float:
            if  start_value >0:
                #print("value====",list_value[i][0])
                cb_code = int(list_value[i][0])
                name1 = list_value[i][1]
                val= getValueOnLine(cb_code)
                print("cb:",cb_code, name1,"total:",int(start_value),'现价:',val)
                #save
                sheet.range('J'+str(i+3)).value = val
    # 读取一段区间内的值
    #a1_c4_value = sht.range('a1:c4').options(ndim=2).value       # 加上 option 读取二维的数据
    #a1_c4_value = sht.range((1,1),(4,3)).options(ndim=2).value   # 和上面读取的内容一样。
    """ 写入 就是把值赋值给读取的单元格就可以了, 行，列"""
    #sheet.range(3,7).value = 'b79'

    wb.save()
    wb.close()
    app.kill()


def select_candidate():
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    path = r"D:\CB\2022"
    wb = app.books.open(path + r'\CB-M-week-lowPrice.xlsx')
    """工作表"""
    # sht = wb.sheets.active                 # 获取当前活动的工作表
    # sht = wb.sheets[0]                     # 按索引获取工作表
    # sht = wb.sheets['Sheet1']              # 按表名获取工作表
    # sht1 = wb.sheets.add()                 # 新建工作表，默认新建的放在最前面。
    # sht1 = wb.sheets.add('新建工作表', after=sht)   # 新建工作表，放在sht工作表后面。

    sheet = wb.sheets['cb']
    #sheet.range('A5').value = ['小兰', 23, '女']
    #sheet.range('B3').value = ['小兰b5', 12, '女']


    """ 读取单元格 """
    #b3 = sheet.range('d3')
    # 获取 b3 中的值
    #v = b3.value
    #print("v=",v)
    print('\n===========candidate==================')
    # 也可以根据行列号读取
    list_value = sheet.range('D3:L300').value
    for i in range(len(list_value)):
        start_value = list_value[i][8]
        #print(type(start_value))
        if type(start_value)==float:
            if start_value > 0:
               total= int(start_value)
               cb_code = int(list_value[i][0])
               name1 = list_value[i][1]
               val = getValueOnLine(cb_code)
               if(total==10 and val>108.2):             continue
               elif (total == 20 and val > 106.6):      continue
               elif (total == 30 and val > 105.1):      continue
               elif (total == 40 and val > 103.6):      continue
               elif (total == 50 and val > 102.2):      continue
               elif (total == 60 and val > 100.87):      continue
               elif (total == 70 and val > 99.56):      continue
               elif (total == 80 and val > 98.2):      continue
               elif (total == 90 and val > 97.2):      continue
               elif (total == 100 and val > 96.3):      continue
               elif (total == 110 and val > 95.31):      continue
               else:
                   print("cb:", cb_code, name1, "total:", total, '现价:', val)

    # 读取一段区间内的值
    #a1_c4_value = sht.range('a1:c4').options(ndim=2).value       # 加上 option 读取二维的数据
    #a1_c4_value = sht.range((1,1),(4,3)).options(ndim=2).value   # 和上面读取的内容一样。
    """ 写入 就是把值赋值给读取的单元格就可以了, 行，列"""
    #sheet.range(3,7).value = 'b79'

    #wb.save()
    wb.close()
    app.kill()


def getNegative():
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    path = r"D:\CB\2022"
    wb = app.books.open(path + r'\CB-M-week-lowPrice.xlsx')
    sheet = wb.sheets['cb']
    print('\n===========negative==================')
    count = 0
    list_value = sheet.range('D3:L300').value
    for i in range(len(list_value)):
        total_value = list_value[i][8]
        #print(type(start_value))
        if type(total_value)==float:
            if total_value > 0:
               total= int(total_value)
               cb_code = int(list_value[i][0])
               name1 = list_value[i][1]
               val = getValueOnLine(cb_code)
               if(val<list_value[i][7]):
                   print("cb:", cb_code, name1, "total:", total, '现价:', val,'成本',list_value[i][7])
                   count=count+1
    print('负收益总数:',count)
    #wb.save()
    wb.close()
    app.kill()

def getPosition():
    app = xw.App(visible=True, add_book=False)
    app.display_alerts = False
    app.screen_updating = False

    path = r"D:\CB\2022"
    wb = app.books.open(path + r'\CB-M-week-lowPrice.xlsx')
    sheet = wb.sheets['cb']
    count = 0
    list_value = sheet.range('D3:L300').value
    #wb.save()
    wb.close()
    app.kill()
    return  list_value

g_list_value = getPosition()

def find_num_in_xls(code):
    #print(g_list_value)
    for i in range(len(g_list_value)):
            total_value = g_list_value[i][8]
            if type(total_value) == float:
                if total_value > 0:
                    if(int(code) == int(g_list_value[i][0])):
                        return total_value
    return 0
'''
序号            379
转债代码       128062
转债名称         亚药转债
转债最新价      104.45
转债涨跌幅       -0.14
正股代码       002370
正股名称         亚太药业
正股最新价        5.12
正股涨跌幅       -0.58
转股价           8.5
转股价值      60.1176
转股溢价率       73.73
纯债溢价率       -4.69
回售触发价        5.95
强赎触发价       11.05
到期赎回价       115.0
纯债价值      109.578
开始转股日    20191009
上市日期     20190424
申购日期     20190402
'''
def find_low110():
    print('\n===========low 110==================')
    for index, row in data.iterrows():
        try:
            if(float(row['转债最新价'])<110 and float(row['转债最新价'])>50):
                if(find_num_in_xls(row['转债代码']) == 0):
                    print(row['转债代码'], row['转债名称'], row['转债最新价'])
        except Exception:
            pass
    #value=val.转债最新价
    #print(val)

if __name__ == '__main__':
    xls_func()
    select_candidate()
    getNegative()
    find_low110()