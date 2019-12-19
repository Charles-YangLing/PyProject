import xlrd #读取
import xlwt #写入
import xlutils #多种操作
from xlutils.copy import copy
from xlutils.filter import process, XLRDReader, XLWTWriter
# import xlsxwriter
import time
import  datetime


print('提示:使用前请将本程序与要执行的文件放置同一文件夹下 \n')
date = str(input("请输入开始日期，比如2019年11月20日开始请输入：20191120 \n"))
#date = '20191126'
#name = '羽毛球经费周统计表11月14日-11月20日活动费明细(1)(1).xls'
name = str(input("请输入完整文件名称，包括后缀名(后缀名必须为.xls类型，不是的话请先转成此类型)： \n"))

def yzdata(date):
    if len(date) > 8 or len(date) <0:
        print('输入格式有误，请重新输入')
        return False
    else:
        yyyy = int(date[0:4])
        mm = int(date[4:6])
        dd = int(date[6:8])
        if mm > 12 or mm < 1:
            print('输入月份格式有误，请重新输入')
            return False
        if dd > 31 or dd < 1:
            print('输入日期格式有误，请重新输入')
            return False
        print('开始日期为'+date[0:4]+'年'+date[4:6]+'月'+date[6:8]+'日，程序开始运行')
        return True

def string_toDatetime(st):
   # print("2.把字符串转成datetime: ", datetime.datetime.strptime(st, "%Y-%m-%d %H:%M:%S"))
    return datetime.datetime.strptime(st, "%Y-%m-%d %H:%M:%S")
# 1.把datetime转成字符串
def datetime_toString(dt):
    #print("1.把datetime转成字符串: ", dt.strftime("%m月%d日"))
    return dt.strftime("%m").replace('0','')+'月'+dt.strftime("%d")+'日'
# 1.把datetime转成字符串
def datetime_toString2(dt):
    #print("1.把datetime转成字符串: ", dt.strftime("%m月%d日"))
    return dt.strftime("%Y/%m/%d")


#把输入的时间转成datetime
def ZhuanHuanDate(st):
    return st[0:4]+'-'+st[4:6]+'-'+st[6:8]+' 00:00:00'

def shuchuhangliezhi(sheet,i,j,k):
    #print('%s 行 %s 列-%s列的值为'+ sheet.col_values(colx=i,start_rowx=j,end_rowx=k)[0:2] %i %j %k)
    return ( sheet.col_values(colx=i, start_rowx=j, end_rowx=k))

def copy2(wb):
    w = XLWTWriter()
    process(XLRDReader(wb, 'unknown.xls'), w)
    return w.output[0][1], w.style_list

# rb = xlrd.open_workbook(name, formatting_info=True, on_demand=True)
# wb, s = copy2(rb)
# wbs = wb.get_sheet(0)
# rbs = rb.get_sheet(0)
# styles = s[rbs.cell_xf_index(0, 0)]
# rb.release_resources()  #关闭模板文件
# wbs.write(0, 0, 'aa', styles)
# wb.save("2.xls")

#开始和结束日期
datebegin = string_toDatetime(ZhuanHuanDate(date))
dateend = datebegin+datetime.timedelta(days=6)

#文件操作
workbook = xlrd.open_workbook(name, formatting_info=True, on_demand=True)
Newexcel, Copyexcel = copy2(workbook)
Newexcelsheet = Newexcel.get_sheet(0)
Copyexcels = workbook.get_sheet(0)
Excel_sheet = workbook.sheets()[0]; #原文件第一页
#styles = Copyexcels[Copyexcels.cell_xf_index(0, 0)]
#vbs = Copyexcels[Copyexcels.vert_split_pos]
#Excel_sheet = Newexcel.get_sheet(0); #原文件第一页
#print(shuchuhangliezhi(sheet=Excel_sheet,i=0,j=1,k=2)) #打印
# print(Newexcel[Newexcelsheet.cell_xf_index(0, 0)])
nrows = Excel_sheet.nrows #行数
ncols = Excel_sheet.ncols #列数
Newexcelsheet.write(0, 0, "正宗俱乐部"+datetime_toString(datebegin)+'-'+datetime_toString(dateend)+"活动费明细",Copyexcel[Copyexcels.cell_xf_index(0, 0)])
i = 0
str2 = ''
while(i < 7):
    str2 += datetime_toString(datebegin+datetime.timedelta(days=i))+': \n场地费折后：\n'
    Newexcelsheet.write(2,i+3,datetime_toString2(datebegin+datetime.timedelta(days=i)),Copyexcel[Copyexcels.cell_xf_index(2, i+3)])
    i+=1
Newexcelsheet.write(1, 0, str2,Copyexcel[Copyexcels.cell_xf_index(1, 0)]) #第二行第0列赋值


hang = 3
while(hang < nrows):
    lie = 3
    tihaun = shuchuhangliezhi(sheet=Excel_sheet, i=11, j=hang, k=hang + 1)#将上周余额替换至本周开始
    Newexcelsheet.write(hang, 2, tihaun[0], Copyexcel[Copyexcels.cell_xf_index(hang, lie)])
    gongshi = 'SUM(C4,-D4,-E4,-F4,-G4,-H4,-I4,-J4,K4)'.replace('4',str(hang+1))
    Newexcelsheet.write(hang, 11, xlwt.Formula(gongshi),Copyexcel[Copyexcels.cell_xf_index(hang, 11)])
    # Newexcelsheet.write_formula(rows=hang,col= 11, formula=gongshi)
    while lie > 2 and lie < 11:
        Newexcelsheet.write(hang, lie, '', Copyexcel[Copyexcels.cell_xf_index(hang, lie)])
        lie+=1
    hang += 1
    if hang == nrows:
        gongshi2 = 'SUM(L4: L64)'.replace('64', str(hang-1))
        gongshi3 = 'SUM(C4: C64)'.replace('64', str(hang-1))
        Newexcelsheet.write(hang-1, 11, xlwt.Formula(gongshi2),Copyexcel[Copyexcels.cell_xf_index(hang-1, 11)])
        Newexcelsheet.write(hang - 1, 2, xlwt.Formula(gongshi3), Copyexcel[Copyexcels.cell_xf_index(hang - 1, 2)])
workbook.release_resources()  #关闭模板文件
Newexcel.save("羽毛球经费周统计表"+datetime_toString(datebegin)+'-'+datetime_toString(dateend)+"活动费明细.xls")
print('程序已成功执行，请查看程序文件夹')
input("按任意键退出")




# 获取整行和整列的值（列表）
# rows = Excel_sheet.row_values(0)  # 获取第一行内容
# cols = Data_sheet.col_values(0)  # 获取第二列内容
# print(rows)
# print(cols)

