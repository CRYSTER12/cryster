import os
import xlrd
n = 1

path = 'D:/Desktop/附件四'     #改名目标路径
os.chdir(path)                                 #切换到当前目录
''''os.listdir(path)操作效果为返回指定路径（path）文件夹中所有文件名'''
filename_list = os.listdir(path) #扫描目标路径文件，将文件名录入列表
# print(filename_list)
# print(os.getcwd())

keyword = '2020年春季学期学生返校申请表'

table = xlrd.open_workbook('D:/Desktop/信息.xlsx').sheet_by_index(0)  #数据表

while n < 25:        #更改数目+1
    f = table.cell_value(n, 1)   #更改名在列表中位置
    g = table.cell_value(n, 0)   #更改名在列表中位置
    for name in filename_list:
       # 遍历目标文件 print(name)

        if not os.path.isdir(name): #判断是否为目录
            if name.find(f) != -1:  #判断文件名是否包含名字，不包含=-1，包含执行下面内容

            #if keyword in name:
                #new_name =name.replace(keyword, '')
                new_name = '%s-%s-健康登记表.docx' % (g, f)
                os.renames(name, new_name)
        else:
            print(path + '\\' + name)
            os.rename(path + '\\' + name, keyword)    #更改目录名，有无路径均可
            os.chdir('...')
    n += 1
else:
    print('匹配完成')
