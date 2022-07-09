import os;
import time;
import openpyxl;


#  acquied localTime
localTime = time.strftime("%Y-%m-%d", time.localtime())
localTime =localTime.replace("-","",2)

print(localTime)


dir ="E:/项目管理/1.项目资料/横向项目/5-智慧出行/3.周报/王琴老师/4.车列周报-20210409/20220708"
path = input('请输入文件路径(结尾加上/)：')
# 获取该目录下所有文件，存入列表中
fileList = os.listdir(path)

n = 0
j = 0
for i in fileList:
    # 设置旧文件名（就是路径+文件名）
    oldname = fileList[n]
    fullPath = path + "/" + oldname;
    if(os.path.isfile(fullPath)):
        fullName = oldname.split('-')
        oldPlanFirstName = fullName[0]
        middleName = fullName[1] + "-"

        firstName = "湖大专题组每周工作总结-"
        firstNamePlan = "湖南大学 同清湖项目每周计划表【2022】徐彪-"
        # 设置新文件名
        if (oldPlanFirstName in firstNamePlan):
            newname = "/"+firstNamePlan + localTime + '.xlsx'
        else:
            newname = "/"+firstName + middleName + localTime + '.xlsx'

        os.rename(fullPath, dir + newname)  # 用os模块中的rename方法对文件改名
        print(newname, '======>')
        wb = openpyxl.load_workbook(dir + newname)
        wb.save(dir+"/{}".format(newname[1:]))


    else:
        dir = dir +"/"+ oldname
        fileNewList = os.listdir(dir);
        for i in fileNewList:
            oldname = fileNewList[j]
            fullPath = path + "/" + oldname;
            fullName = oldname.split('-')

            middleName = fullName[1] + "-"
            lastName = fullName[2]

            firstName = "湖大专题组每周工作总结-"
            firstNamePlan = "同清湖项目每周计划表-"
            hive = "湖南大学-"
            # 设置新文件名
            if (middleName in hive):
                newname =  "/"+firstNamePlan + hive + "车辆调度系统-" + localTime + '.xlsx'
            else:
                newname =  "/"+ firstName + middleName + localTime + '.xlsx'

            os.rename(dir+"/"+ oldname, dir + newname)  # 用os模块中的rename方法对文件改名
            print(newname, '======>')
            wb = openpyxl.load_workbook(dir + newname)
            wb.save(dir+"/{}".format(newname[1:]))

            j+=1
    n += 1