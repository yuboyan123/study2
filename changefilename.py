import os;
import time;
import openpyxl;


#  acquied localTime
localTime = time.strftime("%Y-%m-%d", time.localtime())
localTime =localTime.replace("-","",2)

print(localTime)


dir ="E:/��Ŀ����/1.��Ŀ����/������Ŀ/5-�ǻ۳���/3.�ܱ�/������ʦ/4.�����ܱ�-20210409/20220708"
path = input('�������ļ�·��(��β����/)��')
# ��ȡ��Ŀ¼�������ļ��������б���
fileList = os.listdir(path)

n = 0
j = 0
for i in fileList:
    # ���þ��ļ���������·��+�ļ�����
    oldname = fileList[n]
    fullPath = path + "/" + oldname;
    if(os.path.isfile(fullPath)):
        fullName = oldname.split('-')
        oldPlanFirstName = fullName[0]
        middleName = fullName[1] + "-"

        firstName = "����ר����ÿ�ܹ����ܽ�-"
        firstNamePlan = "���ϴ�ѧ ͬ�����Ŀÿ�ܼƻ���2022�����-"
        # �������ļ���
        if (oldPlanFirstName in firstNamePlan):
            newname = "/"+firstNamePlan + localTime + '.xlsx'
        else:
            newname = "/"+firstName + middleName + localTime + '.xlsx'

        os.rename(fullPath, dir + newname)  # ��osģ���е�rename�������ļ�����
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

            firstName = "����ר����ÿ�ܹ����ܽ�-"
            firstNamePlan = "ͬ�����Ŀÿ�ܼƻ���-"
            hive = "���ϴ�ѧ-"
            # �������ļ���
            if (middleName in hive):
                newname =  "/"+firstNamePlan + hive + "��������ϵͳ-" + localTime + '.xlsx'
            else:
                newname =  "/"+ firstName + middleName + localTime + '.xlsx'

            os.rename(dir+"/"+ oldname, dir + newname)  # ��osģ���е�rename�������ļ�����
            print(newname, '======>')
            wb = openpyxl.load_workbook(dir + newname)
            wb.save(dir+"/{}".format(newname[1:]))

            j+=1
    n += 1