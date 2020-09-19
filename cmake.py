from docx import *
import pandas as pd
import os
a = []
i = 0
c =[]
for root, dirs, files in os.walk("code", topdown=False):#获取本文件夹下文件名称列表  List
    for name in files:
        x2 = os.path.join(root, name)
        # print(x2)
        c.append(x2[5:])

# '''
for m in c:
    a1 = []
    a2 = []
    document = Document("code/"+m)#Cccccccccccccccccccccccccccccccccccccccccccccccchange
    for paragraph in document.paragraphs:  #遍历文档
        x = paragraph.text  #获取每行word的内容为str
        b = len(x)  #获取行字符个数
        e1 = 0
        e2 = 1
        s = 80
        c = "                                                                                           "
        # if(x[:b] != c[:b] and x[:6] != "import"): #去除空行数和非import判断， 
        if(x[:b] != c[:b] and x[:6] == "import" and len(a2)<100): #去除空行数和import判断，     
            # if (b >=80):   
            #     while (b >= 80):
            #         s1 = s * e1
            #         s2 = s * e2
            #         t = x[s1:s2]#按照间隔100字符进行截取
            #         a.append(t)
            #         b = len(t) #获取截取后的 字符串字数
            #         e1 = e1 + 1
            #         e2 = e2 + 1
            # else:
            #     a2.append(x)
            a2.append(x)  
            continue
        elif(x[:b] != c[:b] and x[:6] != "import" and len(a1)<2900):
            if (b >=80):   
                while (b >= 80):
                    s1 = s * e1
                    s2 = s * e2
                    t = x[s1:s2]#按照间隔100字符进行截取
                    a1.append(t)
                    b = len(t) #获取截取后的 字符串字数
                    e1 = e1 + 1
                    e2 = e2 + 1
            else:
                a1.append(x)
        if(len(a1)>=100 and len(a2) >= 2900):
            break
        sum = a2 + a1
    dataframe = pd.DataFrame(sum)
    print(dataframe)
    dataframe.to_excel('import/origin-%s.xls'%m[:9])#Cccccccccccccccccccccccccccccccccccccccccccccccchange
    print("change-%s已完成……"%m)#Cccccccccccccccccccccccccccccccccccccccccccccccchange
            
    # break
    # 只生成一个 测试
# '''
print("All Done!")


    # for name in dirs:
    #     print(os.path.join(root, name))


# dataframe = pd.DataFrame(a)
# print(dataframe)
# dataframe.to_excel('la.xls')
