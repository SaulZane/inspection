
from collections import defaultdict
from datetime import datetime
import locale
import pandas as pd
from colorama import Fore, Back, Style, init
init(autoreset=True)
locale.setlocale(locale.LC_CTYPE, 'chinese')
try:  
    print(Fore.GREEN+u"-----欢迎使用查验检验员违规查验同一车型嫌疑程序导出软件-----©张硕 保留所有权利|2024.2.28|V1.0-----")
    print(Fore.CYAN+u"使用说明：输入表名称必须是“表2.xlsx”，必须有“基础表”数据页。并且必须在程序的同名文件夹下")
    print(Fore.CYAN+u"必须含有“查验日 查验员 社会机构名称 车辆类型” 四个字段，表头名字不能变（其中查验日是字符格式，可以使用公式L2==TRUNC(K2,0)预处理下）")
    print(Fore.CYAN+u"如果运行成功，在同名文件夹下会看到result.xlsx结果表")
    print(Fore.GREEN+u"按任意键开始运行程序...")
    input()
    print(Fore.RED+u"--以下为调试信息，不要乱碰电脑--")
    #使用pandas库读取excel文件
    data = pd.read_excel(u'表2.xlsx',sheet_name=u'基础表',header=0)
    #将数据转换为字典
    df=pd.DataFrame.from_records(data)
    #将数据转换为字典
    records = df.to_dict(orient='records')

    threshold :int =5

    def find_violations(records,threshold=3):
        counts=defaultdict(int)
        violations=[]
        for record in records:
            d=record[u'查验日']
            i=record[u'查验员']
            u=record[u'社会机构名称']
            b=record[u'车辆类型']
            key=(i,u,d,b)      
            counts[key]+=1
        #将大于违规的记录添加到violations列表中
        for key in counts:
            if counts[key]>=threshold:
                #把次数也加入到violations列表中，然后按照次数排序
                violations.append((key,counts[key]))
        violations.sort(key=lambda x:x[1],reverse=True)
                
                
        return violations

    print(find_violations(records,threshold))

    #输出结果，新建一个excel文件，将结果写入到excel文件中，名称为result.xlsx
    result = find_violations(records,threshold)



    data = [(item[0][0], item[0][1], item[0][2].strftime(u'%Y年%m月%d日'), item[0][3],  item[1]) for item in result]
    df = pd.DataFrame(data,columns=[u'查验员',u'社会机构名称',u'查验日',u'车辆类型',u'数量'])
    df.to_excel('result.xlsx', index=False)
    print('-------')
    print(Fore.RED+u'恭喜！！result.xlsx文件已生成，可以操作电脑了!按任意键退出程序...')
    input()
except Exception as e:
    print(Fore.RED+u"程序出现异常，请检查表格是否符合要求，或者联系开发者"+str(e))
    input()


