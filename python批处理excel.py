import xlrd 
import matplotlib.pyplot as plt  
import sys
"""本程序目的是统计EXCEL中每个工作表中的B7单元格的数据，并进行数据清理，最后绘制成柱状图"""
"""打开文件"""
xlsx=xlrd.open_workbook("data.xlsx");
"""获取所有工作表"""
get_worksheetnames=xlsx.sheet_names()   #获得所有工作表的名字
    #获得所有工作表中B7单元格的数据
get_cellvalues=[]    #创造一个空数组保存数据
new_cellvalues=[]    #清洗后数据

for worksheetname in get_worksheetnames[2:30]:   #遍历
    open_sheet=xlsx.sheet_by_name(worksheetname)  #打开工作表
    get_cellvalue=open_sheet.cell(6,1).value       #需要注意，选择第7行第2列的数据用（6，1表示）
    get_cellvalues.append(get_cellvalue)           #向空数组中添加元素     

class Cleanvalue():          
    """清洗数据，去掉异常数据"""
    def __init__(self,value):
        self.value=value
    def cellvalues_avaerage(self):    #取产量平均值
        sum=0
        for key in self: 
            sum=sum+key
        avaerage=sum//len(self)
        return int(avaerage)
    def del_Abnormaldata (self,avaerage):      #去掉异常数据
        for get_cellvalue in self:          
            i=1
            if get_cellvalue<int(avaerage):
                new_cellvalues[i-1]=int(avaerage)
            else:
                new_cellvalues.append(get_cellvalue)
            i=i+1
        return new_cellvalues
my_Cleanvalue=Cleanvalue.cellvalues_avaerage(get_cellvalues)
new_Cleanvalue= Cleanvalue.del_Abnormaldata(get_cellvalues,my_Cleanvalue)

plt.bar(range(len(new_Cleanvalue)),new_Cleanvalue)  #输出图像
plt.show()  
