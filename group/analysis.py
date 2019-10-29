import pandas as pd
import os

#os.chdir("/Users/cilab/PartTime_PythonAnalysis/group")
os.chdir(r"C:\Users\Naichen\Documents\GitHub\stu00608.github.io\PartTime_PythonAnalysis\group")
os.getcwd()

import ct_tool as ct

#記得路徑不同電腦要改

statData = pd.read_csv("002.csv")
#statData = statData.dropna(axis=1,how='all')

#os.chdir("/Users/cilab/PartTime_PythonAnalysis/group/outputs/analysis")
os.chdir(r"C:\Users\Naichen\Documents\GitHub\stu00608.github.io\PartTime_PythonAnalysis\group\outputs\analysis")
os.getcwd()

#-------------------------------------------------------------------------#




df = ct.extraction(statData,"3.工程領域")

df1 = ct.group_table(df,"3.工程領域",["土木營建"])

df2 = df1[['土木營建領域專長 男性 專任 人數','土木營建領域專長 女性 專任 人數','服務年資1~5年','服務年資6~10年','服務年資11~15年',
 '服務年資16~20年','服務年資21~25年','服務年資25年以上','管理職年資1~5年','管理職年資6~10年',
 '管理職年資11~15年','管理職年資16~20年','管理職年資21~25年','管理職年資25年以上',]].reset_index(drop=True)

df2 = df2.fillna(0)

delList=[]
for item in df2:
    tempList = ct.digitCheck(df2,item)
    delList = list(set(delList+tempList))

df2 = df2.join(df1['1.單位名稱'].reset_index(drop=True))


df3 = df2.drop(delList).reset_index(drop=True).fillna(0)



#分成 : 經濟部水利署、公路總局、其他

for i in range(len(df3['1.單位名稱'])):
    if( df3['1.單位名稱'][i].find('經濟部水利署') != -1 ):
        df3.loc[i,'1.單位名稱']='經濟部水利署'
    elif( df3['1.單位名稱'][i].find('公路總局') != -1 ):
        df3.loc[i,'1.單位名稱']='公路總局'
    else:
        result = df3.loc[i,]
        result = pd.DataFrame(result)
        name = df3.loc[i,][-1]
        result = result.rename({0:name},axis='columns')
        result = result.drop('1.單位名稱',axis=0)
        result.to_csv(name+'.csv',encoding='utf_8_sig')

result1 = ct.group_table(df3,'1.單位名稱',['經濟部水利署']).reset_index(drop=True).drop('1.單位名稱',axis=1)

result1 = result1.astype(int).sum()

result1 = pd.DataFrame(result1,columns=['經濟部水利署'])

result1.to_csv('經濟部水利署.csv',encoding='utf_8_sig')

result2 = ct.group_table(df3,'1.單位名稱',['公路總局']).reset_index(drop=True).drop('1.單位名稱',axis=1)

result2 = result2.astype(int).sum()

result2 = pd.DataFrame(result2,columns=['公路總局'])

result2.to_csv('公路總局.csv',encoding='utf_8_sig')


#-------------------------------------------------------------------------#


os.chdir("/Users/cilab/PartTime_PythonAnalysis/group/outputs/leave")
os.getcwd()

df = ct.extraction(statData,"3.工程領域")

df1 = ct.group_table(df,"3.工程領域",["土木營建"])

df2 = df1[['1.單位名稱','2.單位總員工人數','3.工程領域','男性',
 '106年 請假人次','107年  請假人次','108年 請假人次',
 '女性','106年  請假人次','108年  請假人次']].reset_index(drop=True)



for i in range(len(df2['1.單位名稱'])):
    if( df2['1.單位名稱'][i].find('經濟部水利署') != -1 ):
        df2.loc[i,'1.單位名稱']='經濟部水利署'
    elif( df2['1.單位名稱'][i].find('公路總局') != -1 ):
        df2.loc[i,'1.單位名稱']='公路總局'
    else:
        result = df2.loc[i,]
        result.name = result[0]
        name = result.name
        result = pd.DataFrame(result)
        result = result.drop(['1.單位名稱',"3.工程領域",'男性','女性'],axis=0).fillna(0).reset_index(drop=True)

        for j in range(len(result[name])):
            if(not str(result[name][j]).isdigit()):
                result.loc[j,name]=0

        result = result.astype(int).reset_index(drop=True).rename(index={0:'總員工人數',1:'106年男性請假人次',2:'107年男性請假人次',3:'108年男性請假人次',4:'106年女性請假人次',5:'108年女性請假人次'})
        result.to_csv(name+'_請假分析.csv',encoding='utf_8_sig')




doList = ['經濟部水利署','公路總局']
for item in doList:

    result1 = ct.group_table(df2,'1.單位名稱',[item]).reset_index(drop=True).drop(['1.單位名稱',"3.工程領域",'男性','女性'],axis=1).fillna(0)

    delList=[]
    for c in result1:
        tempList = ct.digitCheck(result1,c)
        delList = list(set(delList+tempList))

    result1 = result1.drop(delList).astype(int)

    result2 = result1.describe().rename(index={"mean":"平均數","std":"標準差","min":"最小值","max":"最大值","50%":"中位數"}).astype(int)

    result2.loc['sum']=list(result1.sum().astype(int))

    result2.columns.name=item

    result2 = result2.rename(columns={'106年 請假人次':'106年男性請假人次','107年  請假人次':'107年男性請假人次','108年 請假人次':'108年男性請假人次','106年  請假人次':'106年女性請假人次','108年  請假人次':'108年女性請假人次'})

    result2.to_csv(item+'_請假分析.csv',encoding='utf_8_sig')

#-------------------------------------------------------------------------#

os.chdir("/Users/cilab/PartTime_PythonAnalysis/group/outputs/babycare")
os.getcwd()

#df = ct.extraction(statData,"3.工程領域").reset_index(drop=True)

df = ct.extraction(statData,'可複選：',"貴單位已實施之福利措施").reset_index(drop=True)

ct1 = pd.crosstab(df["貴單位已實施之福利措施"],df["3.工程領域"]).rename(index={"Third Choice托嬰服務（指設有收托二歲以下兒童之服務機構）":"托嬰服務（指設有收托二歲以下兒童之服務機構）"})

ct1.to_csv('貴單位已實施之福利措施.csv',encoding='utf_8_sig')

df1 = ct.group_table(df,"貴單位已實施之福利措施",["哺集乳室"])

df1 = df1[['1.單位名稱','3.工程領域','請問哺集乳室共設有幾處？']].fillna(0).reset_index(drop=True)

delList = []
for i in range(len(df1['請問哺集乳室共設有幾處？'])):
    if(not str(df1['請問哺集乳室共設有幾處？'][i]).isdigit()):
        delList.append(i)

df2 = df1.drop(delList).reset_index(drop=True)


for i in range(len(df2['1.單位名稱'])):
    if( df2['1.單位名稱'][i].find('經濟部水利署') != -1 ):
        df2.loc[i,'1.單位名稱']='經濟部水利署'
    elif( df2['1.單位名稱'][i].find('公路總局') != -1 ):
        df2.loc[i,'1.單位名稱']='公路總局'

doList = ['經濟部水利署','公路總局']
for item in doList:

    result1 = ct.group_table(df2,'1.單位名稱',[item]).reset_index(drop=True).drop(['1.單位名稱','3.工程領域'],axis=1).fillna(0).astype(int)

    result1 = result1.describe().rename(index={"mean":"平均數","std":"標準差","min":"最小值","max":"最大值","50%":"中位數"}).round(2)
    
    result1.name = item

    result1.to_csv(item+'_哺乳室分析.csv',encoding='utf_8_sig')

#-------------------------------------------------------------------------#

