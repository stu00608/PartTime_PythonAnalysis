import pandas as pd
import os
import ct_tool as ct

#記得路徑不同電腦要改
os.chdir(r"C:\Users\Naichen\Documents\GitHub\stu00608.github.io\PartTime_PythonAnalysis\group")
os.getcwd()
statData = pd.read_csv("002.csv")
#statData = statData.dropna(axis=1,how='all')
os.chdir(r"C:\Users\Naichen\Documents\GitHub\stu00608.github.io\PartTime_PythonAnalysis\group\outputs")
os.getcwd()

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



#分成 : 經濟部水利署、公路總局、...

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