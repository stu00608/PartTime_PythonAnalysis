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

data_col = list(df1.columns)

df = ct.extraction(statData,"3.工程領域")

df1 = ct.group_table(df,"3.工程領域",["土木營建"])

df2 = df1[['土木營建領域專長 男性 專任 人數','土木營建領域專長 女性 專任 人數',
'土木營建領域專長 男性 人數','服務年資1~5年','服務年資6~10年','服務年資11~15年',
 '服務年資16~20年','服務年資21~25年','服務年資25年以上','土木營建領域專長 男性 管理職 人數',
 '管理職年資1~5年','管理職年資6~10年','管理職年資11~15年','管理職年資16~20年',
 '管理職年資21~25年','管理職年資25年以上',]].reset_index(drop=True)


delList = []

for j in range(len(df2['土木營建領域專長 男性 專任 人數'])):
    if(not str(df2['土木營建領域專長 男性 專任 人數'][j]).isdigit()):
        if(not delList.count(j)):
            delList.append(j)
    else:
        if(int(df2['土木營建領域專長 男性 專任 人數'][j])>100000):
            if(not delList.count(j)):
                delList.append(j)

df2.drop(delList).reset_index(drop=True).fillna(0).astype(int)


df1 = pd.concat([df2,df1['1.單位名稱']])


#分成 : 經濟部水利署、公路總局、...

for i in range(len(df1['1.單位名稱'])):
    if( df1['1.單位名稱'][i].find('經濟部水利署') != -1 ):
        df1.loc[i,'1.單位名稱']='經濟部水利署'
    elif( df1['1.單位名稱'][i].find('公路總局') != -1 ):
        df1.loc[i,'1.單位名稱']='公路總局'



result1 = ct.group_table(df1,'1.單位名稱',['經濟部水利署']).reset_index(drop=True)

