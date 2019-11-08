import pandas as pd
import os

#os.chdir("/Users/cilab/PartTime_PythonAnalysis/personal")
os.chdir(r"C:\Users\Naichen\Documents\GitHub\stu00608.github.io\PartTime_PythonAnalysis\personal")

os.getcwd()

import ct_tool as ct

statData = pd.read_csv("001.csv")
statData = statData.dropna(axis=1,how='all')
#os.chdir("/Users/cilab/PartTime_PythonAnalysis/personal/outputs")
os.chdir(r"C:\Users\Naichen\Documents\GitHub\stu00608.github.io\PartTime_PythonAnalysis\personal\outputs")
os.getcwd()

#把性別跟學歷的"其他"刪除
for i in range(len(statData['1. 性別：'])):
    if(statData['1. 性別：'][i]=='其他'):
        break
statData = statData.drop(i).reset_index(drop=True)

for i in range(len(statData['3. 最高學歷'])):
    if(statData['3. 最高學歷'][i]=='其他（請說明）'):
        break
statData = statData.drop(i).reset_index(drop=True)

#初始化
resultCSV = []

#---------------------------------------------------------------------------------------#

#性別
ct1 = pd.crosstab(statData["2. 年齡"],statData["1. 性別："],margins=True)
ct1 = ct1.rename(columns={'All':'小計'},index={'All':'合計'})
ct1.columns.name = "性別"
ct1.index.name = "年齡"

ct2 = ct1/ct1['小計'][-1]
ct2 = ((ct2*100).round(1).astype(str)+'%').replace('0.0%','-')

resultCSV.append(ct1)
resultCSV.append(ct2)

# ct1.to_csv('1.年齡.csv',encoding='utf_8_sig')
# ct2.to_csv('1.年齡(比例).csv',encoding='utf_8_sig')

#最高學歷

df1 = statData[['1. 性別：','2. 年齡','3. 最高學歷']].sort_values(by=['1. 性別：']).reset_index(drop=True).rename(columns={'2. 年齡':'年齡',"1. 性別：":'性別','3. 最高學歷':'最高學歷'})

ct1 = pd.crosstab(statData['3. 最高學歷'],[statData['2. 年齡'],statData["1. 性別："]],margins=True)
ct1 = ct1.rename(columns={'All':'小計'},index={'All':'合計'})
ct1.columns.names=["年齡","性別"]
ct1.index.name = "最高學歷"
ct1 = ct1.reindex(['博士','碩士','大學/大專','專科','高職','合計'])
ct1 = ct1/ct1['小計'][-1]
ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-')

resultCSV.append(ct1)
# ct1.to_csv('2.學歷.csv',encoding='utf_8_sig')

#畢業校系
ct1 = pd.crosstab(statData['7. 是否正在從事工程與科技領域職務（含管理與學術研究）？'],[statData['4.\t最高學歷畢業校系'],statData['1. 性別：']],margins=True)
ct1 = ct1/ct1['All'][-1]
ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-')
ct1 = ct1.drop('All')
ct1 = ct1.reindex(['是','否'])
ct1 = ct1.rename(columns={'All':'小計'},index={'是':'工程與科技職務','否':'非工程與科技職務'})
ct1.columns.names=["最高學歷","性別"]
ct1.index.name = ""


resultCSV.append(ct1)
# ct1.to_csv('3.最高學歷.csv',encoding='utf_8_sig')

#總年資

df1 = statData[['1. 性別：','3. 最高學歷','4.\t最高學歷畢業校系','7. 是否正在從事工程與科技領域職務（含管理與學術研究）？','8. 您離開工程職務之最主要原因為何？請勾選最主要原因，至多二項。','(a) 工程與科技領域職務(年)', '(b) 非工程與科技領域職務','10.\t請問您目前服務的職務較接近下列何者？','請問目前受您管理的人員約＿＿人？','9.\t請問您目前的工作狀況為何？']]

delList =[]
for i in range(len(df1)):
    if( (not str(df1['(a) 工程與科技領域職務(年)'][i]).isdigit()) or (not str(df1['(b) 非工程與科技領域職務'][i]).isdigit())) :
        delList.append(i)
df2 = df1.drop(delList).reset_index(drop=True)
t1 = df2['(a) 工程與科技領域職務(年)'].fillna(0).astype(int)
t2 = df2['(b) 非工程與科技領域職務'].fillna(0).astype(int)
df2['總年資']=t1+t2

##     用範圍對df2分類
for i in range(len(df2)):

    v1 = int(df2['總年資'][i])
    if( v1<=5 and v1>=0) :
        df2.loc[i,'總年資']='0~5年'
    elif( v1>5 and v1<11 ):
        df2.loc[i,'總年資']='6~10年'
    elif( v1>=11 and v1<16 ):
        df2.loc[i,'總年資']='11~15年'
    elif( v1>=16 and v1<21 ):
        df2.loc[i,'總年資']='16~20年'
    elif( v1>=21 and v1<26 ):
        df2.loc[i,'總年資']='21~25年'
    elif( v1>=26 and v1<31 ):
        df2.loc[i,'總年資']='26~30年'
    elif( v1>=31 and v1<100 ):
        df2.loc[i,'總年資']='30年以上'
    else:
        df2.loc[i,'總年資']=None
        
ct1 = pd.crosstab(df2['4.\t最高學歷畢業校系'],[df2['總年資'],df2['1. 性別：']],margins=True)
ct1 = ct1/ct1['All'][-1]
ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-')
ct1.columns.names=["總年資","性別"]
ct1.index.name="最高學歷"
ct1 = ct1.rename(columns={'All':'小計'},index={'All':'合計'})

resultCSV.append(ct1)
# ct1.to_csv('4.總年資.csv',encoding='utf_8_sig')

#離開原因

df1 = statData[['1. 性別：','3. 最高學歷','4.\t最高學歷畢業校系','7. 是否正在從事工程與科技領域職務（含管理與學術研究）？','8. 您離開工程職務之最主要原因為何？請勾選最主要原因，至多二項。','10.\t請問您目前服務的職務較接近下列何者？','(a) 工程與科技領域職務(年)']]

delList =[]
for i in range(len(df1)):
    if( not str(df1['(a) 工程與科技領域職務(年)'][i]).isdigit() ) :
        delList.append(i)
df1 = df1.drop(delList).reset_index(drop=True)

df_del = ct.group_table(df1,'(a) 工程與科技領域職務(年)',['0'])

df_del = ct.group_table(df_del,'7. 是否正在從事工程與科技領域職務（含管理與學術研究）？',['否'])

df_del = ct.group_table(df_del,"4.\t最高學歷畢業校系",['國外學校，非工程與科技相關領域','國內學校，非工程與科技相關領域'])

delList = list(df_del.index)
delList.sort()
df2 = df1.drop(delList) 
df2 = df2.reset_index(drop=True)
df2 = df2.dropna(subset=['8. 您離開工程職務之最主要原因為何？請勾選最主要原因，至多二項。']).reset_index(drop=True)
df2 = ct.extraction(df2,'8. 您離開工程職務之最主要原因為何？請勾選最主要原因，至多二項。','您離開工程職務之最主要原因為何？').reset_index(drop=True)

ct3 = pd.crosstab(df2['您離開工程職務之最主要原因為何？'],df2['1. 性別：'],margins=True)
temp = int(ct3['All'][-1])
ct3 = ct3/temp
ct3 = ((ct3*100).round(1).astype(str)+'%').replace('0.0%','-')
ct3 = ct3.drop("其他（請說明）",axis=0)
ct3 = ct3.rename(index={'公司組織調整（例如裁撤、縮編或退出本地市場）':'公司組織調整','All':'人數'})
ct3.columns.name='性別'
ct3.index.name="離開工程職務原因(%d人)"%(temp)

resultCSV.append(ct3)
# ct3.to_csv('5.離開工程職務最主要原因分析.csv',encoding='utf_8_sig')

#目前工作狀況

df1 = ct.group_table(statData,'7. 是否正在從事工程與科技領域職務（含管理與學術研究）？',['否'])

ct1 = pd.crosstab(df1['9.\t請問您目前的工作狀況為何？'],df1['1. 性別：'],margins=True)
temp = int(ct1['All'][-1])
ct1 = ct1/temp
ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-')
ct1 = ct1.rename(index={'All':'合計'},columns={'All':'小計'})
ct1.columns.name='性別'
ct1.index.name="目前的工作狀況(%d人)"%(temp)

resultCSV.append(ct1)
# ct1.to_csv('7.目前工作狀況.csv',encoding='utf_8_sig')

#五年後

df = statData[['1. 性別：','3. 最高學歷','11.\t您對於現任職務之五年後職涯發展的預期為何？請勾選最有可能的一項。']]

ct1 = pd.crosstab(df['11.\t您對於現任職務之五年後職涯發展的預期為何？請勾選最有可能的一項。'],[df['3. 最高學歷'],df['1. 性別：']],margins=True)
ct1 = ct1/ct1['All'][-1]
ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-')
ct1 = ct1[['博士','碩士','大學/大專','專科','高職','All']]
ct1.columns.names=['最高學歷','性別']
ct1.index.name="現任職務之五年後職涯發展的預期"
ct1 = ct1.rename(index={'轉換領域(轉至非工程與科技領域或轉入工程與科技領域)   (請說明原因)':'轉換領域(轉至非工程與科技領域或轉入工程與科技領域)','離職 (請說明原因)':'離職','All':'合計'},columns={'All':'小計'})

resultCSV.append(ct1)
# ct1.to_csv('8.對於現任職務之五年後職涯發展的預期.csv',encoding='utf_8_sig')

#配合因素

df = statData[['1. 性別：','3. 最高學歷','12.\t請問您所預期的職涯發展，需要哪些配合因素來達成？請勾選最主要因素，至多二項。']]

for i in range(len(df)):
    if(not (df['12.\t請問您所預期的職涯發展，需要哪些配合因素來達成？請勾選最主要因素，至多二項。'][i].find('還不清楚（若勾選本項請勿再勾選其他項目）') == -1) ):
        df.loc[i,'12.\t請問您所預期的職涯發展，需要哪些配合因素來達成？請勾選最主要因素，至多二項。'] = '還不清楚'

df1 = ct.extraction(df,'12.\t請問您所預期的職涯發展，需要哪些配合因素來達成？請勾選最主要因素，至多二項。','請問您所預期的職涯發展，需要哪些配合因素來達成？')

ct1 = pd.crosstab(df1['請問您所預期的職涯發展，需要哪些配合因素來達成？'],[df1['3. 最高學歷'],df1['1. 性別：']],margins=True)
ct1 = ct1/ct1['All'][-1]
ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-')
ct1 = ct1[['博士','碩士','大學/大專','專科','高職','All']]
ct1 = ct1.rename(index={'All':'合計'},columns={'All':'小計'})
ct1.columns.names=['最高學歷','性別']
ct1.index.name="預期的職涯發展配合因素"

ct2 = pd.crosstab(df1['請問您所預期的職涯發展，需要哪些配合因素來達成？'],df1['1. 性別：'],margins=True)
ct2 = ct2/ct2['All'][-1]
ct2 = ((ct2*100).round(1).astype(str)+'%').replace('0.0%','-')
ct2 = ct2.rename(index={'All':'合計'},columns={'All':'小計'})
ct2.columns.name='性別'
ct2.index.name="預期的職涯發展配合因素"

resultCSV.append(ct1)
resultCSV.append(ct2)
# ct1.to_csv('9.預期的職涯發展需要哪些配合因素來達成.csv',encoding='utf_8_sig')
# ct2.to_csv('9.預期的職涯發展需要哪些配合因素來達成(性別).csv',encoding='utf_8_sig')

#有助福利

df = statData[['1. 性別：','3. 最高學歷','13.\t哪些福利措施最有助於您留在工程與科技領域就業？（不論現在是否有需求皆可選擇，可複選，最多三項）']]

df1 = ct.extraction(df,'13.\t哪些福利措施最有助於您留在工程與科技領域就業？（不論現在是否有需求皆可選擇，可複選，最多三項）','哪些福利措施最有助於您留在工程與科技領域就業？')

ct1 = pd.crosstab(df1['哪些福利措施最有助於您留在工程與科技領域就業？'],[df1['3. 最高學歷'],df1['1. 性別：']],margins=True).drop(' 請勿再勾選其他選項 )')
ct1 = ct1/ct1['All'][-1]
ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-')
ct1 = ct1[['博士','碩士','大學/大專','專科','高職','All']]
ct1 = ct1.rename(index={'All':'合計','其他（請說明）':'其他','無，個人不需要這些福利 ( 若選此項':'無，個人不需要這些福利'},columns={'All':'小計'})
ct1.columns.names=['最高學歷','性別']
ct1.index.name="福利措施"

ct2 = pd.crosstab(df1['哪些福利措施最有助於您留在工程與科技領域就業？'],df1['1. 性別：'],margins=True).drop(' 請勿再勾選其他選項 )')
ct2 = ct2/ct2['All'][-1]
ct2 = ((ct2*100).round(1).astype(str)+'%').replace('0.0%','-')
ct2 = ct2.rename(index={'All':'合計','其他（請說明）':'其他','無，個人不需要這些福利 ( 若選此項':'無，個人不需要這些福利'},columns={'All':'小計'})
ct2.columns.name='性別'
ct2.index.name="福利措施"

resultCSV.append(ct1)
resultCSV.append(ct2)

# ct1.to_csv('10.哪些福利措施最有助於您留在工程與科技領域就業.csv',encoding='utf_8_sig')
# ct2.to_csv('10.哪些福利措施最有助於您留在工程與科技領域就業(性別).csv',encoding='utf_8_sig')

#單位福利

df1 = statData[['1. 性別：','10.\t請問您目前服務的職務較接近下列何者？','13.\t哪些福利措施最有助於您留在工程與科技領域就業？（不論現在是否有需求皆可選擇，可複選，最多三項）',
       '14.\t您服務的單位或待業前服務的單位提供哪些職場相關措施？（可複選）',
       '15.\t您認為工程與科技領域最需要改善的性別議題有哪些？請勾選其中最重要的議題，至多三項。',]]

df3 =  ct.extraction(df1,'14.\t您服務的單位或待業前服務的單位提供哪些職場相關措施？（可複選）','您服務的單位或待業前服務的單位提供哪些職場相關措施？')

ct1 = pd.crosstab(df3['您服務的單位或待業前服務的單位提供哪些職場相關措施？'],df3['1. 性別：'],margins=True)
ct1 = ct1/ct1['All'][-1]
ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-').drop([' 請勿再勾選其他選項 )','其他，請說明',],axis=0)
ct1 = ct1.rename(index={'All':'合計','無，沒有提供 ( 若選此項':'無，沒有提供'},columns={'All':'小計'})
ct1.columns.name='性別'

resultCSV.append(ct1)
# ct1.to_csv("11.您服務的單位或待業前服務的單位提供哪些職場相關措施.csv",encoding='utf_8_sig')

#需要改善

df4 = ct.extraction(statData,'15.\t您認為工程與科技領域最需要改善的性別議題有哪些？請勾選其中最重要的議題，至多三項。','您認為工程與科技領域最需要改善的性別議題有哪些？')

ct1 = pd.crosstab(df4['您認為工程與科技領域最需要改善的性別議題有哪些？'],df4['1. 性別：'],margins=True).drop(' 請勿再勾選其他選項 )',axis=0)
ct1 = ct1/ct1['All'][-1]
ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-')
ct1 = ct1.rename(index={'All':'合計','無，沒有提供 ( 若選此項':'無，沒有提供','不了解，沒有想法  ( 若選此項':'不了解，沒有想法','無，沒必要改善  ( 若選此項':'無，沒必要改善'},columns={'All':'小計'})
ct1.columns.name='性別'

resultCSV.append(ct1)
# ct1.to_csv("12.您認為工程與科技領域最需要改善的性別議題有哪些.csv",encoding='utf_8_sig')

#---------------------------------------------------------------------------------------#

name = ['年齡','年齡(比例)','學歷','最高學歷','總年資','離開工程職務最主要原因分析','目前工作狀況','對於現任職務之五年後職涯發展的預期','預期的職涯發展需要哪些配合因素來達成',
        '預期的職涯發展需要哪些配合因素來達成','哪些福利措施最有助於您留在工程與科技領域就業','哪些福利措施最有助於您留在工程與科技領域就業(性別)','您服務的單位或待業前服務的單位提供哪些職場相關措施',
        '您認為工程與科技領域最需要改善的性別議題有哪些']

with pd.ExcelWriter('個人版分析.xlsx') as writer:
    for i in range(len(name)):
        resultCSV[i].to_excel(writer,sheet_name=name[i],encoding='utf_8_sig')
