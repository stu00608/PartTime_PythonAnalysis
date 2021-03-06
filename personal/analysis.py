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

#statData.to_excel('testInput.xlsx',encoding='utf_8_sig')

#初始化
resultCSV = []

#---------------------------------------------------------------------------------------#

#性別O
ct1 = pd.crosstab(statData["2. 年齡"],statData["1. 性別："],margins=True)
ct1 = ct1.rename(columns={'All':'總人數'},index={'All':'各年齡小計'})
ct1.columns.name = ""
ct1.index.name = "年齡"

ct1['女性占比'] = (((ct1['女性']/ct1['總人數'])*100).round(1).astype(str)+'%').replace('0.0%','-')
ct1 = ct1.drop(['男性','女性'],axis=1)
ct1 = ct1.reindex(columns=['女性占比','總人數'])

resultCSV.append(ct1)

# ct1.to_csv('1.年齡.csv',encoding='utf_8_sig')
# ct2.to_csv('1.年齡(比例).csv',encoding='utf_8_sig')


#最高學歷

df1 = statData[['1. 性別：','2. 年齡','3. 最高學歷']].sort_values(by=['1. 性別：']).reset_index(drop=True).rename(columns={'2. 年齡':'年齡',"1. 性別：":'性別','3. 最高學歷':'最高學歷'})

ct1 = pd.crosstab(statData['2. 年齡'],[statData['3. 最高學歷'],statData["1. 性別："]],margins=True)

#獨立
result = pd.DataFrame(index=list(ct1.index))
girl_sum = pd.DataFrame(index=list(ct1.index))
girl_sum['All'] = [0]*len(list(ct1.index)) 
for i in ['博士','碩士','大學/大專','專科','高職']:
    girl_sum['All'] += ct1[i]['女性']
    girlRate = ct1[i]['女性']/(ct1[i]['女性']+ct1[i]['男性'])
    girlRate = girlRate.fillna(0)
    result[i+'0'] = ((girlRate*100).round(1).astype(str)+'%').replace('0.0%','-')
    result[i+'1'] = ct1[i]['女性']+ct1[i]['男性']

girlRate = girl_sum['All']/ct1['All']

girlRate = girlRate.fillna(0)
result['合計0'] = ((girlRate*100).round(1).astype(str)+'%').replace('0.0%','-')
result['合計1'] = ct1['All']

result = result.reindex(columns=['博士0','博士1','合計0','合計1','大學/大專0','大學/大專1','專科0','專科1','碩士0','碩士1','高職0','高職1'])

result.columns = [ sorted(['博士','碩士','大學/大專','專科','高職','合計']*2),['女性佔比','總人數']*6 ]

result = result[['合計','博士','碩士','大學/大專','專科','高職']]

resultCSV.append(result)
# ct1.to_csv('2.學歷.csv',encoding='utf_8_sig')

#畢業校系O

col = [ sorted(['工程與科技相關領域','非工程與科技相關領域']*4)+[''] , sorted(['國內學校','國外學校']*2)*2+[''] , ['女性','男性']*4+['小計']  ]

ct1 = pd.crosstab(statData['7. 是否正在從事工程與科技領域職務（含管理與學術研究）？'],[statData['4.\t最高學歷畢業校系'],statData['1. 性別：']],margins=True)
#ct1 = ct1/ct1['All'][-1]
#ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-')
#ct1 = ct1.drop('All')
ct1 = ct1.reindex(['是','否','All'])
ct1 = ct1.rename(columns={'All':'小計'},index={'是':'工程與科技職務','否':'非工程與科技職務','All':'人數'})
ct1.columns.names=["最高學歷","性別"]
ct1.index.name = ""
ct1 = ct1.reindex(columns=[['國內學校，工程與科技相關領域','國內學校，工程與科技相關領域','國外學校，工程與科技相關領域','國外學校，工程與科技相關領域',
'國內學校，非工程與科技相關領域','國內學校，非工程與科技相關領域','國外學校，非工程與科技相關領域','國外學校，非工程與科技相關領域','小計'],['女性','男性']*4+['']])
ct1.columns = col
temp = ct1.copy()

resultCSV.append(ct1)
# ct1.to_csv('3.最高學歷.csv',encoding='utf_8_sig')

ct1 = pd.crosstab(statData['7. 是否正在從事工程與科技領域職務（含管理與學術研究）？'],[statData['4.\t最高學歷畢業校系'],statData['1. 性別：']],margins=True)
ct1 = ct1/ct1['All'][-1]
ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-')
#ct1 = ct1.drop('All')
ct1 = ct1.reindex(['是','否','All'])
ct1 = ct1.rename(columns={'All':'小計'},index={'是':'工程與科技職務','否':'非工程與科技職務','All':'人數'})
ct1.columns.names=["最高學歷","性別"]
ct1.index.name = ""
ct1 = ct1.reindex(columns=[['國內學校，工程與科技相關領域','國內學校，工程與科技相關領域','國外學校，工程與科技相關領域','國外學校，工程與科技相關領域',
'國內學校，非工程與科技相關領域','國內學校，非工程與科技相關領域','國外學校，非工程與科技相關領域','國外學校，非工程與科技相關領域','小計'],['女性','男性']*4+['']])
ct1.columns = col
ct1.loc['人數'] = temp.loc['人數']

resultCSV.append(ct1)

#總年資O

df1 = statData[['1. 性別：','3. 最高學歷','4.\t最高學歷畢業校系','7. 是否正在從事工程與科技領域職務（含管理與學術研究）？','8. 您離開工程職務之最主要原因為何？請勾選最主要原因，至多二項。','(a) 工程與科技領域職務(年)', '(b) 非工程與科技領域職務','10.\t請問您目前服務的職務較接近下列何者？','請問目前受您管理的人員約＿＿人？','9.\t請問您目前的工作狀況為何？']]

delList =[]
for i in range(len(df1)):
    if( (not str(df1['(a) 工程與科技領域職務(年)'][i]).isdigit()) or (not str(df1['(b) 非工程與科技領域職務'][i]).isdigit())) :
        delList.append(i)
df2 = df1.drop(delList).reset_index(drop=True)
t1 = df2['(a) 工程與科技領域職務(年)'].fillna(0).astype(int)
t2 = df2['(b) 非工程與科技領域職務'].fillna(0).astype(int)
df2['總年資']=t1+t2

year = ['0~5年','6~10年','11~15年','16~20年','21~25年','26~30年','30年以上']

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

col = [ ['各年資小計','各年資小計','0~5年','0~5年','6~10年','6~10年','11~15年','11~15年','16~20年','16~20年','21~25年',
  '21~25年','26~30年','26~30年','30年以上','30年以上'],['女性占比','總人數']+['女性','男性']*(len(year)) ]

index = [ sorted(['工程與科技相關領域','非工程與科技相關領域']*2)+[''] , sorted(['國內學校','國外學校'])*2+['合計']  ] 


ct1 = pd.crosstab(df2['4.\t最高學歷畢業校系'],[df2['總年資'],df2['1. 性別：']],margins=True)
ct2 = pd.DataFrame(index=list(ct1.index))
ct3 = pd.crosstab(df2['4.\t最高學歷畢業校系'],[df2['總年資'],df2['1. 性別：']],normalize='columns')


girl_sum = ct1[year[0]]['女性']
for i in range(1,len(year)):
    girl_sum += ct1[year[i]]['女性']

rate = girl_sum/ct1['All']
ct2['女性占比']=((rate*100).round(1).astype(str)+'%').replace('0.0%','-')
ct2['總人數All'] = ct1['All']
ct1 = ct1.drop('All',axis=1)

ct3 = ((ct3*100).round(1).astype(str)+'%').replace('0.0%','-')
ct3.loc['All'] = ct1.loc['All']

ct1 = pd.concat([ct2,ct1],axis=1)
ct2 = pd.concat([ct2,ct3],axis=1)

ct1.columns = col
ct1.index = index
ct2.columns = col
ct2.index = index

ct1.index.name="最高學歷畢業校系"
ct1 = ct1.rename(index={'All':'合計'})
ct2.index.name="最高學歷畢業校系"
ct2 = ct2.rename(index={'All':'合計'})
#temp = list(ct1.index)
#ct1 = ct1.reindex(index=['合計']+temp[:-1])


resultCSV.append(ct1)
resultCSV.append(ct2)
# ct1.to_csv('4.總年資.csv',encoding='utf_8_sig')

#離開原因

df1 = statData[['1. 性別：','3. 最高學歷','4.\t最高學歷畢業校系','7. 是否正在從事工程與科技領域職務（含管理與學術研究）？','8. 您離開工程職務之最主要原因為何？請勾選最主要原因，至多二項。','10.\t請問您目前服務的職務較接近下列何者？','(a) 工程與科技領域職務(年)','9.\t請問您目前的工作狀況為何？']]

delList =[]
for i in range(len(df1)):
    if( not str(df1['(a) 工程與科技領域職務(年)'][i]).isdigit() ) :
        delList.append(i)
df1 = df1.drop(delList).reset_index(drop=True)

df1 = ct.group_table(df1,'7. 是否正在從事工程與科技領域職務（含管理與學術研究）？',['否'])
df_del = ct.group_table(df1,'(a) 工程與科技領域職務(年)',['0'])
df_del = ct.group_table(df_del,'4.\t最高學歷畢業校系',['國內學校，非工程與科技相關領域','國外學校，非工程與科技相關領域'])

delList = list(df_del.index)
delList.sort()
df1 = df1.drop(delList).reset_index(drop=True)
# df_del = ct.group_table(df1,'10.\t請問您目前服務的職務較接近下列何者？',['工程與科技領域之管理職務','工程與科技領域之非管理職務'])
# delList = list(df_del.index)
# delList.sort()
# df1 = df1.drop(delList).reset_index(drop=True)

df2 = ct.extraction(df1,'8. 您離開工程職務之最主要原因為何？請勾選最主要原因，至多二項。','您離開工程職務之最主要原因為何？').reset_index(drop=True)
num = len(df2)
# df_del = ct.group_table(df_del,"4.\t最高學歷畢業校系",['國外學校，非工程與科技相關領域','國內學校，非工程與科技相關領域'])

ct1 = pd.crosstab(df2['您離開工程職務之最主要原因為何？'],df2['1. 性別：'],margins=True)
# ct1 = ct1/temp
# ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-')
ct1 = ct1.drop("其他（請說明）",axis=0)
ct1 = ct1.rename(index={'公司組織調整（例如裁撤、縮編或退出本地市場）':'公司組織調整','All':'合計'},columns={'All':'小計'})
ct1.at['人數'] = [list(df1['1. 性別：']).count('女性'),list(df1['1. 性別：']).count('男性'),list(df1['1. 性別：']).count('女性')+list(df1['1. 性別：']).count('男性')]
index = list(ct1.index)
index.insert(0,index.pop())
ct1 = ct1.reindex(index=index)
ct1.columns.name='性別'
ct1.index.name="離開工程職務原因(%d人)"%(num)
ct1 = ct1.astype(int)

resultCSV.append(ct1)
# ct3.to_csv('5.離開工程職務最主要原因分析.csv',encoding='utf_8_sig')

#目前工作狀況O

ct1 = pd.crosstab(df1['9.\t請問您目前的工作狀況為何？'],df1['1. 性別：'],margins=True)
temp = int(ct1['All'][-1])
#ct1 = ct1/temp
#ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-')
ct1 = ct1.rename(index={'All':'人數'},columns={'All':'小計'})
ct1.columns.name='性別'
ct1.index.name="目前的工作狀況(%d人)"%(temp)

resultCSV.append(ct1)
# ct1.to_csv('7.目前工作狀況.csv',encoding='utf_8_sig')

#五年後O

temp = ['博士','碩士','大學/大專','專科','高職']
col2 = [ ['不分學歷小計']*2 , ['女性占比','總人數'] ]
df = statData[['1. 性別：','3. 最高學歷','11.\t您對於現任職務之五年後職涯發展的預期為何？請勾選最有可能的一項。']]

ct1 = pd.crosstab(df['11.\t您對於現任職務之五年後職涯發展的預期為何？請勾選最有可能的一項。'],[df['3. 最高學歷'],df['1. 性別：']],margins=True)
ct3 = pd.crosstab(df['11.\t您對於現任職務之五年後職涯發展的預期為何？請勾選最有可能的一項。'],[df['3. 最高學歷'],df['1. 性別：']],normalize='columns')
#ct1 = ct1/ct1['All'][-1]
#ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-')
ct1 = ct1[['博士','碩士','大學/大專','專科','高職','All']]

girl_sum = ct1[temp[0]]['女性']
for i in range(1,len(temp)):
    girl_sum += ct1[temp[i]]['女性']

ct2 = pd.DataFrame(columns=['女性占比','總人數'],index = list(ct1.index))

rate = girl_sum/ct1['All']
ct2['女性占比'] = ((rate*100).round(1).astype(str)+'%').replace('0.0%','-')
ct2['總人數'] = ct1['All']
ct2.columns = col2

ct3 = ((ct3*100).round(1).astype(str)+'%').replace('0.0%','-')
ct3.loc['All'] = ct1.loc['All']

ct1 = pd.concat([ct1,ct2],axis=1)
ct2 = pd.concat([ct3,ct2],axis=1)

ct1 = ct1.drop(['All'],axis=1)
ct1 = ct1[['不分學歷小計','博士','碩士','大學/大專','專科','高職']]

ct1.columns.names=['最高學歷','性別']
ct1.index.name="現任職務之五年後職涯發展的預期"
ct1 = ct1.rename(index={'轉換領域(轉至非工程與科技領域或轉入工程與科技領域)   (請說明原因)':'轉換領域(轉至非工程與科技領域或轉入工程與科技領域)','離職 (請說明原因)':'離職','All':'合計'},columns={'All':'小計'})

ct2 = ct2[['不分學歷小計','博士','碩士','大學/大專','專科','高職']]

ct2.columns.names=['最高學歷','性別']
ct2.index.name="現任職務之五年後職涯發展的預期"
ct2 = ct2.rename(index={'轉換領域(轉至非工程與科技領域或轉入工程與科技領域)   (請說明原因)':'轉換領域(轉至非工程與科技領域或轉入工程與科技領域)','離職 (請說明原因)':'離職','All':'合計'},columns={'All':'小計'})


resultCSV.append(ct1)
resultCSV.append(ct2)
# ct1.to_csv('8.對於現任職務之五年後職涯發展的預期.csv',encoding='utf_8_sig')

#配合因素O

df = statData[['1. 性別：','3. 最高學歷','12.\t請問您所預期的職涯發展，需要哪些配合因素來達成？請勾選最主要因素，至多二項。']]
# for i in range(len(df)):
#     if(not (df['12.\t請問您所預期的職涯發展，需要哪些配合因素來達成？請勾選最主要因素，至多二項。'][i].find('還不清楚（若勾選本項請勿再勾選其他項目）') == -1) ):
#         df.loc[i,'12.\t請問您所預期的職涯發展，需要哪些配合因素來達成？請勾選最主要因素，至多二項。'] = '還不清楚'
num = len(df)
df1 = ct.extraction(df,'12.\t請問您所預期的職涯發展，需要哪些配合因素來達成？請勾選最主要因素，至多二項。','請問您所預期的職涯發展，需要哪些配合因素來達成？').reset_index(drop=True)

ct1 = pd.crosstab(df1['請問您所預期的職涯發展，需要哪些配合因素來達成？'],df1['1. 性別：'],margins=True)
# ct1 = ct1/ct1['All'][-1]
# ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-')
ct1 = ct1.rename(index={'All':'合計'},columns={'All':'小計'})
ct1.at['人數'] = [list(df['1. 性別：']).count('女性'),list(df['1. 性別：']).count('男性'),list(df['1. 性別：']).count('女性')+list(df['1. 性別：']).count('男性')]
index = list(ct1.index)
index.insert(0,index.pop())
ct1 = ct1.reindex(index=index)
ct1.columns.name='性別'
ct1.index.name="預期的職涯發展配合因素(%d人)"%(num)

resultCSV.append(ct1)

# ct1.to_csv('9.預期的職涯發展需要哪些配合因素來達成.csv',encoding='utf_8_sig')
# ct2.to_csv('9.預期的職涯發展需要哪些配合因素來達成(性別).csv',encoding='utf_8_sig')

#有助福利('其他'刪除)O 有修改

df = statData[['1. 性別：','3. 最高學歷','13.\t哪些福利措施最有助於您留在工程與科技領域就業？（不論現在是否有需求皆可選擇，可複選，最多三項）']]
num = len(df)
df1 = ct.extraction(df,'13.\t哪些福利措施最有助於您留在工程與科技領域就業？（不論現在是否有需求皆可選擇，可複選，最多三項）','哪些福利措施最有助於您留在工程與科技領域就業？')
df1 = df1[~df1['哪些福利措施最有助於您留在工程與科技領域就業？'].isin([' 請勿再勾選其他選項 )','其他（請說明）'])].reset_index(drop=True)

ct2 = pd.crosstab(df1['哪些福利措施最有助於您留在工程與科技領域就業？'],df1['1. 性別：'],margins=True)
# ct2 = ct2/ct2['All'][-1]
# ct2 = ((ct2*100).round(1).astype(str)+'%').replace('0.0%','-')
ct2 = ct2.rename(index={'All':'合計','其他（請說明）':'其他','無，個人不需要這些福利 ( 若選此項':'無，個人不需要這些福利'},columns={'All':'小計'})
ct2.at['人數'] = [list(df['1. 性別：']).count('女性'),list(df['1. 性別：']).count('男性'),list(df['1. 性別：']).count('女性')+list(df['1. 性別：']).count('男性')]
index = list(ct2.index)
index.insert(0,index.pop())
ct2 = ct2.reindex(index=index)
# ct2 = ct2.drop('其他',axis=0)
ct2.columns.name='性別'
ct2.index.name="福利措施(%d人)"%(num)

resultCSV.append(ct2)

# ct1.to_csv('10.哪些福利措施最有助於您留在工程與科技領域就業.csv',encoding='utf_8_sig')
ct2.to_excel('10.哪些福利措施最有助於您留在工程與科技領域就業(性別).xlsx',encoding='utf_8_sig')

#單位福利O

df1 = statData[['1. 性別：','10.\t請問您目前服務的職務較接近下列何者？','13.\t哪些福利措施最有助於您留在工程與科技領域就業？（不論現在是否有需求皆可選擇，可複選，最多三項）',
       '14.\t您服務的單位或待業前服務的單位提供哪些職場相關措施？（可複選）',
       '15.\t您認為工程與科技領域最需要改善的性別議題有哪些？請勾選其中最重要的議題，至多三項。',]]
num = len(df1)
df3 =  ct.extraction(df1,'14.\t您服務的單位或待業前服務的單位提供哪些職場相關措施？（可複選）','您服務的單位或待業前服務的單位提供哪些職場相關措施？')

ct1 = pd.crosstab(df3['您服務的單位或待業前服務的單位提供哪些職場相關措施？'],df3['1. 性別：'],margins=True)
# ct1 = ct1/ct1['All'][-1]
# ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-').drop([' 請勿再勾選其他選項 )','其他，請說明',],axis=0)
ct1 = ct1.rename(index={'All':'合計','無，沒有提供 ( 若選此項':'無，沒有提供'},columns={'All':'小計'})
ct1.at['人數'] = [list(df1['1. 性別：']).count('女性'),list(df1['1. 性別：']).count('男性'),list(df1['1. 性別：']).count('女性')+list(df1['1. 性別：']).count('男性')]
index = list(ct1.index)
index.insert(0,index.pop())
ct1 = ct1.reindex(index=index)
ct1.columns.name='性別'
ct1.index.name='您服務的單位或待業前服務的單位提供哪些職場相關措施？(%d人)'%(num)

resultCSV.append(ct1)
# ct1.to_csv("11.您服務的單位或待業前服務的單位提供哪些職場相關措施.csv",encoding='utf_8_sig')

#目前服務的職務O

ct1 = pd.crosstab(statData['10.\t請問您目前服務的職務較接近下列何者？'],statData['1. 性別：'],margins=True)
num = len(statData)
ct1 = ct1.rename(index={'All':'合計','其他（請說明）':'其他'},columns={'All':'小計'})
ct1.at['人數'] = [list(statData['1. 性別：']).count('女性'),list(statData['1. 性別：']).count('男性'),list(statData['1. 性別：']).count('女性')+list(statData['1. 性別：']).count('男性')]
index = list(ct1.index)
index.insert(0,index.pop())
ct1 = ct1.reindex(index=index)
ct1.columns.name='性別'
ct1.index.name='您目前服務的職務較接近下列何者？(%d人)'%(num)

resultCSV.append(ct1)

#需要改善O 有修改

num = len(statData)
df4 = ct.extraction(statData,'15.\t您認為工程與科技領域最需要改善的性別議題有哪些？請勾選其中最重要的議題，至多三項。','您認為工程與科技領域最需要改善的性別議題有哪些？')
df4 = df4[~df4['您認為工程與科技領域最需要改善的性別議題有哪些？'].isin([' 請勿再勾選其他選項 )','其他（請說明）'])].reset_index(drop=True)

ct1 = pd.crosstab(df4['您認為工程與科技領域最需要改善的性別議題有哪些？'],df4['1. 性別：'],margins=True)
# ct1 = ct1/ct1['All'][-1]
# ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-')
ct1 = ct1.rename(index={'All':'合計','無，沒有提供 ( 若選此項':'無，沒有提供','不了解，沒有想法  ( 若選此項':'不了解，沒有想法','無，沒必要改善  ( 若選此項':'無，沒必要改善'},columns={'All':'小計'})
ct1.at['人數'] = [list(statData['1. 性別：']).count('女性'),list(statData['1. 性別：']).count('男性'),list(statData['1. 性別：']).count('女性')+list(statData['1. 性別：']).count('男性')]
index = list(ct1.index)
index.insert(0,index.pop())
ct1 = ct1.reindex(index=index)
ct1.columns.name='性別'
ct1.index.name='您認為工程與科技領域最需要改善的性別議題有哪些？(%d人)'%(num)

resultCSV.append(ct1)
ct1.to_excel("您認為工程與科技領域最需要改善的性別議題有哪些.xlsx",encoding='utf_8_sig')

#性別在工程與科技領域的求學階段有否差異

#16 有修改
df = statData.dropna(subset=['16.\t在您或您認識的女性身上，性別在工程與科技領域的求學階段有否差異？'])
ct1 = pd.crosstab(df['16.\t在您或您認識的女性身上，性別在工程與科技領域的求學階段有否差異？'],df['1. 性別：'],margins=True)
# ct1 = ct1/ct1['All'][-1]
# ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-')
ct1 = ct1.rename(index={'All':'人數','有，我有觀察到差異：':'有，我有觀察到差異'},columns={'All':'小計'})
ct1.index.name = '在您或您認識的女性身上，性別在工程與科技領域的求學階段有否差異？'
ct1.columns.name = '性別'
resultCSV.append(ct1)

df = ct.group_table(statData,'16.\t在您或您認識的女性身上，性別在工程與科技領域的求學階段有否差異？',['有，我有觀察到差異：'])
title = ['較用功或成績較好', '較擅長數理學科','動手實驗或實習的機會較少', '遇到挫折時較易放棄', '較不敢接受挑戰', '較容易情緒起伏', '較受人際問題干擾','較容易為情所困']
col = [ sorted(title*3),['女性','男性','小計']*len(title) ]
df1 = df[['1. 性別：']+title].reset_index(drop=True)
ct2 = []
title.sort()
for i in range(len(title)):
    ct2.append( pd.crosstab(df1[title[i]],df1['1. 性別：'],margins=True) )

    for j in range(len(list(ct2[i].index))-1):
        ct2[i].loc[list(ct2[i].index)[j]] = ct2[i].loc[list(ct2[i].index)[j]]/ct2[i].loc['All']

    temp = ct2[i].loc['All']
    ct2[i] = ((ct2[i]*100).round(1).astype(str)+'%').replace('0.0%','-')
    ct2[i].loc['All'] = temp

ct2 = pd.concat(ct2,axis=1)
ct2.columns=col
ct2 = ct2.rename(index={'All':'人數'})
ct2.index.name = '在您或您認識的女性身上，性別在工程與科技領域的求學階段有否差異？'
resultCSV.append(ct2)
ct2.T.to_excel("12.在您或您認識的女性身上，性別在工程與科技領域的求學階段有否差異.xlsx",encoding='utf_8_sig')

#17 有修改
df = statData.dropna(subset=['17.\t在您或您認識的工程與科技領域女性身上，性別在求職時的職務選擇方面有否差異？'])
ct1 = pd.crosstab(df['17.\t在您或您認識的工程與科技領域女性身上，性別在求職時的職務選擇方面有否差異？'],df['1. 性別：'],margins=True)
# ct1 = ct1/ct1['All'][-1]
# ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-')
ct1 = ct1.rename(index={'All':'人數','有，我有觀察到差異：':'有，我有觀察到差異'},columns={'All':'小計'})
ct1.index.name = '在您或您認識的工程與科技領域女性身上，性別在求職時的職務選擇方面有否差異？'
ct1.columns.name = '性別'
resultCSV.append(ct1)

df = ct.group_table(statData,'17.\t在您或您認識的工程與科技領域女性身上，性別在求職時的職務選擇方面有否差異？',['有，我有觀察到差異：'])
title = ['選擇內勤為主的職務', '選擇無須出差的職務', '選擇無須應酬的職務', '選擇無須輪班或值夜班的職務', '選擇可兼顧家庭的職務','選擇較有升遷或加薪機會的職務']
col = [ sorted(title*3),['女性','男性','小計']*len(title) ]
df1 = df[['1. 性別：']+title].reset_index(drop=True)
ct2 = []
title.sort()
for i in range(len(title)):
    ct2.append( pd.crosstab(df1[title[i]],df1['1. 性別：'],margins=True) )

    for j in range(len(list(ct2[i].index))-1):
        ct2[i].loc[list(ct2[i].index)[j]] = ct2[i].loc[list(ct2[i].index)[j]]/ct2[i].loc['All']

    temp = ct2[i].loc['All']
    ct2[i] = ((ct2[i]*100).round(1).astype(str)+'%').replace('0.0%','-')
    ct2[i].loc['All'] = temp

# ct2[3]=ct.Insert_row(2,ct2[3],['0','0','0'])
# ct2[3].index=['女性較有此傾向', '無差異', '男性較有此傾向', 'All']

ct2 = pd.concat(ct2,axis=1)
ct2.columns=col
ct2 = ct2.rename(index={'All':'人數'})
ct2.index.name = '在您或您認識的工程與科技領域女性身上，性別在求職時的職務選擇方面有否差異？'
resultCSV.append(ct2)

ct2.T.to_excel("在您或您認識的工程與科技領域女性身上，性別在求職時的職務選擇方面有否差異.xlsx",encoding='utf_8_sig')

#18
df = df[['1. 性別：','18.\t就您的認知，不同性別在工程與科技領域的求職難易度有否差異？',
    '19.\t在您或您認識的女性身上，性別在工程與科技領域的就業過程中存在的差異：（複選，同意即打勾）',
    '20.\t對您的經驗，不同性別的工程與科技領域主管是否有領導風格的差異（指非因個人而是因為性別所致之差異）？',
    '21.\t您個人比較偏好與何種性別工程與科技領域的主管共事？',
    '22.\t不同性別工程與科技領域的主管領導風格差異主要顯現在哪些層面？（可複選至多三項）']].reset_index(drop=True)

ct1 = pd.crosstab(df['18.\t就您的認知，不同性別在工程與科技領域的求職難易度有否差異？'],df['1. 性別：'],margins=True)
# ct1 = ct1/ct1['All'][-1]
# ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-')
ct1 = ct1.rename(index={'All':'人數'},columns={'All':'小計'})
ct1.index.name = '不同性別在工程與科技領域的求職難易度有否差異？'
ct1.columns.name = '性別'
resultCSV.append(ct1)

#19 有修改
num = len(df)
df1 = ct.extraction(df,'19.\t在您或您認識的女性身上，性別在工程與科技領域的就業過程中存在的差異：（複選，同意即打勾）','19.在您或您認識的女性身上，性別在工程與科技領域的就業過程中存在的差異').reset_index(drop=True)
df1 = df1[~df1['19.在您或您認識的女性身上，性別在工程與科技領域的就業過程中存在的差異'].isin([' 請勿再勾選其他選項 )'])]
ct1 = pd.crosstab(df1['19.在您或您認識的女性身上，性別在工程與科技領域的就業過程中存在的差異'],df1['1. 性別：'],margins=True)
# ct1 = ct1/ct1['All'][-1]
# ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-')
ct1 = ct1.rename(index={'All':'合計','其他，':'其他','沒有差異  ( 若選此項':'沒有差異','沒有接觸經驗，無法回答  ( 若選此項':'沒有接觸經驗，無法回答'},columns={'All':'小計'})
ct1.at['人數'] = [list(df['1. 性別：']).count('女性'),list(df['1. 性別：']).count('男性'),list(df['1. 性別：']).count('女性')+list(df['1. 性別：']).count('男性')]
# ct1 = ct1.drop([' 請勿再勾選其他選項 )'],axis=0)
ct1.index.name = '在您或您認識的女性身上，性別在工程與科技領域的就業過程中存在的差異？(%d人)'%(num)
ct1.columns.name = '性別'
temp = list(ct1.index)
temp.insert(11,temp[0])
temp = temp[1:]
ct1 = ct1.reindex(index=temp)
index = list(ct1.index)
index.insert(0,index.pop())
ct1 = ct1.reindex(index=index)
resultCSV.append(ct1)

ct1.to_excel("在您或您認識的女性身上，性別在工程與科技領域的就業過程中存在的差異.xlsx",encoding='utf_8_sig')

#20

df1 = df.dropna(subset=['20.\t對您的經驗，不同性別的工程與科技領域主管是否有領導風格的差異（指非因個人而是因為性別所致之差異）？'])
ct1 = pd.crosstab(df1['20.\t對您的經驗，不同性別的工程與科技領域主管是否有領導風格的差異（指非因個人而是因為性別所致之差異）？'],df1['1. 性別：'],margins=True)
# ct1 = ct1/ct1['All'][-1]
# ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-')
ct1 = ct1.rename(index={'All':'人數'},columns={'All':'小計'})
ct1.index.name = '對您的經驗，不同性別的工程與科技領域主管是否有領導風格的差異？'
ct1.columns.name = '性別'
resultCSV.append(ct1)

#21

df1 = ct.group_table(df,'20.\t對您的經驗，不同性別的工程與科技領域主管是否有領導風格的差異（指非因個人而是因為性別所致之差異）？',['是，有差異'])
ct1 = pd.crosstab(df1['21.\t您個人比較偏好與何種性別工程與科技領域的主管共事？'],df1['1. 性別：'],margins=True)
# ct1 = ct1/ct1['All'][-1]
# ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-')
ct1 = ct1.rename(index={'All':'人數'},columns={'All':'小計'})
ct1.index.name = '您個人比較偏好與何種性別工程與科技領域的主管共事？'
ct1.columns.name = '性別'
ct1 = ct1.reindex(index=['女性', '男性','無性別偏好' ,'人數'])
resultCSV.append(ct1)

#22

num = len(df1)
df2 = ct.extraction(df1,'22.\t不同性別工程與科技領域的主管領導風格差異主要顯現在哪些層面？（可複選至多三項）','22.\t不同性別工程與科技領域的主管領導風格差異主要顯現在哪些層面？').reset_index(drop=True)
ct1 = pd.crosstab(df2['22.\t不同性別工程與科技領域的主管領導風格差異主要顯現在哪些層面？'],df2['1. 性別：'],margins=True)
# ct1 = ct1/ct1['All'][-1]
# ct1 = ((ct1*100).round(1).astype(str)+'%').replace('0.0%','-')
ct1 = ct1.rename(index={'All':'合計'},columns={'All':'小計'})
ct1.at['人數'] = [list(df1['1. 性別：']).count('女性'),list(df1['1. 性別：']).count('男性'),list(df1['1. 性別：']).count('女性')+list(df1['1. 性別：']).count('男性')]
ct1.index.name = '不同性別工程與科技領域的主管領導風格差異主要顯現在哪些層面？(%d人)'%(num)
ct1.columns.name = '性別'
temp = list(ct1.index)
temp.insert(15,temp[0])
temp = temp[1:]
ct1 = ct1.reindex(index=temp)
index = list(ct1.index)
index.insert(0,index.pop())
ct1 = ct1.reindex(index=index)
resultCSV.append(ct1)

#---------------------------------------------------------------------------------------#

name = ['年齡','最高學歷','最高畢業','最高畢業(比例)','總年資','總年資(比例)','離開工程職務最主要原因分析','目前工作狀況','對於現任職務之五年後職涯發展的預期','對於現任職務之五年後職涯發展的預期(比例)','預期的職涯發展需要哪些配合因素來達成',
        '哪些福利措施最有助於您留在工程與科技領域就業','您服務的單位或待業前服務的單位提供哪些職場相關措施','您目前服務的職務較接近下列何者',
        '您認為工程與科技領域最需要改善的性別議題有哪些','女性身上求學階段有否差異(性別)','女性身上求學階段有否差異','工程與科技領域女性身上職務選擇有否差異(性別)','工程與科技領域女性身上職務選擇有否差異',
        '不同性別在工程與科技領域的求職難易度有否差異(性別)','性別在求職時的職務選擇方面有否差異','不同性別的工程與科技領域主管是否有領導風格的差異(性別)','您個人比較偏好與何種性別工程與科技領域的主管共事(性別)',
        '不同性別工程與科技領域的主管領導風格差異主要在哪些層面(性別)']

with pd.ExcelWriter('個人版分析.xlsx') as writer:
    for i in range(len(name)):
        resultCSV[i].to_excel(writer,sheet_name=name[i],encoding='utf_8_sig')
