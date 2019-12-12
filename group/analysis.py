import pandas as pd
import os
import re
import numpy as np

os.chdir("/Users/cilab/PartTime_PythonAnalysis/group")
#os.chdir(r"C:\Users\Naichen\Documents\GitHub\stu00608.github.io\PartTime_PythonAnalysis\group")
os.getcwd()

import ct_tool as ct

#記得路徑不同電腦要改

#研官產
classification = [] 
data = pd.read_excel("004.xlsx",sheet_name=2,usecols=[1]) #產
data = list(data['單位'])
data += ['國際聯合科技','國際聯合科技股份有限公司','鼎原科技股份有限公司','國際聯合科技','國際聯合科技股份有限公司','世紀離岸風電設備股份有限公司','主動元件事業處','連展科技股份有限公司連接器事業處研發部','呈峰營造','台科大第一宿舍拆除重建工程','泰誠發展營造股份有限公司','點晶科技股份有限公司','環球水泥股份有限公司','中鋼','聚和國際','國光生物科技股份有限公司','鼎漢國際工程顧問公司','台化工務部設計組','台灣自來水公司中區工程處']
classification.append(data)
data = pd.read_excel("004.xlsx",sheet_name=1,usecols=[0]) #研
data = list(data['單位'])
['財團法人中華顧問工程司','財團法人臺灣營建研究院','財團法人國家實驗研究院','財團法人食品工業發展研究所','財團法人國家同步輻射研究中心']
classification.append(data)
data = pd.read_excel("004.xlsx",sheet_name=0,usecols=[1]) #官
data = list(data['單位'])
data += ['交通部高速公路局中區養護工程分局','新北市政府水利局','交通部鐵道局']
classification.append(data)
for i in range(3):
    for j in range(len(classification[i])):
        classification[i][j] = classification[i][j].replace('\u3000','').strip()

statData = pd.read_excel("003.xlsx",skiprows=5)
#statData = statData.dropna(axis=1,how='all')
statData['46-55歲.re']=np.zeros(len(statData.index))
delList = []
unavailableUnit = []
for i in range(len(statData['1.單位名稱'])):
    if(   str(statData['1.單位名稱'][i]).find('test') != -1 ):
        delList.append(i)
        unavailableUnit.append(statData['1.單位名稱'][i])
    elif( str(statData['1.單位名稱'][i]).find('111111111111') != -1 ):
        delList.append(statData['1.單位名稱'][i])
    elif( str(statData['1.單位名稱'][i]).find('測試') != -1 ):
        delList.append(statData['1.單位名稱'][i])

statData = statData.drop(delList).reset_index(drop=True)


os.chdir("/Users/cilab/PartTime_PythonAnalysis/group/outputs")
#os.chdir(r"C:\Users\Naichen\Documents\GitHub\stu00608.github.io\PartTime_PythonAnalysis\group\outputs")
os.getcwd()

#土木營建  建築、都市規劃 電子電機 資訊通訊 化工材料 生技醫工 環工綠能 機械 其他

reportData = pd.DataFrame(columns=['回報資料'])
lost_company = pd.DataFrame(columns=['其他歸類單位'])
allUnit = len(list(statData.index))


#數字檢測(把float獨立出來檢查)
detectList = [ '2.單位總員工人數','土木營建領域專長 男性 專任 人數','建築、都市規劃領域專長 男性 專任 人數','電子電機領域專長 男性 專任 人數','資訊通訊領域專長 男性 專任 人數','化工材料領域專長 男性 專任 人數','生技醫工領域專長 男性 專任 人數','環工綠能領域專長 男性 專任 人數','機械領域專長 男性 專任 人數','其他領域專長 男性 專任 人數','非工程與科技領域專長 男性 專任 人數','土木營建領域專長 女性 專任 人數','建築、都市規劃領域專長 女性 專任 人數','電子電機領域專長 女性 專任 人數','資訊通訊領域專長 女性 專任 人數','化工材料領域專長 女性 專任 人數','生技醫工領域專長 女性 專任 人數','環工綠能領域專長 女性 專任 人數','機械領域專長 女性 專任 人數','其他領域專長 女性 專任 人數','非工程與科技領域專長 女性 專任 人數']+['106年 請假人次','107年 請假人次','108年 請假人次']+['35歲以下','36-45歲','46-55歲','56-65歲','66歲以上']+['服務年資1~5年','服務年資6~10年','服務年資11~15年','服務年資16~20年','服務年資21~25年','服務年資25年以上']+['管理職年資1~5年','管理職年資6~10年','管理職年資11~15年','管理職年資16~20年','管理職年資21~25年','管理職年資25年以上']+list(['服務年資1~5年.'+str(i) for i in range(1,18)])+list(['服務年資6~10年.'+str(i) for i in range(1,18)])+list(['服務年資11~15年.'+str(i) for i in range(1,18)])+list(['服務年資16~20年.'+str(i) for i in range(1,18)])+list(['服務年資21~25年.'+str(i) for i in range(1,18)])+list(['服務年資25年以上.'+str(i) for i in range(1,18)])+list(['管理職年資1~5年.'+str(i) for i in range(1,18)])+list(['管理職年資6~10年.'+str(i) for i in range(1,18)])+list(['管理職年資11~15年.'+str(i) for i in range(1,18)])+list(['管理職年資16~20年.'+str(i) for i in range(1,18)])+list(['管理職年資21~25年.'+str(i) for i in range(1,18)])+list(['管理職年資25年以上.'+str(i) for i in range(1,18)])+list(['35歲以下.'+str(i) for i in range(1,108)])+list(['36-45歲.'+str(i) for i in range(1,108)])+list(['46-55歲.'+str(i) for i in range(1,108)])+list(['56-65歲.'+str(i) for i in range(1,108)])+list(['66歲以上.'+str(i) for i in range(1,108)])+list(['106年 請假人次.'+str(i) for i in range(1,4)])+list(['107年 請假人次.'+str(i) for i in range(1,4)])+list(['108年 請假人次.'+str(i) for i in range(1,4)])

boolList = statData[detectList].isna()
for columns in detectList:     
    for i in range(len(statData.index)):
        if(not boolList[columns][i]):
            temp = statData[columns][i]
            if(str(temp).isdigit()):
                temp = int(temp)
            elif(re.sub('\D','',str(temp))==''):
                temp = 0
            elif(str(temp).find('.')>=0):
                temp = int(temp)
            else:
                temp = int(re.sub('\D','',str(temp)))
            statData.at[i,columns] = temp
            #statData.set_value(i,columns,temp)
            # 1.0 10 'xxx' '10x'
        else:
            statData.at[i,columns] = 0

j=k=m=0
existList = []
existDict = {}
alreadyExistList = []
alreadyExistDict = {}
for i in range(len(statData['1.單位名稱'])):
    if(statData['1.單位名稱'][i] in existList):
        if(alreadyExistDict.get(statData['1.單位名稱'][i],0) == 0):
            alreadyExistDict[statData['1.單位名稱'][i]] = []
        alreadyExistDict[statData['1.單位名稱'][i]] += [i] #第二次以後出現的index
        alreadyExistList.append(statData['1.單位名稱'][i])
        continue
    else:
        existDict[statData.at[i,'1.單位名稱']] = i #第一次出現的index
        existList.append(statData.at[i,'1.單位名稱'])

nameCheck_statData = statData.copy()

water = ['經濟部水利署']
road = ['公路總局','公路局第五區養護工程處-阿里山工務段','新化','埔里工務段','西濱北工處第七工務段','西部濱海公路南區臨時工程處第六工務段','第二區養護工程處台中段','新化工務段','交通部第五區養護工程處新化工務段','西濱北工處工務段','蘇花公路改善工程處勞安科','埔里工務段','埔里工務段','西部濱海公路北區臨時工程處第一工務段','港灣工程部']
electronic = ['台灣電力','臺灣電力','台電','臺電']

for i in range(len(statData['1.單位名稱'])):

    for item in water:
        if(   str(statData['1.單位名稱'][i]).find(item) != -1 ):
            statData.at[i,'1.單位名稱']='水利署'
            j+=1
    for item in road:
        if(   str(statData['1.單位名稱'][i]).find(item) != -1 ):
            statData.at[i,'1.單位名稱']='公路總局'
            k+=1
    for item in electronic:
        if(   str(statData['1.單位名稱'][i]).find(item) != -1 ):
            statData.at[i,'1.單位名稱']='台灣電力股份有限公司'
            m+=1

delList = []
for item in alreadyExistDict:
    delList += alreadyExistDict[item][:-1]
    delList.append(existDict[item])
delList.sort()
delUnit = []
temp = list(statData['1.單位名稱'])
for item in delList:
    delUnit.append(temp[item])
statData = statData.drop(delList,axis=0).reset_index(drop=True)
nameCheck_statData = nameCheck_statData.drop(delList,axis=0).reset_index(drop=True)

reportData.at['水利署總回填數'] = j
reportData.at['水利署重複回填數'] = delUnit.count('水利署')
reportData.at['水利署有效回填數'] = j-delUnit.count('水利署')

reportData.at['公路總局總回填數'] = k
reportData.at['公路總局重複回填數'] = delUnit.count('公路總局')
reportData.at['公路總局有效回填數'] = k-delUnit.count('公路總局')

reportData.at['台灣電力股份有限公司總回填數'] = m
reportData.at['台灣電力股份有限公司重複回填數'] = delUnit.count('台灣電力股份有限公司')
reportData.at['台灣電力股份有限公司有效回填數'] = m-delUnit.count('台灣電力股份有限公司')

reportData.at['其他機構重複回填數'] = len(delUnit)-delUnit.count('水利署')-delUnit.count('公路總局')-delUnit.count('台灣電力股份有限公司')

reportData.at['其他資訊'] = '---'
reportData.at['填答單位數'] = allUnit
reportData.at['有效單位數'] = len(list(statData.index))

category = ['研小計','官小計','產小計']

i=0
delList = []
for j in range(len(statData['1.單位名稱'])):
    if( classification[0].count(statData['1.單位名稱'][j])>=1 ):
        statData.at[j,'1.單位名稱']='產小計'
    elif(classification[1].count(statData['1.單位名稱'][j])>=1):
        statData.at[j,'1.單位名稱']='研小計'
    elif(classification[2].count(statData['1.單位名稱'][j])>=1):
        statData.at[j,'1.單位名稱']='官小計'
    else:
        lost_company.at[i] = statData['1.單位名稱'][j]
        delList.append(statData['1.單位名稱'][j])
        i+=1
        #statData.at[j,'1.單位名稱']='其他'
        


num = ['50人以下','50~250人','250人以上']

# for i in range(len(statData)):
#     v1 = int(statData['2.單位總員工人數'][i])
#     if( v1<=49 and v1>=1) :
#         statData.at[i,'2.單位總員工人數'] = num[0]
#     elif( v1>49 and v1<250 ):
#         statData.at[i,'2.單位總員工人數'] = num[1]
#     elif( v1>249 ):
#         statData.at[i,'2.單位總員工人數'] = num[2]
#     else:
#         statData.at[i,'2.單位總員工人數']=None

#---main---#

colname = [ ['女性占比','女性占比','女性占比','女性占比','女性占比','女性占比','總人數','總人數','總人數','總人數','總人數','總人數'],['1~5年','6~10年','11~15年','16~20年','21~25年','25年以上','1~5年',
 '6~10年','11~15年','16~20年','21~25年','25年以上'] ]

colname2 = [ ['女性佔比']*15+['總人數']*15,(['初階管理職']*5+['中階管理職']*5+['高階管理職']*5)*2,['35歲以下','36-45歲','46-55歲','56-65歲','66歲以上']*6 ]

colname3 = [ ['土木營建','土木營建','建築、都市規劃','建築、都市規劃','電子電機','電子電機','資訊通訊','資訊通訊','化工材料','化工材料','生技醫工','生技醫工','環工綠能','環工綠能','機械','機械','其他','其他','非工程與科技領域','非工程與科技領域'],
            ['女性佔比','總人數']*10 ]

colname4 = [ ['女性佔比']*15+['總人數']*15,(['初階專業職']*5+['中階專業職']*5+['高階專業職']*5)*2,['35歲以下','36-45歲','46-55歲','56-65歲','66歲以上']*6 ]

theme = ['土木營建','建築、都市規劃','電子電機','資訊通訊','化工材料','生技醫工','環工綠能','機械','其他']


#All

df = statData[['1.單位名稱','2.單位總員工人數','土木營建領域專長 男性 專任 人數','建築、都市規劃領域專長 男性 專任 人數','電子電機領域專長 男性 專任 人數','資訊通訊領域專長 男性 專任 人數','化工材料領域專長 男性 專任 人數',
 '生技醫工領域專長 男性 專任 人數','環工綠能領域專長 男性 專任 人數','機械領域專長 男性 專任 人數','其他領域專長 男性 專任 人數','非工程與科技領域專長 男性 專任 人數',
 '土木營建領域專長 女性 專任 人數','建築、都市規劃領域專長 女性 專任 人數','電子電機領域專長 女性 專任 人數','資訊通訊領域專長 女性 專任 人數','化工材料領域專長 女性 專任 人數',
 '生技醫工領域專長 女性 專任 人數','環工綠能領域專長 女性 專任 人數','機械領域專長 女性 專任 人數','其他領域專長 女性 專任 人數','非工程與科技領域專長 女性 專任 人數']]
col = ['土木營建領域專長 男性 專任 人數','建築、都市規劃領域專長 男性 專任 人數','電子電機領域專長 男性 專任 人數','資訊通訊領域專長 男性 專任 人數','化工材料領域專長 男性 專任 人數',
 '生技醫工領域專長 男性 專任 人數','環工綠能領域專長 男性 專任 人數','機械領域專長 男性 專任 人數','其他領域專長 男性 專任 人數','非工程與科技領域專長 男性 專任 人數',
 '土木營建領域專長 女性 專任 人數','建築、都市規劃領域專長 女性 專任 人數','電子電機領域專長 女性 專任 人數','資訊通訊領域專長 女性 專任 人數','化工材料領域專長 女性 專任 人數',
 '生技醫工領域專長 女性 專任 人數','環工綠能領域專長 女性 專任 人數','機械領域專長 女性 專任 人數','其他領域專長 女性 專任 人數','非工程與科技領域專長 女性 專任 人數']
result = pd.DataFrame(columns=col)
for name in num:
    try:
        df1 = ct.group_table(df,'2.單位總員工人數',[name])
    except:
        continue
    df1 = pd.DataFrame(data=df1[col].sum(),columns=[name])
    result.at[name]=list(df1[name])

for name in category:
    try:
        df1 = ct.group_table(df,'1.單位名稱',[name])
    except:
        continue
    df1 = pd.DataFrame(data=df1[col].sum(),columns=[name])
    result.at[name]=list(df1[name])



result = result.astype(int)

boy = result[col[:10]]
boy.columns = theme+['非工程與科技領域']
girl = result[col[10:]]
girl.columns = theme+['非工程與科技領域']
member = (boy+girl).astype(int)
girlRate = (girl/member)*100
girlRate = girlRate.fillna(0)
temp = (girl/member).fillna(0)
member.at['產官研合計'] = list(member.sum())
member = member.astype(int)
girlRate.at['產官研合計'] = list(girlRate.sum())
girlRate = (girlRate.round(1).astype(str) + '%').replace('0.0%','-')

result = pd.DataFrame(index=list(girlRate.index))
for i in range(len(theme)):
    result[theme[i]+str(0)]=girlRate[theme[i]]
    result[theme[i]+str(1)]=member[theme[i]]

result['非工程與科技領域'+str(0)]=girlRate['非工程與科技領域']
result['非工程與科技領域'+str(1)]=member['非工程與科技領域']

temp = pd.concat([temp,member],axis=1,sort=True)
member_analysis_n = temp
result.columns = colname3
result.columns.names = ['','']
member_analysis = result.copy()
index = list(member_analysis.index)
index.insert(0,index.pop())
member_analysis = member_analysis.reindex(index=index)
#result1.to_csv('總人數分析.csv',encoding='utf_8_sig')

#------------------------#

i=0
themeList = [ ['1.單位名稱','2.單位總員工人數',theme[i],theme[i]+'領域專長 男性 專任 人數',theme[i]+'領域專長 女性 專任 人數']+['服務年資1~5年','服務年資6~10年','服務年資11~15年','服務年資16~20年','服務年資21~25年','服務年資25年以上']+['服務年資1~5年.'+str(i*2+1),'服務年資6~10年.'+str(i*2+1),'服務年資11~15年.'+str(i*2+1),'服務年資16~20年.'+str(i*2+1),'服務年資21~25年.'+str(i*2+1),'服務年資25年以上.'+str(i*2+1)]+['管理職年資1~5年','管理職年資6~10年','管理職年資11~15年','管理職年資16~20年','管理職年資21~25年','管理職年資25年以上']+['管理職年資1~5年.'+str(i*2+1),'管理職年資6~10年.'+str(i*2+1),'管理職年資11~15年.'+str(i*2+1),'管理職年資16~20年.'+str(i*2+1),'管理職年資21~25年.'+str(i*2+1),'管理職年資25年以上.'+str(i*2+1)]+['35歲以下','36-45歲','46-55歲','56-65歲','66歲以上']+['35歲以下.'+str(i*2+1),'36-45歲.'+str(i*2+1),'46-55歲.'+str(i*2+1),'56-65歲.'+str(i*2+1),'66歲以上.'+str(i*2+1)]+['35歲以下.'+str(i*2+18),'36-45歲.'+str(i*2+18),'46-55歲.'+str(i*2+18),'56-65歲.'+str(i*2+18),'66歲以上.'+str(i*2+18)]+['35歲以下.'+str(i*2+18+1),'36-45歲.'+str(i*2+18+1),'46-55歲.'+str(i*2+18+1),'56-65歲.'+str(i*2+18+1),'66歲以上.'+str(i*2+18+1)]+['35歲以下.'+str(i*2+36),'36-45歲.'+str(i*2+36),'46-55歲.'+str(i*2+36),'56-65歲.'+str(i*2+36),'66歲以上.'+str(i*2+36)]+['35歲以下.'+str(i*2+36+1),'36-45歲.'+str(i*2+36+1),'46-55歲.'+str(i*2+36+1),'56-65歲.'+str(i*2+36+1),'66歲以上.'+str(i*2+36+1)]+['35歲以下.'+str(i*2+54),'36-45歲.'+str(i*2+54),'46-55歲.'+str(i*2+54),'56-65歲.'+str(i*2+54),'66歲以上.'+str(i*2+54)]+['35歲以下.'+str(i*2+54+1),'36-45歲.'+str(i*2+54+1),'46-55歲.'+str(i*2+54+1),'56-65歲.'+str(i*2+54+1),'66歲以上.'+str(i*2+54+1)]+['35歲以下.'+str(i*2+72),'36-45歲.'+str(i*2+72),'46-55歲.'+str(i*2+72),'56-65歲.'+str(i*2+72),'66歲以上.'+str(i*2+72)]+['35歲以下.'+str(i*2+72+1),'36-45歲.'+str(i*2+72+1),'46-55歲.'+str(i*2+72+1),'56-65歲.'+str(i*2+72+1),'66歲以上.'+str(i*2+72+1)]+['35歲以下.'+str(i*2+90),'36-45歲.'+str(i*2+90),'46-55歲.'+str(i*2+90),'56-65歲.'+str(i*2+90),'66歲以上.'+str(i*2+90)]+['35歲以下.'+str(i*2+90+1),'36-45歲.'+str(i*2+90+1),'46-55歲.'+str(i*2+90+1),'56-65歲.'+str(i*2+90+1),'66歲以上.'+str(i*2+90+1)] ] 

for i in range(1,len(theme)):
    themeList.append(['1.單位名稱','2.單位總員工人數',theme[i],theme[i]+'領域專長 男性 專任 人數',theme[i]+'領域專長 女性 專任 人數']+['服務年資1~5年.'+str(i*2),'服務年資6~10年.'+str(i*2),'服務年資11~15年.'+str(i*2),'服務年資16~20年.'+str(i*2),'服務年資21~25年.'+str(i*2),'服務年資25年以上.'+str(i*2)]+['服務年資1~5年.'+str(i*2+1),'服務年資6~10年.'+str(i*2+1),'服務年資11~15年.'+str(i*2+1),'服務年資16~20年.'+str(i*2+1),'服務年資21~25年.'+str(i*2+1),'服務年資25年以上.'+str(i*2+1)]+['管理職年資1~5年.'+str(i*2),'管理職年資6~10年.'+str(i*2),'管理職年資11~15年.'+str(i*2),'管理職年資16~20年.'+str(i*2),'管理職年資21~25年.'+str(i*2),'管理職年資25年以上.'+str(i*2)]+['管理職年資1~5年.'+str(i*2+1),'管理職年資6~10年.'+str(i*2+1),'管理職年資11~15年.'+str(i*2+1),'管理職年資16~20年.'+str(i*2+1),'管理職年資21~25年.'+str(i*2+1),'管理職年資25年以上.'+str(i*2+1)]+['35歲以下.'+str(i*2),'36-45歲.'+str(i*2),'46-55歲.'+str(i*2),'56-65歲.'+str(i*2),'66歲以上.'+str(i*2)]+['35歲以下.'+str(i*2+1),'36-45歲.'+str(i*2+1),'46-55歲.'+str(i*2+1),'56-65歲.'+str(i*2+1),'66歲以上.'+str(i*2+1)]+['35歲以下.'+str(i*2+18),'36-45歲.'+str(i*2+18),'46-55歲.'+str(i*2+18),'56-65歲.'+str(i*2+18),'66歲以上.'+str(i*2+18)]+['35歲以下.'+str(i*2+18+1),'36-45歲.'+str(i*2+18+1),'46-55歲.'+str(i*2+18+1),'56-65歲.'+str(i*2+18+1),'66歲以上.'+str(i*2+18+1)]+['35歲以下.'+str(i*2+36),'36-45歲.'+str(i*2+36),'46-55歲.'+str(i*2+36),'56-65歲.'+str(i*2+36),'66歲以上.'+str(i*2+36)]+['35歲以下.'+str(i*2+36+1),'36-45歲.'+str(i*2+36+1),'46-55歲.'+str(i*2+36+1),'56-65歲.'+str(i*2+36+1),'66歲以上.'+str(i*2+36+1)]+['35歲以下.'+str(i*2+54),'36-45歲.'+str(i*2+54),'46-55歲.'+str(i*2+54),'56-65歲.'+str(i*2+54),'66歲以上.'+str(i*2+54)]+['35歲以下.'+str(i*2+54+1),'36-45歲.'+str(i*2+54+1),'46-55歲.'+str(i*2+54+1),'56-65歲.'+str(i*2+54+1),'66歲以上.'+str(i*2+54+1)]+['35歲以下.'+str(i*2+72),'36-45歲.'+str(i*2+72),'46-55歲.'+str(i*2+72),'56-65歲.'+str(i*2+72),'66歲以上.'+str(i*2+72)]+['35歲以下.'+str(i*2+72+1),'36-45歲.'+str(i*2+72+1),'46-55歲.'+str(i*2+72+1),'56-65歲.'+str(i*2+72+1),'66歲以上.'+str(i*2+72+1)]+['35歲以下.'+str(i*2+90),'36-45歲.'+str(i*2+90),'46-55歲.'+str(i*2+90),'56-65歲.'+str(i*2+90),'66歲以上.'+str(i*2+90)]+['35歲以下.'+str(i*2+90+1),'36-45歲.'+str(i*2+90+1),'46-55歲.'+str(i*2+90+1),'56-65歲.'+str(i*2+90+1),'66歲以上.'+str(i*2+90+1)])

#放所有個別輸出資料
seniority = []
seniority_n = []
management_seniority = []
management_seniority_n = []
management = []
management_n = []
professional = []
professional_n = []



for i in range(9):

    inputList = themeList[i]

    df = statData[inputList]

    df = df.dropna(subset=[theme[i]],axis=0) #i

    col = inputList[3:]
    result = pd.DataFrame(columns=col)

    # for name in num:
    #     try:
    #         df1 = ct.group_table(df,'2.單位總員工人數',[name])
    #     except:
    #         continue
    #     df1 = pd.DataFrame(data=df1[col].sum(),columns=[name])
    #     result.at[name]=list(df1[name])

    for name in category:
        try:
            df1 = ct.group_table(df,'1.單位名稱',[name])
        except:
            continue
        df1 = pd.DataFrame(data=df1[col].sum(),columns=[name])
        result.at[name] = list(df1[name])

    result.at['產官研合計'] = list(result.sum())

    #result.index.name='單位總員工人數'

    #年資
    #總人數
    target = inputList[5:17]
    boy = result[target[:6]]
    boy.columns = ['服務年資1~5年','服務年資6~10年','服務年資11~15年','服務年資16~20年','服務年資21~25年','服務年資25年以上']
    girl = result[target[6:]]
    girl.columns = ['服務年資1~5年','服務年資6~10年','服務年資11~15年','服務年資16~20年','服務年資21~25年','服務年資25年以上']
    boy = boy.T
    girl = girl.T
    boy.at['合計'] = list(boy.sum())
    girl.at['合計'] = list(girl.sum())
    member = (boy+girl).astype(int)
    #女性佔比
    temp = (girl/member).fillna(0)
    girlRate = (girl/member)*100
    girlRate = girlRate.fillna(0)
    temp = (girl/member).fillna(0)
    girlRate = (girlRate.round(1).astype(str) + '%').replace('0.0%','-')
    col = list(member.columns) 
    index = list(member.index)
    result1 = pd.DataFrame(index=index)
    for item in col:
        result1[item+'0']=girlRate[item]
        result1[item+'1']=member[item]
        
    result1.columns = [ct.listDoubleMerge(col),['女性占比','總人數']*len(col)]
    col.insert(0,col.pop()) #尾換到頭
    result1 = result1[col]
    index.insert(0,index.pop())
    result1 = result1.reindex(index=index)
    result1.columns.names = [theme[i],'服務年資'] #i
    result1 = result1.rename(index={'合計':'各年資合計'})
    seniority.append(result1)

    #管理職年資
    target = inputList[17:29]
    boy = result[target[:6]]
    boy.columns = ['管理職年資1~5年','管理職年資6~10年','管理職年資11~15年','管理職年資16~20年','管理職年資21~25年','管理職年資25年以上']
    girl = result[target[6:]]
    girl.columns = ['管理職年資1~5年','管理職年資6~10年','管理職年資11~15年','管理職年資16~20年','管理職年資21~25年','管理職年資25年以上']
    boy = boy.T
    girl = girl.T
    boy.at['合計'] = list(boy.sum())
    girl.at['合計'] = list(girl.sum())
    member = (boy+girl).astype(int)
    #女性佔比
    girlRate = (girl/member)*100
    girlRate = girlRate.fillna(0)
    temp = (girl/member).fillna(0)
    girlRate = (girlRate.round(1).astype(str) + '%').replace('0.0%','-')
    col = list(member.columns) 
    index = list(member.index)
    result1 = pd.DataFrame(index=index)
    for item in col:
        result1[item+'0']=girlRate[item]
        result1[item+'1']=member[item]
        
    result1.columns = [ct.listDoubleMerge(col),['女性占比','總人數']*len(col)]
    col.insert(0,col.pop()) #尾換到頭
    result1 = result1[col]
    index.insert(0,index.pop())
    result1 = result1.reindex(index=index)
    result1.columns.names = [theme[i],'管理職服務年資'] #i
    result1 = result1.rename(index={'合計':'各年資合計'})
    management_seniority.append(result1)

    #各階管理職人員人數

    target = inputList[29:59]
    ageCol = ['35歲以下', '36-45歲', '46-55歲', '56-65歲', '66歲以上']
    level = ['初階管理職','中階管理職','高階管理職']
    result2 = pd.DataFrame()
    all_sum = []
    for k in range(3):

        data = result[target[k*10:k*10+10]].astype(int)
        col = list(data.columns)
        index = list(data.index)

        member = pd.DataFrame(index=index)
        for j in range(len(ageCol)):
            member[ageCol[j]]=list(data[col[j]]+data[col[j+5]])
        girl = data[col[5:]]
        girl.columns = ageCol
        member = member.T
        member.at['各年齡合計'] = list(member.sum())
        member = member.astype(int)
        if(member_sum):
            member_sum = list(member.loc['各年齡合計'])
        else:
            member_sum += list(member.loc['各年齡合計'])
        
        girl=girl.T
        girl.at['各年齡合計'] = list(girl.sum())
        if(girl_sum):
            girl_sum = list(girl.loc['各年齡合計'])
        else:
            girl_sum += list(girl.loc['各年齡合計'])
        
        girlRate = (girl/member)*100
        girlRate = girlRate.fillna(0)

        girlRate = (girlRate.round(1).astype(str) + '%').replace('0.0%','-')
        result1 = pd.DataFrame(index=list(girlRate.index))
        for item in index[:5]:
            result1[item+'0']=girlRate[item]
            result1[item+'1']=member[item]

        result1.columns = [ ct.listDoubleMerge(girlRate.columns) , ['女性占比','總人數']*len(list(girlRate.columns))]
        col = list(girlRate.columns)
        col.insert(0,col.pop())
        result1 = result1[col]

        index = list(girlRate.index)
        index.insert(0,index.pop())
        result1 = result1.reindex(index=index)
        
        

        result1.index = [ [level[k]]*len(index) , index]
        if(result2.empty):
            result2 = result1.copy()
        else:
            result2 = pd.concat([result2,result1],axis=0)

        if(not all_sum):
            all_sum = girl_sum + member_sum
        else:
            tempAllSum = girl_sum + member_sum
            for n in range(len(all_sum)):
                all_sum[n]+=tempAllSum[n] 
        
#all_sum.append(all_sum.pop(0))
#all_sum.append(all_sum.pop(0))
    newAllSum = []
    for n in range(len(all_sum)//2):
        girlRate = (all_sum[n]/all_sum[n+len(all_sum)//2])*100
        newAllSum.append( (str(round(girlRate,1)) + '%').replace('0.0%','-') )
        newAllSum.append(all_sum[n+len(all_sum)//2])
    newAllSum.insert(0,newAllSum.pop())
    newAllSum.insert(0,newAllSum.pop())
    result2.ix[('不分管理職','合計'),:]=newAllSum

    #all_sum.index = ('各年齡合計','各年齡合計') 
    #result2.append(all_sum)
    management.append(result2)

    #專業職人員

    target = inputList[59:89]
    ageCol = ['35歲以下', '36-45歲', '46-55歲', '56-65歲', '66歲以上']
    level = ['初階專業職','中階專業職','高階專業職']
    result2 = pd.DataFrame()
    all_sum = []
    for k in range(3):

        data = result[target[k*10:k*10+10]].astype(int)
        col = list(data.columns)
        index = list(data.index)

        member = pd.DataFrame(index=index)
        for j in range(len(ageCol)):
            member[ageCol[j]]=list(data[col[j]]+data[col[j+5]])
        girl = data[col[5:]]
        girl.columns = ageCol
        member = member.T
        member.at['各年齡合計'] = list(member.sum())
        member = member.astype(int)
        if(member_sum):
            member_sum = list(member.loc['各年齡合計'])
        else:
            member_sum += list(member.loc['各年齡合計'])
        
        girl=girl.T
        girl.at['各年齡合計'] = list(girl.sum())
        if(girl_sum):
            girl_sum = list(girl.loc['各年齡合計'])
        else:
            girl_sum += list(girl.loc['各年齡合計'])
        
        girlRate = (girl/member)*100
        girlRate = girlRate.fillna(0)

        girlRate = (girlRate.round(1).astype(str) + '%').replace('0.0%','-')
        result1 = pd.DataFrame(index=list(girlRate.index))
        for item in index[:5]:
            result1[item+'0']=girlRate[item]
            result1[item+'1']=member[item]

        result1.columns = [ ct.listDoubleMerge(girlRate.columns) , ['女性占比','總人數']*len(list(girlRate.columns))]
        col = list(girlRate.columns)
        col.insert(0,col.pop())
        result1 = result1[col]

        index = list(girlRate.index)
        index.insert(0,index.pop())
        result1 = result1.reindex(index=index)
        
        result1.index = [ [level[k]]*len(index) , index]
        if(result2.empty):
            result2 = result1.copy()
        else:
            result2 = pd.concat([result2,result1],axis=0)

        if(not all_sum):
            all_sum = girl_sum + member_sum
        else:
            tempAllSum = girl_sum + member_sum
            for n in range(len(all_sum)):
                all_sum[n]+=tempAllSum[n] 
        
#all_sum.append(all_sum.pop(0))
#all_sum.append(all_sum.pop(0))
    newAllSum = []
    for n in range(len(all_sum)//2):
        girlRate = (all_sum[n]/all_sum[n+len(all_sum)//2])*100
        newAllSum.append( (str(round(girlRate,1)) + '%').replace('0.0%','-') )
        newAllSum.append(all_sum[n+len(all_sum)//2])
    newAllSum.insert(0,newAllSum.pop())
    newAllSum.insert(0,newAllSum.pop())
    result2.ix[('不分專業職','合計'),:]=newAllSum
    professional.append(result2)    



#第三部分：性別相關福利措施，以及轉入工程與科技領域之案例

col = ['106年 請假人次','107年 請假人次','108年 請假人次','106年 請假人次.1','107年 請假人次.1','108年 請假人次.1','106年 請假人次.2','107年 請假人次.2','108年 請假人次.2','106年 請假人次.3','107年 請假人次.3','108年 請假人次.3']
inputList = statData[['1.單位名稱']+col+['哺集乳室','女性生理假（不扣薪）','Third Choice托嬰服務（指設有收托二歲以下兒童之服務機構）','托兒服務（指設有收托二歲至六歲兒童之服務機構）','育兒津貼','（因照顧家庭因素可申請）彈性工時']]
yearCol = ['106年','107年','108年']
title = ['過去三年間機構獲准之育嬰留職停薪人次','去三年間機構獲准之家庭照謢假人次']

result = pd.DataFrame(columns=col)
for name in category:
    try:
        df1 = ct.group_table(inputList,'1.單位名稱',[name])
    except:
        continue
    result.at[name] = list(df1[col].sum())

result.at['產官研小計'] = list(result.sum())
resultList = []
for i in range(2):
    boy = result[col[i*len(yearCol):i*len(yearCol)+len(yearCol)]]
    boy.columns = yearCol
    boy = boy.T
    boy.at['各年小計'] = list(boy.sum())
    girl = result[col[i*len(yearCol)+len(yearCol):i*len(yearCol)+len(yearCol)*2]]
    girl.columns = yearCol
    girl = girl.T
    girl.at['各年小計'] = list(girl.sum())
    member = (boy+girl).astype(int)
    girlRate = (girl/member)*100
    #girlRate = girlRate.T
    girlRate = girlRate.fillna(0)
    girlRate = (girlRate.round(1).astype(str) + '%').replace('0.0%','-')
    column = list(member.columns) 
    index = list(member.index)
    result1 = pd.DataFrame(index=index)
    for item in column:
        result1[item+'0']=girlRate[item]
        result1[item+'1']=member[item]
    result1.columns = [ct.listDoubleMerge(column),['女性占比','總人數']*len(column)]
    column.insert(0,column.pop()) #尾換到頭
    result1 = result1[column]
    index.insert(0,index.pop())
    result1 = result1.reindex(index=index)
    result1.columns.names=[title[i],'']
    resultList.append(result1)

#---------------#

col = ['哺集乳室','女性生理假（不扣薪）','Third Choice托嬰服務（指設有收托二歲以下兒童之服務機構）','托兒服務（指設有收托二歲至六歲兒童之服務機構）','育兒津貼','（因照顧家庭因素可申請）彈性工時']

result = statData[['1.單位名稱','哺集乳室','女性生理假（不扣薪）','Third Choice托嬰服務（指設有收托二歲以下兒童之服務機構）','托兒服務（指設有收托二歲至六歲兒童之服務機構）','育兒津貼','（因照顧家庭因素可申請）彈性工時']]

answer = pd.DataFrame(columns=col)

for name in category:
    answerList = []
    try:
        result1 = ct.group_table(result,'1.單位名稱',[name])
    except:
        continue
    result1 = result1.fillna(0).drop('1.單位名稱',axis=1).reset_index(drop=True)
 
    for item in col:
        count=0

        for i in range(len(result1[item])):
            if (result1[item][i] is not 0):
                count+=1
            
        answerList.append(count)

    answer.at[name] = answerList

answer.at['產官研合計'] = list(answer.sum())
answer = answer.T
column = list(answer.columns)
column.insert(0,column.pop()) #尾換到頭
answer = (answer[column]).astype(int)

with pd.ExcelWriter('單位版分析.xlsx') as writer:
    reportData.to_excel(writer,sheet_name='回報參數',encoding='utf_8_sig')
    lost_company.to_excel(writer,sheet_name='其他歸類單位',encoding='utf_8_sig')
    member_analysis.to_excel(writer,sheet_name='總人數分析',encoding='utf_8_sig')
    for i in range(9):
        seniority[i].to_excel(writer,sheet_name=str(i+1)+'.'+theme[i]+'服務年資',encoding='utf_8_sig')
        management_seniority[i].to_excel(writer,sheet_name=str(i+1)+'.'+theme[i]+'管理職服務年資',encoding='utf_8_sig')
        management[i].to_excel(writer,sheet_name=str(i+1)+'.'+theme[i]+'各階管理職人數',encoding='utf_8_sig')
        professional[i].to_excel(writer,sheet_name=str(i+1)+'.'+theme[i]+'各階專業職人數',encoding='utf_8_sig')
    resultList[0].to_excel(writer,sheet_name=title[0],encoding='utf_8_sig')
    resultList[1].to_excel(writer,sheet_name=title[1],encoding='utf_8_sig')
    answer.to_excel(writer,sheet_name='各機構已實施之福利措施',encoding='utf_8_sig')