import pandas as pd
import os
import re
import numpy as np

#os.chdir("/Users/cilab/PartTime_PythonAnalysis/group")
os.chdir(r"C:\Users\Naichen\Documents\GitHub\stu00608.github.io\PartTime_PythonAnalysis\group")
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
for i in range(len(statData['1.單位名稱'])):
    if(   str(statData['1.單位名稱'][i]).find('test') != -1 ):
        delList.append(i)
    elif( str(statData['1.單位名稱'][i]).find('111111111111') != -1 ):
        delList.append(i)
    elif( str(statData['1.單位名稱'][i]).find('測試') != -1 ):
        delList.append(i)

statData = statData.drop(delList).reset_index(drop=True)


#os.chdir("/Users/cilab/PartTime_PythonAnalysis/group/outputs")
os.chdir(r"C:\Users\Naichen\Documents\GitHub\stu00608.github.io\PartTime_PythonAnalysis\group\outputs")
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

category = ['研小計','官小計','產小計','其他']

i=0
for j in range(len(statData['1.單位名稱'])):
    if( classification[0].count(statData['1.單位名稱'][j])>=1 ):
        statData.at[j,'1.單位名稱']='產小計'
    elif(classification[1].count(statData['1.單位名稱'][j])>=1):
        statData.at[j,'1.單位名稱']='研小計'
    elif(classification[2].count(statData['1.單位名稱'][j])>=1):
        statData.at[j,'1.單位名稱']='官小計'
    else:
        lost_company.at[i] = statData['1.單位名稱'][j]
        i+=1
        statData.at[j,'1.單位名稱']='其他'
        


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
member.at['合計'] = list(member.sum())
member = member.astype(int)
girlRate.at['合計'] = list(girlRate.sum())
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
#result1.to_csv('總人數分析.csv',encoding='utf_8_sig')

#------------------------#

i=0
themeList = [ ['1.單位名稱','2.單位總員工人數',theme[i],theme[i]+'領域專長 男性 專任 人數',theme[i]+'領域專長 女性 專任 人數']+['服務年資1~5年','服務年資6~10年','服務年資11~15年','服務年資16~20年','服務年資21~25年','服務年資25年以上']+['服務年資1~5年.'+str(i*2+1),'服務年資6~10年.'+str(i*2+1),'服務年資11~15年.'+str(i*2+1),'服務年資16~20年.'+str(i*2+1),'服務年資21~25年.'+str(i*2+1),'服務年資25年以上.'+str(i*2+1)]+['管理職年資1~5年','管理職年資6~10年','管理職年資11~15年','管理職年資16~20年','管理職年資21~25年','管理職年資25年以上']+['管理職年資1~5年.'+str(i*2+1),'管理職年資6~10年.'+str(i*2+1),'管理職年資11~15年.'+str(i*2+1),'管理職年資16~20年.'+str(i*2+1),'管理職年資21~25年.'+str(i*2+1),'管理職年資25年以上.'+str(i*2+1)]+['35歲以下','36-45歲','46-55歲','56-65歲','66歲以上']+['35歲以下.'+str(i*2+1),'36-45歲.'+str(i*2+1),'46-55歲.'+str(i*2+1),'56-65歲.'+str(i*2+1),'66歲以上.'+str(i*2+1)]+['35歲以下.'+str(i*2+18),'36-45歲.'+str(i*2+18),'46-55歲.'+str(i*2+18),'56-65歲.'+str(i*2+18),'66歲以上.'+str(i*2+18)]+['35歲以下.'+str(i*2+18+1),'36-45歲.'+str(i*2+18+1),'46-55歲.'+str(i*2+18+1),'56-65歲.'+str(i*2+18+1),'66歲以上.'+str(i*2+18+1)]+['35歲以下.'+str(i*2+36),'36-45歲.'+str(i*2+36),'46-55歲.'+str(i*2+36),'56-65歲.'+str(i*2+36),'66歲以上.'+str(i*2+36)]+['35歲以下.'+str(i*2+36+1),'36-45歲.'+str(i*2+36+1),'46-55歲.'+str(i*2+36+1),'56-65歲.'+str(i*2+36+1),'66歲以上.'+str(i*2+36+1)]+['35歲以下.'+str(i*2+54),'36-45歲.'+str(i*2+54),'46-55歲.'+str(i*2+54),'56-65歲.'+str(i*2+54),'66歲以上.'+str(i*2+54)]+['35歲以下.'+str(i*2+54+1),'36-45歲.'+str(i*2+54+1),'46-55歲.'+str(i*2+54+1),'56-65歲.'+str(i*2+54+1),'66歲以上.'+str(i*2+54+1)]+['35歲以下.'+str(i*2+72),'36-45歲.'+str(i*2+72),'46-55歲.'+str(i*2+72),'56-65歲.'+str(i*2+72),'66歲以上.'+str(i*2+72)]+['35歲以下.'+str(i*2+72+1),'36-45歲.'+str(i*2+72+1),'46-55歲.'+str(i*2+72+1),'56-65歲.'+str(i*2+72+1),'66歲以上.'+str(i*2+72+1)]+['35歲以下.'+str(i*2+90),'36-45歲.'+str(i*2+90),'46-55歲.'+str(i*2+90),'56-65歲.'+str(i*2+90),'66歲以上.'+str(i*2+90)]+['35歲以下.'+str(i*2+90+1),'36-45歲.'+str(i*2+90+1),'46-55歲.'+str(i*2+90+1),'56-65歲.'+str(i*2+90+1),'66歲以上.'+str(i*2+90+1)] ] 

for i in range(1,len(theme)):
    themeList.append(['1.單位名稱','2.單位總員工人數',theme[i],theme[i]+'領域專長 男性 專任 人數',theme[i]+'領域專長 女性 專任 人數']+['服務年資1~5年.'+str(i*2),'服務年資6~10年.'+str(i*2),'服務年資11~15年.'+str(i*2),'服務年資16~20年.'+str(i*2),'服務年資21~25年.'+str(i*2),'服務年資25年以上.'+str(i*2)]+['服務年資1~5年.'+str(i*2+1),'服務年資6~10年.'+str(i*2+1),'服務年資11~15年.'+str(i*2+1),'服務年資16~20年.'+str(i*2+1),'服務年資21~25年.'+str(i*2+1),'服務年資25年以上.'+str(i*2+1)]+['管理職年資1~5年.'+str(i*2),'管理職年資6~10年.'+str(i*2),'管理職年資11~15年.'+str(i*2),'管理職年資16~20年.'+str(i*2),'管理職年資21~25年.'+str(i*2),'管理職年資25年以上.'+str(i*2)]+['管理職年資1~5年.'+str(i*2+1),'管理職年資6~10年.'+str(i*2+1),'管理職年資11~15年.'+str(i*2+1),'管理職年資16~20年.'+str(i*2+1),'管理職年資21~25年.'+str(i*2+1),'管理職年資25年以上.'+str(i*2+1)]+['35歲以下.'+str(i*2),'36-45歲.'+str(i*2),'46-55歲.'+str(i*2),'56-65歲.'+str(i*2),'66歲以上.'+str(i*2)]+['35歲以下.'+str(i*2+1),'36-45歲.'+str(i*2+1),'46-55歲.'+str(i*2+1),'56-65歲.'+str(i*2+1),'66歲以上.'+str(i*2+1)]+['35歲以下.'+str(i*2+18),'36-45歲.'+str(i*2+18),'46-55歲.'+str(i*2+18),'56-65歲.'+str(i*2+18),'66歲以上.'+str(i*2+18)]+['35歲以下.'+str(i*2+18+1),'36-45歲.'+str(i*2+18+1),'46-55歲.'+str(i*2+18+1),'56-65歲.'+str(i*2+18+1),'66歲以上.'+str(i*2+18+1)]+['35歲以下.'+str(i*2+36),'36-45歲.'+str(i*2+36),'46-55歲.'+str(i*2+36),'56-65歲.'+str(i*2+36),'66歲以上.'+str(i*2+36)]+['35歲以下.'+str(i*2+36+1),'36-45歲.'+str(i*2+36+1),'46-55歲.'+str(i*2+36+1),'56-65歲.'+str(i*2+36+1),'66歲以上.'+str(i*2+36+1)]+['35歲以下.'+str(i*2+54),'36-45歲.'+str(i*2+54),'46-55歲.'+str(i*2+54),'56-65歲.'+str(i*2+54),'66歲以上.'+str(i*2+54)]+['35歲以下.'+str(i*2+54+1),'36-45歲.'+str(i*2+54+1),'46-55歲.'+str(i*2+54+1),'56-65歲.'+str(i*2+54+1),'66歲以上.'+str(i*2+54+1)]+['35歲以下.'+str(i*2+72),'36-45歲.'+str(i*2+72),'46-55歲.'+str(i*2+72),'56-65歲.'+str(i*2+72),'66歲以上.'+str(i*2+72)]+['35歲以下.'+str(i*2+72+1),'36-45歲.'+str(i*2+72+1),'46-55歲.'+str(i*2+72+1),'56-65歲.'+str(i*2+72+1),'66歲以上.'+str(i*2+72+1)]+['35歲以下.'+str(i*2+90),'36-45歲.'+str(i*2+90),'46-55歲.'+str(i*2+90),'56-65歲.'+str(i*2+90),'66歲以上.'+str(i*2+90)]+['35歲以下.'+str(i*2+90+1),'36-45歲.'+str(i*2+90+1),'46-55歲.'+str(i*2+90+1),'56-65歲.'+str(i*2+90+1),'66歲以上.'+str(i*2+90+1)])

'''

themeList = [ ['1.單位名稱','2.單位總員工人數', #0
                '土木營建', #1
                '土木營建領域專長 男性 專任 人數','土木營建領域專長 女性 專任 人數', #2
                '服務年資1~5年','服務年資6~10年','服務年資11~15年','服務年資16~20年','服務年資21~25年','服務年資25年以上','服務年資1~5年.1','服務年資6~10年.1','服務年資11~15年.1','服務年資16~20年.2','服務年資21~25年.1','服務年資25年以上.1', #4
                '管理職年資1~5年','管理職年資6~10年','管理職年資11~15年','管理職年資16~20年','管理職年資21~25年','管理職年資25年以上','管理職年資1~5年.1','管理職年資6~10年.1','管理職年資11~15年.1','管理職年資16~20年.1','管理職年資21~25年.1','管理職年資25年以上.1', #16
                '35歲以下','36-45歲','46-55歲','56-65歲','66歲以上','35歲以下.1','36-45歲.1','46-55歲.1','56-65歲.1','66歲以上.1', #28
                '35歲以下.18','36-45歲.18','46-55歲.17','56-65歲.18','66歲以上.18','35歲以下.19','36-45歲.19','46-55歲.18','56-65歲.19','66歲以上.19', #38
                '35歲以下.36','36-45歲.36','46-55歲.35','56-65歲.36','66歲以上.36','35歲以下.37','36-45歲.37','46-55歲.36','56-65歲.37','66歲以上.37'] ,#48
              ['1.單位名稱','2.單位總員工人數',
                '建築、都市規劃',
                '建築、都市規劃領域專長 男性 專任 人數','建築、都市規劃領域專長 女性 專任 人數',
                '服務年資1~5年.2','服務年資6~10年.2','服務年資11~15年.2','服務年資16~20年.3','服務年資21~25年.2','服務年資25年以上.2','服務年資1~5年.3','服務年資6~10年.3','服務年資11~15年.3','服務年資16~20年.4','服務年資21~25年.3','服務年資25年以上.3',
                '管理職年資1~5年.2','管理職年資6~10年.2','管理職年資11~15年.2','管理職年資16~20年.2','管理職年資21~25年.2','管理職年資25年以上.2','管理職年資1~5年.3','管理職年資6~10年.3','管理職年資11~15年.3','管理職年資16~20年.3','管理職年資21~25年.3','管理職年資25年以上.3',
                '35歲以下.2','36-45歲.2','46-55歲.2','56-65歲.2','66歲以上.2','35歲以下.3','36-45歲.3','46-55歲.3','56-65歲.3','66歲以上.3',
                '35歲以下.20','36-45歲.20','46-55歲.19','56-65歲.20','66歲以上.20','35歲以下.21','36-45歲.21','46-55歲.20','56-65歲.21','66歲以上.21',
                '35歲以下.38','36-45歲.38','46-55歲.37','56-65歲.38','66歲以上.38','35歲以下.39','36-45歲.39','46-55歲.38','56-65歲.39','66歲以上.39'],
              ['1.單位名稱','2.單位總員工人數',
                '電子電機',
                '電子電機領域專長 男性 專任 人數','電子電機領域專長 女性 專任 人數',
                '服務年資1~5年.4','服務年資6~10年.4','服務年資11~15年.4','服務年資16~20年.5','服務年資21~25年.4','服務年資25年以上.4','服務年資1~5年.5','服務年資6~10年.5','服務年資11~15年.5','服務年資16~20年.6','服務年資21~25年.5','服務年資25年以上.5',
                '管理職年資1~5年.4','管理職年資6~10年.4','管理職年資11~15年.4','管理職年資16~20年.4','管理職年資21~25年.4','管理職年資25年以上.4','管理職年資1~5年.5','管理職年資6~10年.5','管理職年資11~15年.5','管理職年資16~20年.5','管理職年資21~25年.5','管理職年資25年以上.5',
                '35歲以下.4','36-45歲.4','46-55歲.4','56-65歲.4','66歲以上.4','35歲以下.5','36-45歲.5','46-55歲.5','56-65歲.5','66歲以上.5',
                '35歲以下.22','36-45歲.22','46-55歲.21','56-65歲.22','66歲以上.22','35歲以下.23','36-45歲.23','46-55歲.22','56-65歲.23','66歲以上.23',
                '35歲以下.40','36-45歲.40','45-55歲','56-65歲.40','66歲以上.40','35歲以下.41','36-45歲.41','46-55歲.39','56-65歲.41','66歲以上.41'],
              ['1.單位名稱','2.單位總員工人數',
                '資訊通訊',
                '資訊通訊領域專長 男性 專任 人數','資訊通訊領域專長 女性 專任 人數',
                '服務年資1~5年.6','服務年資6~10年.6','服務年資11~15年.6','服務年資16~20年.7','服務年資21~25年.6','服務年資25年以上.6','服務年資1~5年.7','服務年資6~10年.7','服務年資11~15年.7','服務年資16~20年.8','服務年資21~25年.7','服務年資25年以上.7',
                '管理職年資1~5年.6','管理職年資6~10年.6','管理職年資11~15年.6','管理職年資16~20年.6','管理職年資21~25年.6','管理職年資25年以上.6','管理職年資1~5年.7','管理職年資6~10年.7','管理職年資11~15年.7','管理職年資16~20年.7','管理職年資21~25年.7','管理職年資25年以上.7',
                '35歲以下.6','36-45歲.6','46-55歲.6','56-65歲.6','66歲以上.6','35歲以下.7','36-45歲.7','46-55歲.7','56-65歲.7','66歲以上.7',
                '35歲以下.24','36-45歲.24','46-55歲.23','56-65歲.24','66歲以上.24','35歲以下.25','36-45歲.25','46-55歲.24','56-65歲.25','66歲以上.25',
                '35歲以下.42','36-45歲.42','46-55歲.40','56-65歲.42','66歲以上.42','35歲以下.43','36-45歲.43','46-55歲.41','56-65歲.43','66歲以上.43'],

              ['1.單位名稱','2.單位總員工人數',
                '化工材料',
                '化工材料領域專長 男性 專任 人數','化工材料領域專長 女性 專任 人數',
                '服務年資1~5年.8','服務年資6~10年.8','服務年資11~15年.8','服務年資16~20年.10','服務年資21~25年.8','服務年資25年以上.8','服務年資1~5年.9','服務年資6~10年.9','服務年資11~15年.9','服務年資16~20年.11','服務年資21~25年.9','服務年資25年以上.9',
                '管理職年資1~5年.8','管理職年資6~10年.8','管理職年資11~15年.8','管理職年資16~20年.8','管理職年資21~25年.8','管理職年資25年以上.8','管理職年資1~5年.9','管理職年資6~10年.9','管理職年資11~15年.9','管理職年資16~20年.9','管理職年資21~25年.9','管理職年資25年以上.9',
                '35歲以下.8','36-45歲.8','46-55歲.8','56-65歲.8','66歲以上.8','35歲以下.9','36-45歲.9','46-55歲.9','56-65歲.9','66歲以上.9',
                '35歲以下.26','36-45歲.26','46-55歲.25','56-65歲.26','66歲以上.26','35歲以下.27','36-45歲.27','46-55歲.26','56-65歲.27','66歲以上.27',
                '35歲以下.44','36-45歲.44','46-55歲.42','56-65歲.44','66歲以上.44','35歲以下.45','36-45歲.45','46-55歲.43','56-65歲.45','66歲以上.45'],
            
              ['1.單位名稱','2.單位總員工人數',
                '生技醫工',
                '生技醫工領域專長 男性 專任 人數','生技醫工領域專長 女性 專任 人數',
                '服務年資1~5年.10','服務年資6~10年.10','服務年資11~15年.10','服務年資16~20年.12','服務年資21~25年.10','服務年資25年以上.10','服務年資1~5年.11','服務年資6~10年.11','服務年資11~15年.11','服務年資16~20年.13','服務年資21~25年.11','服務年資25年以上.11',
                '管理職年資1~5年.11','管理職年資6~10年.10','管理職年資11~15年.10','管理職年資16~20年.10','管理職年資21~25年.10','管理職年資25年以上.10','管理職年資1~5年.12','管理職年資6~10年.11','管理職年資11~15年.11','管理職年資16~20年.11','管理職年資21~25年.11','管理職年資25年以上.11',
                '35歲以下.10','36-45歲.10','46-55歲.10','56-65歲.10','66歲以上.10','35歲以下.11','36-45歲.11','46-55歲.11','56-65歲.11','66歲以上.11',
                '35歲以下.28','36-45歲.28','46-55歲.27','56-65歲.28','66歲以上.28','35歲以下.29','36-45歲.29','46-55歲.28','56-65歲.29','66歲以上.29',
                '35歲以下.46','36-45歲.46','46-55歲.44','56-65歲.46','66歲以上.46','35歲以下.47','36-45歲.47','46-55歲.45','56-65歲.47','66歲以上.47'],
            
              ['1.單位名稱','2.單位總員工人數',
                '環工綠能',
                '環工綠能領域專長 男性 專任 人數','環工綠能領域專長 女性 專任 人數',
                '服務年資1~5年.12','服務年資6~10年.12','服務年資11~15年.12','服務年資16~20年.14','服務年資21~25年.12','服務年資25年以上.12','服務年資1~5年.13','服務年資6~10年.13','服務年資11~15年.13','服務年資16~20年.15','服務年資21~25年.13','服務年資25年以上.13',
                '管理職年資1~5年.13','管理職年資6~10年.12','管理職年資11~15年.12','管理職年資16~20年.12','管理職年資21~25年.12','管理職年資25年以上.12','管理職年資1~5年.14','管理職年資6~10年.13','管理職年資11~15年.13','管理職年資16~20年.13','管理職年資21~25年.13','管理職年資25年以上.13',
                '35歲以下.12','36-45歲.12','46-55歲.12','56-65歲.12','66歲以上.12','35歲以下.13','36-45歲.13','46-55歲.13','56-65歲.13','66歲以上.13',
                '35歲以下.30','36-45歲.30','46-55歲.29','56-65歲.30','66歲以上.30','35歲以下.31','36-45歲.31','46-55歲.30','56-65歲.31','66歲以上.31',
                '35歲以下.48','36-45歲.48','46-55歲.46','56-65歲.48','66歲以上.48','35歲以下.49','36-45歲.49','46-55歲.47','56-65歲.49','66歲以上.49'],

              ['1.單位名稱','2.單位總員工人數',
                '機械',
                '機械領域專長 男性 專任 人數','機械領域專長 女性 專任 人數',
                '服務年資1~5年.14','服務年資6~10年.14','服務年資11~15年.14','服務年資16~20年.16','服務年資21~25年.14','服務年資25年以上.14','服務年資1~5年.15','服務年資6~10年.15','服務年資11~15年.15','服務年資16~20年.17','服務年資21~25年.15','服務年資21~25年.16','服務年資25年以上.15',
                '管理職年資1~5年.15','管理職年資6~10年.14','管理職年資11~15年.14','管理職年資16~20年.14','管理職年資21~25年.14','管理職年資25年以上.14','管理職年資1~5年.16','管理職年資6~10年.15','管理職年資11~15年.15','管理職年資16~20年.15','管理職年資21~25年.15','管理職年資25年以上.15',
                '35歲以下.14','36-45歲.14','46-55歲.14','56-65歲.14','66歲以上.14','35歲以下.15','36-45歲.15','46-55歲.15','56-65歲.15','66歲以上.15',
                '35歲以下.32','36-45歲.32','46-55歲.31','56-65歲.32','66歲以上.32','35歲以下.33','36-45歲.33','46-55歲.32','56-65歲.33','66歲以上.33',
                '35歲以下.50','36-45歲.50','46-55歲.48','56-65歲.50','66歲以上.50','35歲以下.51','36-45歲.51','46-55歲.49','56-65歲.51','66歲以上.51',],

              ['1.單位名稱','2.單位總員工人數',
                '其他',
                '其他領域專長 男性 專任 人數','其他領域專長 女性 專任 人數',
                '服務年資1~5年.16','服務年資6~10年.16','服務年資11~15年.16','服務年資16~20年.18','服務年資21~25年.17','服務年資25年以上.16','服務年資1~5年.17','服務年資6~10年.17','服務年資11~15年.17','服務年資16~20年.19','服務年資21~25年.18','服務年資25年以上.17',
                '管理職年資1~5年.17','管理職年資6~10年.16','管理職年資11~15年.16','管理職年資16~20年.16','管理職年資21~25年.16','管理職年資25年以上.16','管理職年資1~5年.18','管理職年資6~10年.17','管理職年資11~15年.17','管理職年資16~20年.17','管理職年資21~25年.17','管理職年資25年以上.17',
                '35歲以下.16','36-45歲.16','46-55歲.16','56-65歲.16','66歲以上.16','35歲以下.17','36-45歲.17','46-55歲.re','56-65歲.17','66歲以上.17',
                '35歲以下.34','36-45歲.34','46-55歲.33','56-65歲.34','66歲以上.34','35歲以下.35','36-45歲.35','46-55歲.34','56-65歲.35','66歲以上.35',
                '35歲以下.52','36-45歲.52','46-55歲.50','56-65歲.52','66歲以上.52','35歲以下.53','36-45歲.53','46-55歲.51','56-65歲.53','66歲以上.53',]
            
            ]

'''

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

    result.at['合計'] = list(result.sum())

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
    management_seniority.append(result1)

    #各階管理職人員人數

    target = inputList[29:59]
    ageCol = ['35歲以下', '36-45歲', '46-55歲', '56-65歲', '66歲以上']
    level = ['初階管理職','中階管理職','高階管理職']
    result2 = pd.DataFrame()
    for i in range(3):

        data = result[target[i*10:i*10+10]].astype(int)
        col = list(data.columns)
        index = list(data.index)

        member = pd.DataFrame(index=index)
        for j in range(len(ageCol)):
            member[ageCol[j]]=list(data[col[j]]+data[col[j+5]])
        girl = data[col[5:]]
        girl.columns = ageCol
        member = member.T
        member.at['合計'] = list(member.sum())
        member = member.astype(int)
        girl=girl.T
        girl.at['合計'] = list(girl.sum())
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
        result1.index = [ [level[i]]*len(index) , index]
        if(result2.empty):
            result2 = result1.copy()
        else:
            result2 = pd.concat([result2,result1],axis=0)

    management.append(result2)

    #專業職人員

    target = inputList[29:59]
    ageCol = ['35歲以下', '36-45歲', '46-55歲', '56-65歲', '66歲以上']
    level = ['初階專業職','中階專業職','高階專業職']
    result2 = pd.DataFrame()
    for i in range(3):

        data = result[target[i*10:i*10+10]].astype(int)
        col = list(data.columns)
        index = list(data.index)

        member = pd.DataFrame(index=index)
        for j in range(len(ageCol)):
            member[ageCol[j]]=list(data[col[j]]+data[col[j+5]])
        girl = data[col[5:]]
        girl.columns = ageCol
        member = member.T
        member.at['合計'] = list(member.sum())
        member = member.astype(int)
        girl=girl.T
        girl.at['合計'] = list(girl.sum())
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
        result1.index = [ [level[i]]*len(index) , index]
        if(result2.empty):
            result2 = result1.copy()
        else:
            result2 = pd.concat([result2,result1],axis=0)

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
    girlRate = girlRate.T
    girlRate = (girl/member)*100
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

result = inputList[['1.單位名稱','哺集乳室','女性生理假（不扣薪）','Third Choice托嬰服務（指設有收托二歲以下兒童之服務機構）','托兒服務（指設有收托二歲至六歲兒童之服務機構）','育兒津貼','（因照顧家庭因素可申請）彈性工時']]

answer = pd.DataFrame(columns=col)
answerList = []
for name in category:
    
    try:
        result1 = ct.group_table(result,'1.單位名稱',[name])
    except:
        continue
    result1 = result1.fillna(0).drop('1.單位名稱',axis=1).reset_index(drop=True)
    for item in col:
        count=0
        for i in range(len(result[item])):
            if (result[item][i] is not 0):
                count+=1
            
        answerList.append(count)
    
    answer.at[name] = answerList


with pd.ExcelWriter('單位版分析.xlsx') as writer:
    reportData.to_excel(writer,sheet_name='回報參數',encoding='utf_8_sig')
    lost_company.to_excel(writer,sheet_name='其他歸類單位',encoding='utf_8_sig')
    member_analysis.to_excel(writer,sheet_name='總人數分析',encoding='utf_8_sig')
    for i in range(9):
        seniority[i].to_excel(writer,sheet_name=str(i+1)+'.'+theme[i]+'服務年資',encoding='utf_8_sig')
        management_seniority[i].to_excel(writer,sheet_name=str(i+1)+'.'+theme[i]+'管理職服務年資',encoding='utf_8_sig')
        management[i].to_excel(writer,sheet_name=str(i+1)+'.'+theme[i]+'各階管理職人數',encoding='utf_8_sig')
        professional[i].to_excel(writer,sheet_name=str(i+1)+'.'+theme[i]+'各階專業職人數',encoding='utf_8_sig')
