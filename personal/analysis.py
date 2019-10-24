import pandas as pd
import os

os.chdir("/Users/cilab/PartTime_PythonAnalysis/personal")
os.getcwd()

import ct_tool as ct

statData = pd.read_csv("001.csv")
statData = statData.dropna(axis=1,how='all')
os.chdir("/Users/cilab/PartTime_PythonAnalysis/personal/outputs")
os.getcwd()

for i in range(len(statData['1. 性別：'])):
    if(statData['1. 性別：'][i]=='其他'):
        break
statData = statData.drop(i).reset_index(drop=True)

for i in range(len(statData['3. 最高學歷'])):
    if(statData['3. 最高學歷'][i]=='其他（請說明）'):
        break
statData = statData.drop(i).reset_index(drop=True)



'''
['Submitted', '1. 性別：', '2. 年齡', '3. 最高學歷', '4.\t最高學歷畢業校系',
       '5. 最高學歷畢業時間：請填寫西元年', '(a) 工程與科技領域職務(年)', '(b) 非工程與科技領域職務',
       '7. 是否正在從事工程與科技領域職務（含管理與學術研究）？', '8. 您離開工程職務之最主要原因為何？請勾選最主要原因，至多二項。',
       '原因 - 其他（請說明）', '9.\t請問您目前的工作狀況為何？', '10.\t請問您目前服務的職務較接近下列何者？',
       '請問目前受您管理的人員約＿＿人？', '11.\t您對於現任職務之五年後職涯發展的預期為何？請勾選最有可能的一項。', '請說明原因:',
       '12.\t請問您所預期的職涯發展，需要哪些配合因素來達成？請勾選最主要因素，至多二項。', '其他，請說明',
       '13.\t哪些福利措施最有助於您留在工程與科技領域就業？（不論現在是否有需求皆可選擇，可複選，最多三項）',
       '14.\t您服務的單位或待業前服務的單位提供哪些職場相關措施？（可複選）',
       '15.\t您認為工程與科技領域最需要改善的性別議題有哪些？請勾選其中最重要的議題，至多三項。',
       '16.\t在您或您認識的女性身上，性別在工程與科技領域的求學階段有否差異？', '較用功或成績較好', '較擅長數理學科',
       '動手實驗或實習的機會較少', '遇到挫折時較易放棄', '較不敢接受挑戰', '較容易情緒起伏', '較受人際問題干擾',
       '較容易為情所困', '還有其他差異?（請說明）', '17.\t在您或您認識的工程與科技領域女性身上，性別在求職時的職務選擇方面有否差異？',
       '選擇內勤為主的職務', '選擇無須出差的職務', '選擇無須應酬的職務', '選擇無須輪班或值夜班的職務', '選擇可兼顧家庭的職務',
       '選擇較有升遷或加薪機會的職務', '18.\t就您的認知，不同性別在工程與科技領域的求職難易度有否差異？',
       '19.\t在您或您認識的女性身上，性別在工程與科技領域的就業過程中存在的差異：（複選，同意即打勾）', '其他（請說明）',
       '20.\t對您的經驗，不同性別的工程與科技領域主管是否有領導風格的差異（指非因個人而是因為性別所致之差異）？',
       '21.\t您個人比較偏好與何種性別工程與科技領域的主管共事？',
       '22.\t不同性別工程與科技領域的主管領導風格差異主要顯現在哪些層面？（可複選至多三項）', '其他（請說明）：', '歡迎提供意見或建議',
       'Submitted From']
'''


# crosstab() Input --> margins=顯示總和、normalize=以百分比顯示

cross_tab_gender_to_age = pd.crosstab(statData["2. 年齡"],statData["1. 性別："],margins=True)

#output
cross_tab_gender_to_age.to_csv('2.性別年齡.csv',encoding='utf_8_sig')

#-----------------------------------------------------------------------------------------------------#
#最高學歷分析

df1 = statData[['1. 性別：','2. 年齡','3. 最高學歷']].sort_values(by=['1. 性別：']).reset_index(drop=True).rename(columns={'2. 年齡':'年齡',"1. 性別：":'性別','3. 最高學歷':'最高學歷'})

ct1 = ct.crosstab_generator(df1,'最高學歷',"年齡","性別","3.性別年齡_最高學歷")

#-----------------------------------------------------------------------------------------------------#
#畢業校系-->目前職務

df1 = statData[['1. 性別：','3. 最高學歷','4.\t最高學歷畢業校系','7. 是否正在從事工程與科技領域職務（含管理與學術研究）？','8. 您離開工程職務之最主要原因為何？請勾選最主要原因，至多二項。','(a) 工程與科技領域職務(年)', '(b) 非工程與科技領域職務','10.\t請問您目前服務的職務較接近下列何者？','請問目前受您管理的人員約＿＿人？','9.\t請問您目前的工作狀況為何？']]

#ct1 = ct.crosstab_generator(df1,'7. 是否正在從事工程與科技領域職務（含管理與學術研究）？','4.\t最高學歷畢業校系','1. 性別：',"性別最高學歷_目前職務")

#---年資---

#總年資


#數字處理
delList =[]
for i in range(len(df1)):
    if( (not str(df1['(a) 工程與科技領域職務(年)'][i]).isdigit()) or (not str(df1['(b) 非工程與科技領域職務'][i]).isdigit())) :
        delList.append(i)

df2 = df1.drop(delList).reset_index(drop=True)

t1 = df2['(a) 工程與科技領域職務(年)'].fillna(0).astype(int)
t2 = df2['(b) 非工程與科技領域職務'].fillna(0).astype(int)

df2['總年資']=t1+t2

#用範圍對df2分類
for i in range(len(df2)):
    v1 = int(df2['(a) 工程與科技領域職務(年)'][i])
    if( v1<=5 and v1>=0) :
        df2['(a) 工程與科技領域職務(年)'][i]='0~5年'
    elif( v1>5 and v1<11 ):
        df2['(a) 工程與科技領域職務(年)'][i]='6~10年'
    elif( v1>=11 and v1<16 ):
        df2['(a) 工程與科技領域職務(年)'][i]='11~15年'
    elif( v1>=16 and v1<21 ):
        df2['(a) 工程與科技領域職務(年)'][i]='16~20年'
    elif( v1>=21 and v1<26 ):
        df2['(a) 工程與科技領域職務(年)'][i]='21~25年'
    elif( v1>=26 and v1<31 ):
        df2['(a) 工程與科技領域職務(年)'][i]='26~30年'
    elif( v1>=31 and v1<100 ):
        df2['(a) 工程與科技領域職務(年)'][i]='30年以上'
    else:
        df2['(a) 工程與科技領域職務(年)'][i]=None

    v1 = int(df2['(b) 非工程與科技領域職務'][i])
    if( v1<=5 and v1>=0) :
        df2['(b) 非工程與科技領域職務'][i]='0~5年'
    elif( v1>5 and v1<11 ):
        df2['(b) 非工程與科技領域職務'][i]='6~10年'
    elif( v1>=11 and v1<16 ):
        df2['(b) 非工程與科技領域職務'][i]='11~15年'
    elif( v1>=16 and v1<21 ):
        df2['(b) 非工程與科技領域職務'][i]='16~20年'
    elif( v1>=21 and v1<26 ):
        df2['(b) 非工程與科技領域職務'][i]='21~25年'
    elif( v1>=26 and v1<31 ):
        df2['(b) 非工程與科技領域職務'][i]='26~30年'
    elif( v1>=31 and v1<100 ):
        df2['(b) 非工程與科技領域職務'][i]='30年以上'
    else:
        df2['(b) 非工程與科技領域職務'][i]=None

    v1 = int(df2['總年資'][i])
    if( v1<=5 and v1>=0) :
        df2['總年資'][i]='0~5年'
    elif( v1>5 and v1<11 ):
        df2['總年資'][i]='6~10年'
    elif( v1>=11 and v1<16 ):
        df2['總年資'][i]='11~15年'
    elif( v1>=16 and v1<21 ):
        df2['總年資'][i]='16~20年'
    elif( v1>=21 and v1<26 ):
        df2['總年資'][i]='21~25年'
    elif( v1>=26 and v1<31 ):
        df2['總年資'][i]='26~30年'
    elif( v1>=31 and v1<100 ):
        df2['總年資'][i]='30年以上'
    else:
        df2['總年資'][i]=None
        


ct2 = pd.crosstab(df2['4.\t最高學歷畢業校系'],[df2['(a) 工程與科技領域職務(年)'],df2['1. 性別：']],margins=True)
ct2 = ct2/ct2['All'][-1]
ct2 = ((ct2*100).round(1).astype(str)+'%').replace('0.0%','-')
ct3 = pd.crosstab(df2['4.\t最高學歷畢業校系'],[df2['(b) 非工程與科技領域職務'],df2['1. 性別：']],margins=True)
ct3 = ct3/ct3['All'][-1]
ct3 = ((ct3*100).round(1).astype(str)+'%').replace('0.0%','-')
ct4 = pd.crosstab(df2['4.\t最高學歷畢業校系'],[df2['總年資'],df2['1. 性別：']],margins=True)
ct4 = ct4/ct4['All'][-1]
ct4 = ((ct4*100).round(1).astype(str)+'%').replace('0.0%','-')


ct2.to_csv('6.1 工程領域年資.csv',encoding='utf_8_sig')
ct3.to_csv('6.2 非工程領域年資.csv',encoding='utf_8_sig')
ct4.to_csv('6.3 總年資.csv',encoding='utf_8_sig')

#------修改------#
df1 = statData[['1. 性別：','3. 最高學歷','4.\t最高學歷畢業校系','7. 是否正在從事工程與科技領域職務（含管理與學術研究）？','8. 您離開工程職務之最主要原因為何？請勾選最主要原因，至多二項。','10.\t請問您目前服務的職務較接近下列何者？']]
df2 = ct.group_table(df1,"4.\t最高學歷畢業校系",['國外學校，非工程與科技相關領域','國內學校，非工程與科技相關領域'])
df2 = ct.group_table(df2,"10.\t請問您目前服務的職務較接近下列何者？",["非工程與科技領域之管理職務","非工程與科技領域之非管理職務"])
delList = list(df2.index)
delList.sort()
df2 = df1.drop(delList) 
df2 = df2.reset_index(drop=True)
df2 = df2.dropna(subset=['8. 您離開工程職務之最主要原因為何？請勾選最主要原因，至多二項。']).reset_index(drop=True)
df2 = ct.extraction(df2,'8. 您離開工程職務之最主要原因為何？請勾選最主要原因，至多二項。','您離開工程職務之最主要原因為何？').reset_index(drop=True)
ct2 = pd.crosstab(df2['您離開工程職務之最主要原因為何？'],[df2['3. 最高學歷'],df2['1. 性別：']],margins=True)
ct2 = ct2[['博士','碩士','大學/大專','專科','高職','All']]
ct2 = ct2/ct2['All'][-1]
ct2 = ((ct2*100).round(1).astype(str)+'%').replace('0.0%','-')
ct2.to_csv('8.離開工程職務最主要原因分析.csv',encoding='utf_8_sig')
ct3 = pd.crosstab(df2['您離開工程職務之最主要原因為何？'],df2['1. 性別：'],margins=True)
ct3 = ct3/ct3['All'][-1]
ct3 = ((ct3*100).round(1).astype(str)+'%').replace('0.0%','-')
ct3.to_csv('8.離開工程職務最主要原因分析(性別).csv',encoding='utf_8_sig')



#最高學歷對目前工作狀況（第七題選否的）
ct5 = pd.crosstab(df5['9.\t請問您目前的工作狀況為何？'],[df5['3. 最高學歷'],df5['1. 性別：']],margins=True)

ct5.to_csv("9.目前的工作狀況（第七題選否）.csv",encoding='utf_8_sig')


df5 = ct.group_table(df1,'7. 是否正在從事工程與科技領域職務（含管理與學術研究）？','是')

ct5 = pd.crosstab(df5['10.\t請問您目前服務的職務較接近下列何者？'],df5['1. 性別：'],margins=True).sort_values(by=['All'],ascending=False).drop('All',axis=0)

df6 = ct.group_table(df5,'10.\t請問您目前服務的職務較接近下列何者？',['工程與科技領域之管理職務','非工程與科技領域之管理職務'])

df1 = df6.reset_index(drop=True)

delList =[]
for i in range(len(df1)):
    if( not str(df1['請問目前受您管理的人員約＿＿人？'][i]).isdigit() ) :
        delList.append(i)

df2 = df1.drop(delList).reset_index(drop=True)

#用範圍對df2分類
for i in range(len(df2)):
    v1 = int(df2['請問目前受您管理的人員約＿＿人？'][i])
    if( v1<=10 and v1>0) :
        df2['請問目前受您管理的人員約＿＿人？'][i]='1. 1~10人'
    elif( v1>10 and v1<31 ):
        df2['請問目前受您管理的人員約＿＿人？'][i]='2. 11~30人'
    elif( v1>=31 and v1<51 ):
        df2['請問目前受您管理的人員約＿＿人？'][i]='3. 31~50人'
    elif( v1>=51 and v1<101 ):
        df2['請問目前受您管理的人員約＿＿人？'][i]='4. 51~100人'
    elif( v1>=101 and v1<151 ):
        df2['請問目前受您管理的人員約＿＿人？'][i]='5. 101~150人'
    elif( v1>=151 and v1<201 ):
        df2['請問目前受您管理的人員約＿＿人？'][i]='6. 151~200人'
    elif( v1>=201 ):
        df2['請問目前受您管理的人員約＿＿人？'][i]='7. 200人以上'
    else:
        df2['請問目前受您管理的人員約＿＿人？'][i]=None


df3 = pd.crosstab(df2['請問目前受您管理的人員約＿＿人？'],[df2['10.\t請問您目前服務的職務較接近下列何者？'],df2['1. 性別：']],margins=True)

df3.to_csv('9.1 受您管理的人員.csv',encoding='utf_8_sig')

#-----------------------------------------------------------#

df1 = statData[['1. 性別：','10.\t請問您目前服務的職務較接近下列何者？','13.\t哪些福利措施最有助於您留在工程與科技領域就業？（不論現在是否有需求皆可選擇，可複選，最多三項）',
       '14.\t您服務的單位或待業前服務的單位提供哪些職場相關措施？（可複選）',
       '15.\t您認為工程與科技領域最需要改善的性別議題有哪些？請勾選其中最重要的議題，至多三項。',]]

df3 =  ct.extraction(df1,'14.\t您服務的單位或待業前服務的單位提供哪些職場相關措施？（可複選）','您服務的單位或待業前服務的單位提供哪些職場相關措施？')

ct3 = pd.crosstab(df3['10.\t請問您目前服務的職務較接近下列何者？'],[df3['您服務的單位或待業前服務的單位提供哪些職場相關措施？'],df3['1. 性別：']],margins=True).drop(' 請勿再勾選其他選項 )',axis=1)

ct3.to_csv("14.您服務的單位或待業前服務的單位提供哪些職場相關措施.csv",encoding='utf_8_sig')

df4 =  ct.extraction(df1,'15.\t您認為工程與科技領域最需要改善的性別議題有哪些？請勾選其中最重要的議題，至多三項。','您認為工程與科技領域最需要改善的性別議題有哪些？')

ct4 = pd.crosstab(df4['10.\t請問您目前服務的職務較接近下列何者？'],[df4['您認為工程與科技領域最需要改善的性別議題有哪些？'],df4['1. 性別：']],margins=True).drop(' 請勿再勾選其他選項 )',axis=1)

ct4.to_csv("15.您認為工程與科技領域最需要改善的性別議題有哪些.csv",encoding='utf_8_sig')


#-----上面待整理 有夠亂＝＝-----#


#第11題
df = statData[['1. 性別：','3. 最高學歷','11.\t您對於現任職務之五年後職涯發展的預期為何？請勾選最有可能的一項。']]

ct1 = pd.crosstab(df['11.\t您對於現任職務之五年後職涯發展的預期為何？請勾選最有可能的一項。'],[df['3. 最高學歷'],df['1. 性別：']],margins=True)

ct1.to_csv('11.對於現任職務之五年後職涯發展的預期.csv',encoding='utf_8_sig')

#第12題
df = statData[['1. 性別：','3. 最高學歷','12.\t請問您所預期的職涯發展，需要哪些配合因素來達成？請勾選最主要因素，至多二項。']]

for i in range(len(df)):
    if(not (df['12.\t請問您所預期的職涯發展，需要哪些配合因素來達成？請勾選最主要因素，至多二項。'][i].find('還不清楚（若勾選本項請勿再勾選其他項目）') == -1) ):
        df['12.\t請問您所預期的職涯發展，需要哪些配合因素來達成？請勾選最主要因素，至多二項。'][i] = '還不清楚'

df1 = ct.extraction(df,'12.\t請問您所預期的職涯發展，需要哪些配合因素來達成？請勾選最主要因素，至多二項。','請問您所預期的職涯發展，需要哪些配合因素來達成？')

ct1 = pd.crosstab(df1['請問您所預期的職涯發展，需要哪些配合因素來達成？'],[df1['3. 最高學歷'],df1['1. 性別：']],margins=True)

ct1.to_csv('12.預期的職涯發展需要哪些配合因素來達成.csv',encoding='utf_8_sig')

#第13題

df = statData[['1. 性別：','3. 最高學歷','13.\t哪些福利措施最有助於您留在工程與科技領域就業？（不論現在是否有需求皆可選擇，可複選，最多三項）']]

df1 = ct.extraction(df,'13.\t哪些福利措施最有助於您留在工程與科技領域就業？（不論現在是否有需求皆可選擇，可複選，最多三項）','哪些福利措施最有助於您留在工程與科技領域就業？')

ct1 = pd.crosstab(df1['哪些福利措施最有助於您留在工程與科技領域就業？'],[df1['3. 最高學歷'],df1['1. 性別：']],margins=True).drop(' 請勿再勾選其他選項 )')

ct1.to_csv('13.哪些福利措施最有助於您留在工程與科技領域就業.csv',encoding='utf_8_sig')

#第16題
df = statData[['1. 性別：','3. 最高學歷','16.\t在您或您認識的女性身上，性別在工程與科技領域的求學階段有否差異？']]

ct1 = pd.crosstab(df['16.\t在您或您認識的女性身上，性別在工程與科技領域的求學階段有否差異？'],[df['3. 最高學歷'],df['1. 性別：']],margins=True)

##觀察到的差異

df = statData[['1. 性別：','3. 最高學歷','16.\t在您或您認識的女性身上，性別在工程與科技領域的求學階段有否差異？','較用功或成績較好', '較擅長數理學科',
       '動手實驗或實習的機會較少', '遇到挫折時較易放棄', '較不敢接受挑戰', '較容易情緒起伏', '較受人際問題干擾',
       '較容易為情所困']]

df1 = ct.group_table(df,"16.\t在您或您認識的女性身上，性別在工程與科技領域的求學階段有否差異？",['有，我有觀察到差異：'])

checkList = ['較用功或成績較好', '較擅長數理學科',
       '動手實驗或實習的機會較少', '遇到挫折時較易放棄', '較不敢接受挑戰', '較容易情緒起伏', '較受人際問題干擾',
       '較容易為情所困']

for item in checkList:
    ct1 = pd.crosstab(df1[item],[df1['3. 最高學歷'],df1['1. 性別：']],margins=True)
    filename = '16.'+str(checkList.index(item))+' '+item+'.csv'
    ct1.to_csv(filename,encoding='utf_8_sig')

#第17~22題
df = statData[['1. 性別：','3. 最高學歷',"16.\t在您或您認識的女性身上，性別在工程與科技領域的求學階段有否差異？",
        '17.\t在您或您認識的工程與科技領域女性身上，性別在求職時的職務選擇方面有否差異？',
       '選擇內勤為主的職務', '選擇無須出差的職務', '選擇無須應酬的職務', '選擇無須輪班或值夜班的職務', '選擇可兼顧家庭的職務',
       '選擇較有升遷或加薪機會的職務','18.\t就您的認知，不同性別在工程與科技領域的求職難易度有否差異？',
       '19.\t在您或您認識的女性身上，性別在工程與科技領域的就業過程中存在的差異：（複選，同意即打勾）',
       '20.\t對您的經驗，不同性別的工程與科技領域主管是否有領導風格的差異（指非因個人而是因為性別所致之差異）？',
       '21.\t您個人比較偏好與何種性別工程與科技領域的主管共事？',
       '22.\t不同性別工程與科技領域的主管領導風格差異主要顯現在哪些層面？（可複選至多三項）']]

df0 = ct.group_table(df,"16.\t在您或您認識的女性身上，性別在工程與科技領域的求學階段有否差異？",['有，我有觀察到差異：'])

df1 = ct.group_table(df0,'17.\t在您或您認識的工程與科技領域女性身上，性別在求職時的職務選擇方面有否差異？',['有，我有觀察到差異：'])

checkList = ['選擇內勤為主的職務', '選擇無須出差的職務', '選擇無須應酬的職務', '選擇無須輪班或值夜班的職務', '選擇可兼顧家庭的職務',
       '選擇較有升遷或加薪機會的職務']

for item in checkList:
    ct1 = pd.crosstab(df1[item],[df1['3. 最高學歷'],df1['1. 性別：']],margins=True)
    filename = '17.'+str(checkList.index(item))+' '+item+'.csv'
    ct1.to_csv(filename,encoding='utf_8_sig')

ct1 = pd.crosstab(df1['18.\t就您的認知，不同性別在工程與科技領域的求職難易度有否差異？'],[df1['3. 最高學歷'],df1['1. 性別：']],margins=True)
ct1.to_csv('18.就您的認知，不同性別在工程與科技領域的求職難易度有否差異.csv',encoding='utf_8_sig')

df2 = ct.extraction(df1,'19.\t在您或您認識的女性身上，性別在工程與科技領域的就業過程中存在的差異：（複選，同意即打勾）','在您或您認識的女性身上，性別在工程與科技領域的就業過程中存在的差異')

ct1 = pd.crosstab(df2['在您或您認識的女性身上，性別在工程與科技領域的就業過程中存在的差異'],[df2['3. 最高學歷'],df2['1. 性別：']],margins=True).drop(' 請勿再勾選其他選項 )')
ct1.to_csv('19.就您的認知，不同性別在工程與科技領域的求職難易度有否差異.csv',encoding='utf_8_sig')

ct1 = pd.crosstab(df1['20.\t對您的經驗，不同性別的工程與科技領域主管是否有領導風格的差異（指非因個人而是因為性別所致之差異）？'],[df1['3. 最高學歷'],df1['1. 性別：']],margins=True)
ct1.to_csv('20.對您的經驗，不同性別的工程與科技領域主管是否有領導風格的差異.csv',encoding='utf_8_sig')

ct1 = pd.crosstab(df1['21.\t您個人比較偏好與何種性別工程與科技領域的主管共事？'],[df1['3. 最高學歷'],df1['1. 性別：']],margins=True)
ct1.to_csv('21.您個人比較偏好與何種性別工程與科技領域的主管共事.csv',encoding='utf_8_sig')

df2 = ct.extraction(df1,'22.\t不同性別工程與科技領域的主管領導風格差異主要顯現在哪些層面？（可複選至多三項）','不同性別工程與科技領域的主管領導風格差異主要顯現在哪些層面？')

ct1 = pd.crosstab(df2['不同性別工程與科技領域的主管領導風格差異主要顯現在哪些層面？'],[df2['3. 最高學歷'],df2['1. 性別：']],margins=True)
ct1.to_csv('22.不同性別工程與科技領域的主管領導風格差異主要顯現在哪些層面.csv',encoding='utf_8_sig')
