import pandas as pd
import os

#os.chdir("/Users/cilab/PartTime_PythonAnalysis/group")
os.chdir(r"C:\Users\Naichen\Documents\GitHub\stu00608.github.io\PartTime_PythonAnalysis\group")
os.getcwd()

import ct_tool as ct

#記得路徑不同電腦要改

statData = pd.read_excel("003.xlsx",skiprows=4)
#statData = statData.dropna(axis=1,how='all')

#os.chdir("/Users/cilab/PartTime_PythonAnalysis/group/outputs/analysis")
os.chdir(r"C:\Users\Naichen\Documents\GitHub\stu00608.github.io\PartTime_PythonAnalysis\group\outputs\analysis")
os.getcwd()

#土木營建

df = statData[['1.單位名稱','2.單位總員工人數','土木營建','土木營建領域專長 男性 專任 人數','土木營建領域專長 女性 專任 人數','服務年資1~5年',
 '服務年資6~10年','服務年資11~15年','服務年資16~20年','服務年資16~20年.1','服務年資21~25年','服務年資25年以上','服務年資1~5年.1',
 '服務年資6~10年.1','服務年資11~15年.1','服務年資16~20年.2','服務年資21~25年.1','服務年資25年以上.1',]]

df = df.dropna(subset=['土木營建'],axis=0)
