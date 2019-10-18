import pandas as pd
import os
import ct_tool as ct


#記得路徑不同電腦要改
os.chdir("/Users/cilab/PartTime_PythonAnalysis/group")
os.getcwd()
statData = pd.read_csv("002.csv")
statData = statData.dropna(axis=1,how='all')
os.chdir("/Users/cilab/PartTime_PythonAnalysis/group/outputs")
os.getcwd()



