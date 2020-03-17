import pandas as pd
import os
os.getcwd()

# def double_crosstab_generator(statData,data1,data2,data3,filename):
#     #2to1
#     #中括號有bug
#     ct1 = pd.crosstab([statData[data1],statData[data2]],
#                                     statData[data3], normalize='index')
#     ct2 = pd.crosstab([statData[data1],statData[data2]],
#                                     statData[data3], normalize='columns')

#     ct1 = ((ct1*100).round(2).astype(str)+'%').replace('0.0%','-')
#     ct2 = ((ct2*100).round(2).astype(str)+'%').replace('0.0%','-')
    
#     ct1.to_csv(filename+'（橫向）.csv',encoding='utf_8_sig')
#     ct2.to_csv(filename+'（直向）.csv',encoding='utf_8_sig')


# def crosstab_generator(statData,data1,data2,data3,filename):
#     #1to2
    
#     ct1 = pd.crosstab(statData[data1],[statData[data2],
#                                     statData[data3]],margins=True)

#     #ct1 = ((ct1*100).round(2).astype(str)+'%').replace('0.0%','-')
    
#     ct1.to_csv(filename+'.csv',encoding='utf_8_sig')

#     return ct1

def group_table(statData,columns,data=[]):
    
    #分類一個內容下的多種不同資料並合併成一個DataFrame

    dfg1 = statData.groupby(columns,sort=False)
    mixData = []
    for contents in data:
        mixData.append(dfg1.get_group(contents))

    return pd.concat([c for c in mixData],axis=0)

def extraction(statData,title,changeName=None):
    t = statData[title].str.split(',',expand=True).stack()
    if(changeName==None):
        changeName=title
    t = t.reset_index(level=1,drop=True).rename(changeName)

    statData = statData.drop(title,axis=1)

    return statData.join(t)

def Insert_row(row_number, df, row_value): 
    # Starting value of upper half 
    start_upper = 0
   
    # End value of upper half 
    end_upper = row_number 
   
    # Start value of lower half 
    start_lower = row_number 
   
    # End value of lower half 
    end_lower = df.shape[0] 
   
    # Create a list of upper_half index 
    upper_half = [*range(start_upper, end_upper, 1)] 
   
    # Create a list of lower_half index 
    lower_half = [*range(start_lower, end_lower, 1)] 
   
    # Increment the value of lower half by 1 
    lower_half = [x.__add__(1) for x in lower_half] 
   
    # Combine the two lists 
    index_ = upper_half + lower_half 
   
    # Update the index of the dataframe 
    df.index = index_ 
   
    # Insert a row at the end 
    df.loc[row_number] = row_value 
    
    # Sort the index labels 
    df = df.sort_index() 
   
    # return the dataframe 
    return df 