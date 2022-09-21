""" import pandas as pd
data_frame=pd.read_excel('D:\code\parsexls\aaa.xls')
writer=pd.ExcelWriter('D:\code\parsexls\bbb.xls')
data_frame.to_excel(writer,sheet_name='xbbbb',index=False)
writer.save() """

from cmath import tan
import math
import pandas as pd


    
def read_excel(filename,sheetname):
    df=pd.read_excel(filename,sheet_name=sheetname)
    print(df.columns)
    result={}
    # [result.update({col:[]}) for col in df.columns[1:]]


    col_key=df.columns[1:]
    for index in range(0,df.index.size):
        row=list(df.loc[index])
        if row[0]=="平面度":
            for key in col_key:
                v=df.loc[index][key]
                tofs=v.split('|')
                tof_index=0
                for tof_v in tofs:
                    tof_key=str(key)+"_RMS_"+str(tof_index)
                    tof_index+=1
                    if tof_key in result.keys():
                        result[tof_key].append(tof_v)
                    else:
                        result[tof_key]=[tof_v]
                
   
    with pd.ExcelWriter(".\\tof_r.xlsx") as writer:
        new_df=pd.DataFrame(result)
        new_df.to_excel(writer,'Sheet1',index=False)
        writer.save()

    pass
#result = data_frame.loc[data['Distance']]
# print(data_frame.head(3))

""""#print(data_frame.index)
print(data_frame.columns)
print(data_frame.loc[0])"""
""" writer=pd.ExcelWriter('D:\code\parsexls\bbb.xls')
data_frame.to_excel(writer,sheet_name='xbbbb',index=False)
writer.save() """

# print(data_frame)

if __name__=="__main__":
    read_excel("cccc.xlsx","TL460-GTI8-LD-CUT")
    pass