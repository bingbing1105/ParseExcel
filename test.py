""" import pandas as pd
data_frame=pd.read_excel('D:\code\parsexls\aaa.xls')
writer=pd.ExcelWriter('D:\code\parsexls\bbb.xls')
data_frame.to_excel(writer,sheet_name='xbbbb',index=False)
writer.save() """

from cmath import isnan, nan, tan
import math
import pandas as pd


    

def read_excel(filename,sheetname):
    df=pd.read_excel(filename,sheet_name=sheetname)
    print(df.columns)
    result={}
    # [result.update({col:[]}) for col in df.columns[1:]]

    col_key=df.columns[1:]
    last_tof_nums=0
    for index in range(0,df.index.size):
        row=list(df.loc[index])
        if row[0]=="单点抖动幅度":
            for key in col_key:
                v=df.loc[index][key]
                tof_index=0

                if not pd.isna(v):
                    tofs=str(v).split('|')
                    for tof_v in tofs:
                        tof_key=str(key)+"_tof_"+str(tof_index)
                        tof_index+=1
                        if tof_key in result.keys():
                            result[tof_key].append(tof_v)
                        else:
                            result[tof_key]=[tof_v]
                    if last_tof_nums==0:
                        last_tof_nums=tof_index

                # 需补齐相同行
                if last_tof_nums> tof_index:
                    for i in range(tof_index,last_tof_nums):
                        tof_key=str(key)+"_tof_"+str(i)
                        if tof_key in result.keys():
                            result[tof_key].append('NA')
                        else:
                            result[tof_key]=['NA']


    # for key,value in result.items():
    #     print("col[%s] len=%s"%(key,len(value)))         
    # 
    with pd.ExcelWriter(".\\x.xlsx") as writer:
        new_df=pd.DataFrame(result)
        new_df.to_excel(writer,sheet_name='Sheet1',index=False)
        # tof_0
        data_of_tof0={}
        data_of_tof1={}
        data_of_tof2={}
        data_of_tof3={}
        for key,value in result.items():
            if 'tof_0' in key:
                data_of_tof0.update({key:value})
            elif 'tof_1' in key:
                data_of_tof1.update({key:value})
            elif 'tof_2' in key:
                data_of_tof2.update({key:value})
            elif 'tof_3' in key:
                data_of_tof3.update({key:value})
        new_df=pd.DataFrame(data_of_tof0)
        new_df.to_excel(writer,sheet_name='tof0',index=False)
        new_df=pd.DataFrame(data_of_tof1)
        new_df.to_excel(writer,sheet_name='tof1',index=False)
        new_df=pd.DataFrame(data_of_tof2)
        new_df.to_excel(writer,sheet_name='tof2',index=False)
        new_df=pd.DataFrame(data_of_tof3)
        new_df.to_excel(writer,sheet_name='tof3',index=False)
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
    read_excel("ccc.xlsx","AS830-GTIS8-45-F121-CGP")
    pass