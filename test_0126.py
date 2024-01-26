""" import pandas as pd
data_frame=pd.read_excel('D:\code\parsexls\aaa.xls')
writer=pd.ExcelWriter('D:\code\parsexls\bbb.xls')
data_frame.to_excel(writer,sheet_name='xbbbb',index=False)
writer.save() """

from cmath import isnan, nan, tan
import math
import pandas as pd
import re

with open("TM461-G4-CA.INI", "r") as file:
    lines = file.readlines()
    lines = [line.strip() for line in lines if line.strip()]
    content = '\n'.join(lines)
    cleaned_content = re.sub(r'[a-zA-Z]', '', content)
    distance = cleaned_content.split()
# xls = pd.ExcelFile("x.xlsx")
# sheet_names = xls.sheet_names
# for name_id in sheet_names:
#     print(name_id)
#     df_existing = pd.read_excel("x.xlsx", sheet_name = name_id)
#     df_existing.insert(0, "New Column", lines)
#     df_existing.to_excel("x.xlsx", sheet_name = name_id, index=False)
    

def read_excel(filename,sheetname):
    df=pd.read_excel(filename,sheet_name=sheetname)
    print(df.columns)
    tof_channel_result={}
    plarity_result={}
    # [tof_channel_result.update({col:[]}) for col in df.columns[1:]]
    doudong_result={}
    col_key = df.columns[1:]
    last_tof_nums = 0
    plane_num =0
    doudong_num =0
    for index in range(0,df.index.size):
        row=list(df.loc[index])
        if row[0]=="TOF多通道深度精度":
            for key in col_key:
                v=df.loc[index][key]
                tof_index = 0
                if not pd.isna(v):
                    tofs=str(v).split('|')
                    for tof_v in tofs:
                        tof_key=str(key)+"_tof_"+str(tof_index)
                        tof_index += 1
                        if tof_key in tof_channel_result.keys():
                            tof_channel_result[tof_key].append(tof_v)
                        else:
                            tof_channel_result[tof_key]=[tof_v]
                    if last_tof_nums == 0:
                        last_tof_nums = len(tofs)

                # 需补齐相同行
                if last_tof_nums > tof_index:
                    for i in range(tof_index,last_tof_nums):
                        tof_key=str(key)+"_tof_"+str(i)
                        if tof_key in tof_channel_result.keys():
                            tof_channel_result[tof_key].append('NA')
                        else:
                            tof_channel_result[tof_key]=['NA']
        if row[0]=="平面度":
            for key in col_key:
                v=df.loc[index][key]
                tof_index = 0
                if not pd.isna(v):
                    plane=str(v).split('|')
                    for tof_v in plane:
                        tof_key=str(key)+"_plane_"+str(tof_index)
                        tof_index += 1
                        if tof_key in plarity_result.keys():
                            plarity_result[tof_key].append(tof_v)
                        else:
                            plarity_result[tof_key]=[tof_v]
                    if plane_num == 0:
                        plane_num = len(plane)

                # 需补齐相同行
                if plane_num > tof_index:
                    for i in range(tof_index,plane_num):
                        tof_key=str(key)+"_plane_"+str(i)
                        if tof_key in plarity_result.keys():
                            plarity_result[tof_key].append('NA')
                        else:
                            plarity_result[tof_key]=['NA']
        if row[0]=="单点抖动幅度":
            for key in col_key:
                v=df.loc[index][key]
                tof_index = 0
                if not pd.isna(v):
                    doudong=str(v).split('|')
                    for tof_v in doudong:
                        tof_key=str(key)+"_doudong_"+str(tof_index)
                        tof_index += 1
                        if tof_key in doudong_result.keys():
                            doudong_result[tof_key].append(tof_v)
                        else:
                            doudong_result[tof_key]=[tof_v]
                    if doudong_num == 0:
                        doudong_num = len(doudong)

                # 需补齐相同行
                if doudong_num > tof_index:
                    for i in range(tof_index,doudong_num):
                        tof_key=str(key)+"_plane_"+str(i)
                        if tof_key in doudong_result.keys():
                            doudong_result[tof_key].append('NA')
                        else:
                            doudong_result[tof_key]=['NA']



    with pd.ExcelWriter(".\\x.xlsx") as writer:
        tof_channel=pd.DataFrame(tof_channel_result)
        tof_channel.to_excel(writer,sheet_name ='TOF多通道深度精度',index=False)                        
        for i in range(0,last_tof_nums):
            data_of_tof = {}
            for key,value in tof_channel_result.items():
                if 'tof_{}'.format(i) in key:
                    data_of_tof.update({key:value})
                tof_channel=pd.DataFrame(data_of_tof)
                tof_channel.insert(0, "Distance", distance)
                tof_channel.to_excel(writer,sheet_name='tof{}'.format(i),index=False)
        plarity=pd.DataFrame(plarity_result)
        plarity.to_excel(writer,sheet_name ='平面度',index=False)                        
        for i in range(0,plane_num):
            data_of_tof = {}
            for key,value in plarity_result.items():
                if 'plane_{}'.format(i) in key:
                    data_of_tof.update({key:value})
                plarity=pd.DataFrame(data_of_tof)
                plarity.insert(0, "Distance", distance)
                plarity.to_excel(writer,sheet_name='plane_{}'.format(i),index=False)


        jitter=pd.DataFrame(doudong_result)
        jitter.to_excel(writer,sheet_name ='单点抖动幅度',index=False)                        
        for i in range(0,doudong_num):
            data_of_tof = {}
            for key,value in doudong_result.items():
                if 'plane_{}'.format(i) in key:
                    data_of_tof.update({key:value})
                jitter=pd.DataFrame(data_of_tof)
                jitter.insert(0, "Distance", distance)
                jitter.to_excel(writer,sheet_name='单点抖动幅度_{}'.format(i),index=False)        
        writer._save()
       

""""#print(data_frame.index)
print(data_frame.columns)
print(data_frame.loc[0])"""
""" writer=pd.ExcelWriter('D:\code\parsexls\bbb.xls')
data_frame.to_excel(writer,sheet_name='xbbbb',index=False)
writer.save() """

# print(data_frame)

if __name__=="__main__":
    read_excel("TM460.xlsx","1")
    pass