""" import pandas as pd
data_frame=pd.read_excel('D:\code\parsexls\aaa.xls')
writer=pd.ExcelWriter('D:\code\parsexls\bbb.xls')
data_frame.to_excel(writer,sheet_name='xbbbb',index=False)
writer.save() """

from cmath import isnan, nan, tan
import math
import pandas as pd
import re
import os

with open("AS850-GT6-410-CB0S0.INI", "r") as file:
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
    juli_result = {}
    Hfov_result={}
    Vfov_result={}
    col_key = df.columns[1:]
    last_tof_nums = 0
    plane_num =0
    doudong_num =0
    juli_num =0
    Hfov_num=0
    Vfov_num=0
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
        if row[0]=="水平视场角":
            for key in col_key:
                v=df.loc[index][key]
                tof_index = 0
                if not pd.isna(v):
                    Hfov=str(v).split('|')
                    for tof_v in Hfov:
                        tof_key=str(key)+"_Hfov_"+str(tof_index)
                        tof_index += 1
                        if tof_key in Hfov_result.keys():
                            Hfov_result[tof_key].append(tof_v)
                        else:
                            Hfov_result[tof_key]=[tof_v]
                    if Hfov_num == 0:
                        Hfov_num = len(Hfov)

                # 需补齐相同行
                if Hfov_num > tof_index:
                    for i in range(tof_index,Hfov_num):
                        tof_key=str(key)+"_Hfov_"+str(i)
                        if tof_key in Hfov_result.keys():
                            Hfov_result[tof_key].append('NA')
                        else:
                            Hfov_result[tof_key]=['NA']
        if row[0]=="垂直视场角":
            for key in col_key:
                v=df.loc[index][key]
                tof_index = 0
                if not pd.isna(v):
                    Vfov=str(v).split('|')
                    for tof_v in Vfov:
                        tof_key=str(key)+"_Vfov_"+str(tof_index)
                        tof_index += 1
                        if tof_key in Vfov_result.keys():
                            Vfov_result[tof_key].append(tof_v)
                        else:
                            Vfov_result[tof_key]=[tof_v]
                    if Vfov_num == 0:
                        Vfov_num = len(Vfov)

                # 需补齐相同行
                if Vfov_num > tof_index:
                    for i in range(tof_index,Vfov_num):
                        tof_key=str(key)+"_Vfov_"+str(i)
                        if tof_key in Vfov_result.keys():
                            Vfov_result[tof_key].append('NA')
                        else:
                            Vfov_result[tof_key]=['NA']
        if row[0]=="Distance":
            for key in col_key:
                v=df.loc[index][key]
                tof_index = 0
                if not pd.isna(v):
                    juli=str(v).split('|')
                    for tof_v in juli:
                        tof_key=str(key)+"_juli_"+str(tof_index)
                        tof_index += 1
                        if tof_key in juli_result.keys():
                            juli_result[tof_key].append(tof_v)
                        else:
                            juli_result[tof_key]=[tof_v]
                    if juli_num == 0:
                        juli_num = len(juli)

                # 需补齐相同行
                if juli_num > tof_index:
                    for i in range(tof_index,juli_num):
                        tof_key=str(key)+"_juli_"+str(i)
                        if tof_key in juli_result.keys():
                            juli_result[tof_key].append('NA')
                        else:
                            juli_result[tof_key]=['NA']
    reslut_name=str(sheetname)+'-'+str(df.columns[1])+".xlsx"
    with pd.ExcelWriter(reslut_name) as writer:
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
                if 'doudong_{}'.format(i) in key:
                    data_of_tof.update({key:value})
                jitter=pd.DataFrame(data_of_tof)
                jitter.insert(0, "Distance", distance)
                jitter.to_excel(writer,sheet_name='单点抖动幅度_{}'.format(i),index=False)

        jitter=pd.DataFrame(Hfov_result)
        jitter.to_excel(writer,sheet_name ='水平视场角',index=False)
        for i in range(0,Hfov_num):
            data_of_tof = {}
            for key,value in Hfov_result.items():
                if 'Hfov_{}'.format(i) in key:
                    data_of_tof.update({key:value})
                jitter=pd.DataFrame(data_of_tof)
                jitter.insert(0, "Distance", distance)
                jitter.to_excel(writer,sheet_name='水平视场角_{}'.format(i),index=False)

        jitter=pd.DataFrame(Vfov_result)
        jitter.to_excel(writer,sheet_name ='垂直视场角',index=False)
        for i in range(0,Vfov_num):
            data_of_tof = {}
            for key,value in Vfov_result.items():
                if 'Vfov_{}'.format(i) in key:
                    data_of_tof.update({key:value})
                jitter=pd.DataFrame(data_of_tof)
                jitter.insert(0, "Distance", distance)
                jitter.to_excel(writer,sheet_name='垂直视场角_{}'.format(i),index=False)

        Distance=pd.DataFrame(juli_result)
        Distance.to_excel(writer,sheet_name ='距离',index=False)
        for i in range(0,juli_num):
            data_of_tof = {}
            for key,value in juli_result.items():
                if 'juli_{}'.format(i) in key:
                    data_of_tof.update({key:value})
                jitter=pd.DataFrame(data_of_tof)
                jitter.insert(0, "Distance", distance)
                jitter.to_excel(writer,sheet_name='距离_{}'.format(i),index=False)
        # writer._save()
       

""""#print(data_frame.index)
print(data_frame.columns)
print(data_frame.loc[0])"""
""" writer=pd.ExcelWriter('D:\code\parsexls\bbb.xls')
data_frame.to_excel(writer,sheet_name='xbbbb',index=False)
writer.save() """

# print(data_frame)

def get_all_excel():
    dir=os.path.abspath(__file__)
    dir=os.path.dirname(dir)
    # print( dir)
    file_list = []
    # for root, dirs, files in os.walk(dir):
    #     for file in files:
    #         if file.endswith('.xlsx'):
    #             file_list.append(os.path.join(root, file))
    #             filename=file
    for file in os.listdir( dir):
        if file.endswith('.xlsx'):
            filexlsx=file
        if file.endswith('.INI'):
            fileini=file.split( '.')[0]
        file_list.append(os.path.join(dir, file))
    # print('filename:{}'.format(filename))
    return filexlsx,fileini,

def main():
    file_xlsz,file_ini = get_all_excel()
    # print('file_xlsz:{}'.format(file_xlsz))
    # print('file_ini:{}'.format(file_ini))
    read_excel(file_xlsz,file_ini)
    pass

if __name__=="__main__":
    main()
    # read_excel("qa_xml_qa_result_2024. 2.21_17. 0.xlsx","PS801-TIS3-410-CB0S0-MT4")
    pass