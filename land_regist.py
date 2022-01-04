import tkinter as tk
from tkinter import filedialog,dialog
import os
from pathlib import Path
import pandas as pd
import csv
import numpy as np
import warnings
import sys  
warnings.filterwarnings("ignore") 
#參考https://www.796t.com/article.php?id=10889
sys.setrecursionlimit(2000000)

window = tk.Tk()
window.title('地籍謄本拆分') # 標題
window.geometry('300x150') # 視窗尺寸

file_path = ''

file_text = ''

def strQ2B(s):
    rstring = ""
    for uchar in s:
        u_code = ord(uchar)
        if u_code == 12288:  # 全形空格直接轉換
            u_code = 32
        elif 65281 <= u_code <= 65374:  # 全形字元（除空格）根據關係轉化
            u_code -= 65248
        rstring += chr(u_code)
    return rstring

def all_combine():
    ## 讀取資料
    global file_path
    global file_text
    file_path = filedialog.askopenfilenames(title=u'選擇檔案',initialdir=(os.path.expanduser('C:/Users/Desktop')))
    print('開啟檔案：',file_path)
    if file_path is not None:
        address= list(file_path)
        print(address)
    for i in address:
        df1 = pd.read_csv(i,header=None,engine="python",encoding='utf-8')
        df_str1 = df1[0].astype(str)
        q2b1=[]
        for i in range(len(df_str1)):
            a = strQ2B(df_str1.loc[i])
            q2b1.append(a)

        df_new1 = pd.DataFrame(q2b1,columns=[0])
    ##############################字元取代&格式整理
        for i in range(len(df_new1)):
            df_new1[0][i] = df_new1[0][i].replace("標的登記次序"," 標的登記次序")
            df_new1[0][i] = df_new1[0][i].replace("設定權利範圍"," 設定權利範圍")
            df_new1[0][i] = df_new1[0][i].replace("地號合併自","地號 合併自")
            df_new1[0][i] = df_new1[0][i].replace("本謄本僅係 ","本謄本僅係")
            df_new1[0][i] = df_new1[0][i].replace("清償日期"," 清償日期")
            df_new1[0][i] = df_new1[0][i].replace("債權額比例"," 債權額比例")
            df_new1[0][i] = df_new1[0][i].replace("相關他項權利登記次序"," 相關他項權利登記次序")
            df_new1[0][i] = df_new1[0][i].replace("原因發生日期"," 原因發生日期")
            df_new1[0][i] = df_new1[0][i].replace("統一編號"," 統一編號")
            df_new1[0][i] = df_new1[0][i].replace("權狀字號"," 權狀字號")
            df_new1[0][i] = df_new1[0][i].replace("全部 ","全部")
            df_new1[0][i] = df_new1[0][i].replace("                                                          CE","")
            df_new1[0][i] = df_new1[0][i].replace("主  任 ","主任 ")
            df_new1[0][i] = df_new1[0][i].replace("分割自:","分割自")
            df_new1[0][i] = df_new1[0][i].replace("主 任  ","主任")
            df_new1[0][i] = df_new1[0][i].replace("違 約 金","違約金")
            df_new1[0][i] = df_new1[0][i].replace("面    積"," 面積")
            df_new1[0][i] = df_new1[0][i].replace("權 利 人","權利人")
            df_new1[0][i] = df_new1[0][i].replace("住    址","住址")
            df_new1[0][i] = df_new1[0][i].replace("  *","")
            df_new1[0][i] = df_new1[0][i].replace("*","")
            df_new1[0][i] = df_new1[0][i].replace("(","")
            df_new1[0][i] = df_new1[0][i].replace(")","")
            df_new1[0][i] = df_new1[0][i].replace("〈","")
            df_new1[0][i] = df_new1[0][i].replace("〉","")
            df_new1[0][i] = df_new1[0][i].replace("空白","無 ")
        b1 = pd.DataFrame(df_new1[0].str.split(' ').tolist()).stack()
        b1.columns = ['B', 'index_col']
        b1 = b1.reset_index()[0]
        b1 = pd.DataFrame(b1)
        right_df1 = b1[b1[0]!='']
        right_df1 = pd.DataFrame(right_df1)
        colnames=['col']
        right_df1.columns = colnames
        right_df1.index = range(1,len(right_df1)+1)
        right_df1 = right_df1['col'].str.split(':',expand=True)
        colnames=['col1','col2','col3']
        right_df1.columns = colnames

        ##############################找出擷取節點
        sep_index1 = right_df1[right_df1['col1']=="土地標示部"].index.tolist()[:]
        sep_index2 = right_df1[right_df1['col1']=="土地所有權部"].index.tolist()[:]
        sep_index3 = right_df1[right_df1['col1']=="土地他項權利部"].index.tolist()[:]
        sep_index4 = right_df1[right_df1['col1']=="本謄本列印完畢"].index.tolist()[:]
        sep_index5 = right_df1[right_df1['col1']==right_df1.loc[1][0]].index.tolist()[:]

        ##############################謄本格式判別 
        land_mark = pd.DataFrame(columns=['col1','col2','col3'])
        land_owner = pd.DataFrame(columns=['col1','col2','col3'])
        land_others = pd.DataFrame(columns=['col1','col2','col3'])
        land_combine = pd.DataFrame(columns=['col1','col2','col3'])


        for i in range(len(sep_index1)):
            if sep_index3[0]>sep_index4[i]:
                #landmark
                land_mark.loc[len(land_mark)+1] = right_df1['col1'][sep_index5[i]+1]+" "+right_df1['col1'][sep_index5[i]+2]
                lm = right_df1.loc[sep_index1[i]:sep_index2[i]-1]
                land_mark = land_mark.append(lm)
                land_mark = land_mark.fillna(np.nan)
                land_mark.index= range(1,len(land_mark)+1)


                #landowner
                land_owner.loc[len(land_owner)+1] = right_df1['col1'][sep_index5[i]+1]+" "+right_df1['col1'][sep_index5[i]+2]
                lo = right_df1.loc[sep_index2[i]:sep_index4[i]-1]
                land_owner = land_owner.append(lo)
                land_owner = land_owner.fillna(np.nan)
                land_owner.index= range(1,len(land_owner)+1)

                land_combine.loc[len(land_combine)+1] = right_df1['col1'][sep_index5[i]+1]+" "+right_df1['col1'][sep_index5[i]+2]
                land_combine = land_combine.append([lm,lo])
                land_combine.index= range(1,len(land_combine)+1)

            else:
                for j in range(len(sep_index3)):
                    #landmark
                    land_mark.loc[len(land_mark)+1] = right_df1['col1'][sep_index5[i+j]+1]+" "+right_df1['col1'][sep_index5[i+j]+2]
                    lm = right_df1.loc[sep_index1[i+j]:sep_index2[i+j]-1]
                    land_mark = land_mark.append(lm)
                    land_mark = land_mark.fillna(np.nan)
                    land_mark.index= range(1,len(land_mark)+1)


                    #landowner
                    land_owner.loc[len(land_owner)+1] = right_df1['col1'][sep_index5[i+j]+1]+" "+right_df1['col1'][sep_index5[i+j]+2]
                    lo = right_df1.loc[sep_index2[i+j]:sep_index3[j]-1]
                    land_owner = land_owner.append(lo)
                    land_owner = land_owner.fillna(np.nan)
                    land_owner.index= range(1,len(land_owner)+1)


                    #land_others
                    land_others.loc[len(land_others)+1] = right_df1['col1'][sep_index5[i+j]+1]+" "+right_df1['col1'][sep_index5[i+j]+2]
                    l_others = right_df1.loc[sep_index3[j]:sep_index4[i+j]-1]
                    land_others = land_others.append(l_others)
                    land_others = land_others.fillna(np.nan)
                    land_others.index= range(1,len(land_others)+1)

                    land_combine.loc[len(land_combine)+1] = right_df1['col1'][sep_index5[i+j]+1]+" "+right_df1['col1'][sep_index5[i+j]+2]
                    land_combine = land_combine.append([lm,lo,l_others])
                    land_combine.index= range(1,len(land_combine)+1)
                break

        ##############
        index_owner = land_combine[land_combine['col1']=="前次移轉現值或原規定地價"].index.tolist()[:]
        for i in range(len(index_owner)):
            land_combine.loc[index_owner[i]][1] = land_combine.loc[index_owner[i]+1][0]+land_combine.loc[index_owner[i]+2][0]
            land_combine = land_combine.drop([index_owner[i]+1,index_owner[i]+2],axis=0)

        index_owner1 = land_combine[land_combine['col1']=="本謄本僅係所有權個人全部節本,詳細權利狀態請參閱全部謄本"].index.tolist()
        for i in range(len(index_owner1)):
            if index_owner1 is not None:
                land_combine = land_combine.drop(list(range(index_owner1[i],index_owner1[i]+1)))
            else:
                pass

        ##############################
        others_index1 = land_combine[land_combine['col1']=="擔保債權總金額"].index.tolist()
        others_index2 = land_combine[land_combine['col1']=="擔保債權確定期日"].index.tolist()
        others_index3 = land_combine[land_combine['col1']=="權利標的"].index.tolist()
        others_index4 = land_combine[land_combine['col1']=="設定權利範圍"].index.tolist()

        for i in range(len(others_index1)):
            drop_list = list(range(others_index1[i]+1,others_index2[i]))+list(range(others_index2[i]+1,others_index3[i]))+list(range(others_index3[i]+1,others_index4[i]))
            land_combine = land_combine.drop(drop_list,axis=0)

        land_combine.index= range(1,len(land_combine)+1)
        print(i)
        filename_save=land_combine.loc[1][0]+".xlsx"
        savefile = land_combine.to_excel(filename_save,encoding='big5',index=0)
    
bt1 = tk.Button(window,text='轉換檔案',width=15,height=2,command=all_combine).pack(anchor=tk.CENTER,pady=50)

window.mainloop() # 顯示