#!/usr/bin/env python
# coding: utf-8

import datetime
import os
import re
import shutil
import time
import zipfile
import chromedriver_autoinstaller
import numpy as np
import pandas as pd
import selenium.webdriver.support.ui as ui
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

# from webdriver_manager.chrome import ChromeDriverManager
# driver = webdriver.Chrome(executable_path='位置/chromedriver',chrome_options=opts)
##############
chromedriver_autoinstaller.install()
# 下載時，依下載順序新建買賣、預售屋、租賃三個資料夾，下載完一類後再開第二個資料夾

# 宣告全域變數，讓變數可以在函式內使用

global path
global pre_vol_filename,pre_vol_namelist
global path_main_list

#取得程式碼(這個檔案)所在的絕對路徑
#note:在本機端和NAS上的路徑會差一個"/"，換地方執行需要注意
path =os.path.dirname(os.path.realpath(__file__))+'\\'



#建立一個時間log來紀錄更新執行時間

if not os.path.exists('log.txt'):
    now = datetime.datetime.now()
    txt = '上次更新時間為：' + str(now)
    newfile=open('log.txt', "a+")
    newfile.write(txt+'\n')
    newfile.close()
else:
    logfile = 'log.txt'
    now = datetime.datetime.now()
    txt = '上次更新時間為：' + str(now)
    with open(logfile, 'a') as f:
        f.write(txt+'\n')

# 前季下載&本期下載
def land_dl(data_type_file,data_type):
    #指定下載路徑
    out_path = path+data_type_file  # 下載想指定的路徑
    #webdriver執行時chrome的瀏覽器設定
    #關掉彈出視窗、設定下載資料夾、不要跳出帳號密碼管理員、允許自動下載(重要!)
    prefs = {'profile.default_content_settings.popups': 0,
     'download.default_directory': out_path,
     "profile.password_manager_enabled": False,
     "credentials_enable_service": False,
     "profile.default_content_setting_values.automatic_downloads" : 1}
    options = webdriver.ChromeOptions()#預設瀏覽器設定
    options.add_argument("--disable-infobars")
    options.add_argument("--disable-extensions")
    options.add_experimental_option("excludeSwitches", ["enable-automation","enable-logging"])
    options.add_experimental_option('prefs', prefs)#加入個人喜好設定

    #進入不動產成交案件實際資訊資料供應系統
    #因為進入實價登錄下載頁面會跳出內政部開放資料平台的網頁所以必須把這個頁面關掉
    #並且把主要的視窗切換成我們要進行下載的頁面
    driver = webdriver.Chrome(options=options)
    driver.get("https://plvr.land.moi.gov.tw/DownloadOpenData") 
    now_handle = driver.current_window_handle 
    #獲取所有視窗控制程式碼
    all_handles = driver.window_handles 
    for handle in all_handles:
        if handle!=now_handle:
            #輸出待選擇的視窗控制程式碼
            print(handle)
            driver.switch_to.window(handle)
            time.sleep(1)
            driver.close() #關閉當前視窗
    #輸出主視窗控制程式碼
    print(now_handle)
    driver.switch_to.window(now_handle) #返回主視窗
    #進入非本期下載
    button1 = driver.find_element(By.ID, "ui-id-2")
    button1.click()
    time.sleep(2)
    #選擇進階下載
    button2 = driver.find_element(By.ID, "downloadTypeId2")
    button2.click()
    time.sleep(2)
    #接收函式傳進來的資料種類(買賣、新成屋、租賃)，在外面用迴圈執行此函式
    button3 = driver.find_element(By.ID, data_type)
    button3.click()
    #選擇不同季度
    selector = Select(driver.find_element_by_name("season"))
    opt_season = selector.options
    for i in range(len(opt_season)):#切換不同季度
        if True:
            wait = ui.WebDriverWait(driver,10)
            wait.until(lambda driver: driver.find_element_by_xpath('//*[@id="historySeason_id"]'))
            year = Select(driver.find_element_by_name("season"))
            year.select_by_index(i)
            #選擇下載檔案格式
            select = Select(driver.find_element_by_name('fileFormat'))
            select.select_by_value("csv")
            #下載按鈕
            element_dl = driver.find_element_by_name("button9")
            driver.execute_script("arguments[0].click();",element_dl)
            time.sleep(20)
        else:
            pass
    button3.click()#取消選取當前資料種類
    #本期下載
    button4 = driver.find_element(By.ID, "ui-id-1")
    button4.click()
    time.sleep(5)
    select = Select(driver.find_element_by_name('fileFormat'))
    select.select_by_value("csv")
    button5 = driver.find_element(By.ID, "downloadTypeId2")
    button5.click()
    button6 = driver.find_element(By.ID, data_type)
    button6.click()
    driver.find_element_by_name("button9").click()


#執行前季與本期下載
def seperate_download():
    download_root_path = path
    evident_type = ['1.不動產買賣','2.預售屋買賣','3.不動產租賃']#欲建立的資料夾名稱
    evident_type_ID = ["typeASelectAll","typeBSelectAll","typeCSelectAll"]#網頁按鈕的id
    #在本機端建立儲存用的資料夾並且執行下載
    for i in range(len(evident_type)):
        if not os.path.exists(download_root_path+'\\'+evident_type[i]):
            os.mkdir(download_root_path+'\\'+evident_type[i])#建資料夾
            land_dl(evident_type[i],evident_type_ID[i])#下載到資料夾
        else:
            land_dl(evident_type[i],evident_type_ID[i])#更新資料用


# 前期下載
def previous_vol():    
    out_path = path+'4.前期下載'  # 下載想指定的路徑
    prefs = {'profile.default_content_settings.popups': 0,
     'download.default_directory': out_path,
     "profile.password_manager_enabled": False,
     "credentials_enable_service": False,
     "profile.default_content_setting_values.automatic_downloads" : 1}
    options = webdriver.ChromeOptions()
    options.add_experimental_option("excludeSwitches", ["enable-automation","enable-logging"])
    options.add_experimental_option('prefs', prefs)
    #進入不動產成交案件實際資訊資料供應系統
    driver = webdriver.Chrome(executable_path=r'chromedriver.exe', options=options)
    driver.get("https://plvr.land.moi.gov.tw/DownloadOpenData")
    now_handle = driver.current_window_handle 
    #獲取所有視窗控制程式碼
    all_handles = driver.window_handles 
    for handle in all_handles:
        if handle!=now_handle:
            #輸出待選擇的視窗控制程式碼
            print(handle)
            driver.switch_to.window(handle)
            time.sleep(1)
            driver.close() #關閉當前視窗
    #輸出主視窗控制程式碼
    print(now_handle)
    driver.switch_to.window(now_handle) #返回主視窗
    time.sleep(5)
    button1 = driver.find_element(By.ID, "ui-id-2")
    button1.click()
    time.sleep(1)
    ####################前期下載
    button5 = driver.find_elements_by_class_name('i_download')
    if button5==[]:
        pass
    else:
        for i in range(len(button5)):
            button5[i].click()
            time.sleep(10)

## 執行前期下載
def pre_vol_download():
    try:
        if not os.path.exists(path+'4.前期下載'):
            os.mkdir(path+'4.前期下載')#建資料夾
            previous_vol()#下載到資料夾
        else:
            previous_vol()#更新資料用
    except:
        print("error")


# 解壓縮並刪除不需要之檔案
def zip_extract_clean_data(download_root_path):
    #依一般數字順序(1,2,3,4...)列出資料夾內zip檔案名稱
    #newlist為每個解壓縮資料夾的名稱
    #old_filename為資料夾內實際資料的名稱(例:download1.zip)
    old_filename=os.listdir(download_root_path)
    old_filename = pd.Series(old_filename)
    old_filename = list(old_filename[old_filename.str.contains('download')])
    old_filename.sort(key=len,reverse=True)
    newlist = []
    for names in old_filename:
        if names.endswith(".zip"):
            newlist.append(names)         
    newlist.sort(key=len)
    for i in range(len(newlist)):#刪掉難處理的符號
        newlist[i] = newlist[i].replace(' ', "")
        newlist[i] = newlist[i].replace('(', "")
        newlist[i] = newlist[i].replace(')', "")
        newlist[i] = newlist[i].replace('.zip', "")
    #解壓縮每個zip檔
    crack_file=[]
    for i in range(len(newlist)):
        extract_path =download_root_path +'/'+ old_filename[i]
        file_path = download_root_path+'/' + newlist[i]
        try:
            with zipfile.ZipFile(extract_path, 'r') as zip_ref:
                zip_ref.extractall(file_path)
                zip_ref.close()#關掉當前zip
        except zipfile.BadZipfile:
            print(file_path)
            crack_file.append(file_path)
                
    #刪除資料夾內用不到的資料(欄位說明等)
    #os.walk : root,dirs,files->回傳檔案路徑
    # root 目前資料夾根路徑
    # dirs 根路徑下的資料夾路徑
    # files 資料夾內的檔案路徑
    for i in range(len(newlist)):
        each_file_path = download_root_path+'/' + newlist[i]
        for files in os.walk(each_file_path):
            path_file_list = list(files)[2]
        #各縣市資料路徑加入real_datalist
        #刪除資料路徑加入not_datalist
        #用try...except...讓程式保持運行，有錯誤就跳過    
        df_list = pd.DataFrame(path_file_list,columns=['Name'])
        real_datalist = df_list[df_list['Name'].str.contains('lvr_land')]
        real_datalist = list(real_datalist['Name'])
        not_datalist = df_list[~df_list['Name'].str.contains('lvr_land')]#波浪鍵反選
        not_datalist = list(not_datalist['Name'])
        for j in range(len(not_datalist)):
            try:
                os.remove(each_file_path+'/'+not_datalist[j])
            except OSError as e:
                print(e)
            else:
                pass
    return (newlist,real_datalist,crack_file)

#執行前季下載
def extract_pre_vol(download_root_path):
    #依數字順序列出資料夾內zip檔案名稱
    old_filename=os.listdir(download_root_path)
    print(old_filename)
    old_filename = pd.Series(old_filename,dtype=pd.StringDtype())
    old_filename = list(old_filename[old_filename.str.contains('opendata')])
    old_filename.sort(key=len)
    newlist = []
    for names in old_filename:
        if names.endswith(".zip"):
            newlist.append(names)         
    newlist.sort(key=len)
    for i in range(len(newlist)):
        newlist[i] = newlist[i].replace(' ', "")
        newlist[i] = newlist[i].replace('(', "")
        newlist[i] = newlist[i].replace(')', "")
        newlist[i] = newlist[i].replace('.zip', "")
    #解壓縮每個zip檔
    pre_crack_file=[]
    for i in range(len(newlist)):
        extract_path =download_root_path+ '/' + old_filename[i]
        file_path = download_root_path+'/' + newlist[i]
        with zipfile.ZipFile(extract_path, 'r') as zip_ref:
            try:
                zip_ref.extractall(file_path)
                zip_ref.close()
            except zipfile.BadZipfile:
                print(file_path)
                pre_crack_file.append(file_path)
    for i in range(len(newlist)):
        each_file_path = download_root_path+'/' + newlist[i]
        for files in os.walk(each_file_path):#取得每季資料夾內所有檔案
            path_file_list = list(files)[2]            
        df_list = pd.DataFrame(path_file_list,columns=['Name'])
        real_datalist = df_list[df_list['Name'].str.contains('.xls')]
        real_datalist = list(real_datalist['Name'])
        not_datalist = df_list[~df_list['Name'].str.contains('.xls')]#波浪鍵反選
        not_datalist = list(not_datalist['Name'])           
        for j in range(len(not_datalist)):
            try:
                os.remove(each_file_path+'/'+not_datalist[j])
            except OSError as e:
                print(e)
            else:
                pass 
    return(newlist,pre_crack_file)


#真正檔案下載
seperate_download()
#取得三類資料夾的路徑
path_main_list = []
for path, dirs, files in os.walk(path):
    evident_type = dirs
    break
for i in range(len(evident_type)):
    path_main_list.append(path+'/'+evident_type[i])
pre_vol_path = path_main_list.pop()    
pre_vol_download()



global newlist
real_datalist=[]
crack_file_list =[]
#執行解壓縮與清除檔案
for i in range(len(path_main_list)):
    newlist,b,crack_file = zip_extract_clean_data(path_main_list[i])
    real_datalist.append(b)
    crack_file_list.append(crack_file)



while any(crack_file_list) ==True:
    print(crack_file_list)
    extracted_file = next(os.walk(path_main_list[1]))[1]
    zip_data = next(os.walk(path_main_list[0]))[2]
    extracted_file_df = pd.DataFrame(extracted_file,columns=['Name'])
    extracted_zip_df = pd.DataFrame(zip_data,columns=['Name'])
    delete_list_zip = extracted_zip_df[extracted_zip_df['Name'].str.contains('zip')]
    delete_list_zip = list(delete_list_zip['Name'])
    delete_list_file = extracted_file_df[extracted_file_df['Name'].str.contains('download')]#用"|"表示"or"                                
    delete_list_file = list(delete_list_file['Name'])
    evident_type =next(os.walk(path))[1]
    all_file_path = []
    for i in range(len(evident_type)):
        all_file_path.append(path+'/'+evident_type[i])
    ###迴圈刪除
    for j in range(len(all_file_path)):
        for i in range(len(delete_list_zip)):
            dirPath_zip = all_file_path[j]+'/'+delete_list_zip[i]
            try:
                os.remove(dirPath_zip)
            except OSError as e:
                print(e)
            else:
                pass
    for j in range(len(all_file_path)):
        for i in range(len(delete_list_file)):
            dirPath_file = all_file_path[j]+'/'+delete_list_file[i]
            try:
                shutil.rmtree(dirPath_file)
            except OSError as e:
                print(e)
            else:
                pass
    seperate_download()
    real_datalist=[]
    #執行解壓縮與清除檔案
    for i in range(len(path_main_list)):
        newlist,b,crack_file = zip_extract_clean_data(path_main_list[i])
        real_datalist.append(b)
    if crack_file==[]:
        break

pre_crack_list=[]
try: 
    pre_vol_filename,pre_crack_file = extract_pre_vol(pre_vol_path)
    pre_crack_list.append(pre_crack_file)
    if pre_vol_filename==[]:#前期下載一季彙整一次，防止前期下載空白
        pass
except:
    pass

while any(pre_crack_list) ==True:
    extracted_file = next(os.walk(pre_vol_path))[1]
    zip_data = next(os.walk(pre_vol_path))[2]
    extracted_file_df = pd.DataFrame(extracted_file,columns=['Name'])
    extracted_zip_df = pd.DataFrame(zip_data,columns=['Name'])
    delete_list_zip = extracted_zip_df[extracted_zip_df['Name'].str.contains('zip')]
    delete_list_zip = list(delete_list_zip['Name'])
    delete_list_file = extracted_file_df[extracted_file_df['Name'].str.contains('opendata')]#用"|"表示"or"                                
    delete_list_file = list(delete_list_file['Name'])
    evident_type =next(os.walk(path))[1]
    all_file_path = []
    for i in range(len(evident_type)):
        all_file_path.append(path+'/'+evident_type[i])

    ###迴圈刪除
    for j in range(len(all_file_path)):
        for i in range(len(delete_list_zip)):
            dirPath_zip = all_file_path[j]+'/'+delete_list_zip[i]
            try:
                os.remove(dirPath_zip)
            except OSError as e:
                print(e)
            else:
                pass
                
    for j in range(len(all_file_path)):
        for i in range(len(delete_list_file)):
            dirPath_file = all_file_path[j]+'/'+delete_list_file[i]
            try:
                shutil.rmtree(dirPath_file)
            except OSError as e:
                print(e)
            else:
                pass
    pre_vol_download()
    #執行解壓縮與清除檔案
    pre_vol_filename=[]
    try: 
        pre_vol_filename,pre_crack_file = extract_pre_vol(pre_vol_path)
        pre_crack_list.append(pre_crack_file)
        if pre_vol_filename==[]:#前期下載一季彙整一次，防止前期下載空白
            pass
    except:
        pass
    if pre_crack_list==[]:
        break



#把所有檔名一律小寫(配合前期下載)
#後續讀取檔案方便，可以用一個list就可以對應前期與前季
#因為前期將各縣市三類(買賣新成屋租賃)資料合併成一個檔案，與前季下載不同
for i in range(len(path_main_list)):
    for j in range(len(newlist)):
        files=os.listdir(path_main_list[i]+'/'+newlist[j])
        for k in range(len(files)):            
            old_name = path_main_list[i]+'/'+newlist[j]+'/'+files[k]
            new_name = path_main_list[i]+'/'+newlist[j]+'/'+files[k].lower()
            os.rename(old_name,new_name)
for i in range(len(path_main_list)):
    for j in range(len(real_datalist[0])):
        try:
            real_datalist[i][j]= real_datalist[i][j].lower()
        except:
            pass

if pre_vol_filename == []:
    pass
else:
    for i in range(len(pre_vol_filename)):
        pre_files=os.listdir(pre_vol_path+'/'+pre_vol_filename[i])
        for k in range(len(pre_files)):            
            pre_old_name = pre_vol_path+'/'+pre_vol_filename[i]+'/'+pre_files[k]
            pre_new_name = pre_vol_path+'/'+pre_vol_filename[i]+'/'+pre_files[k].lower()        
            os.rename(pre_old_name,pre_new_name)


# 依照縣市合併案例
def county_combine(county_filename):#接收縣市名稱
    #買賣租賃新成屋表格欄位不同
    #取得第一筆資料作為欄位模板，並另存成新的dataframe
    col_na_A = pd.read_csv(path_main_list[0]+'/'+newlist[0]+'/'+real_datalist[0][0],engine = "python")
    col_name_A = col_na_A.columns.tolist()
    combine_df_A = pd.DataFrame([],columns=col_name_A)
    
    col_na_B = pd.read_csv(path_main_list[1]+'/'+newlist[0]+'/'+real_datalist[1][0],engine = "python")
    col_name_B = col_na_B.columns.tolist()
    combine_df_B = pd.DataFrame([],columns=col_name_B)
    
    col_na_C = pd.read_csv(path_main_list[2]+'/'+newlist[0]+'/'+real_datalist[2][0],engine = "python")
    col_name_C = col_na_C.columns.tolist()
    combine_df_C = pd.DataFrame([],columns=col_name_C)
        

    for j in range(len(path_main_list)):
        for i in range(len(newlist)):
            try:
                new_df = pd.read_csv(path_main_list[j]+'/'+newlist[i]+'/'+county_filename,engine = "python")
                if '不動產買賣' in path_main_list[j]:#利用路徑中包含不同類別名稱判斷
                    combine_df_A = combine_df_A.append(new_df)
                    #刪除重複欄位(英文欄位)，留下最上面的欄位名稱
                    combine_df_A = combine_df_A[~combine_df_A['鄉鎮市區'].isin(['The villages and towns urban district'])]
                    combine_df_A.index = range(1, len(combine_df_A)+1)

                elif '預售屋買賣'in path_main_list[j]:
                    combine_df_B = combine_df_B.append(new_df)
                    combine_df_B = combine_df_B[~combine_df_B['鄉鎮市區'].isin(['The villages and towns urban district'])]
                    combine_df_B.index = range(1, len(combine_df_B)+1)

                elif '不動產租賃' in path_main_list[j]:
                    combine_df_C = combine_df_C.append(new_df)
                    combine_df_C = combine_df_C[~combine_df_C['鄉鎮市區'].isin(['The villages and towns urban district'])]
                    combine_df_C.index = range(1, len(combine_df_C)+1)
            except:
                pass
    return combine_df_A,combine_df_B,combine_df_C

#根據各縣市英文代碼(身分證字號)排列，後續可以改進
countyname = ['臺北市','臺中市','基隆市','臺南市','高雄市','新北市','宜蘭縣','桃園市','嘉義市','新竹縣','苗栗縣','南投縣','彰化縣','新竹市','雲林縣','嘉義縣','屏東縣','花蓮縣','臺東縣','金門縣','澎湖縣','連江縣']

for j in range(len(path_main_list)):
    for i in range(len(real_datalist[0])):
        try:
            #合併
            county_evident_A,county_evident_B,county_evident_C = county_combine(real_datalist[j][i])
        except:
            pass
        if not os.path.exists(path_main_list[j]+'/'+countyname[i]):#如果路徑不存在就建立資料夾
            os.mkdir(path_main_list[j]+'/'+countyname[i])
            if county_evident_A.size!=0:
                try:
                    A_save = county_evident_A.to_csv(path_main_list[0]+'/'+countyname[i]+'/'+countyname[i]+".csv",encoding='utf-8',index=False)
                except:
                    pass
            elif county_evident_B.size!=0:
                try:
                    B_save = county_evident_B.to_csv(path_main_list[1]+'/'+countyname[i]+'/'+countyname[i]+".csv",encoding='utf-8',index=False)
                except:
                    pass
            elif county_evident_C.size!=0:
                try:
                    C_save = county_evident_C.to_csv(path_main_list[2]+'/'+countyname[i]+'/'+countyname[i]+".csv",encoding='utf-8',index=False)
                except:
                    pass
        else:
            #資料夾存在就直接存檔覆蓋
            if county_evident_A.size!=0:
                try:
                    A_save = county_evident_A.to_csv(path_main_list[0]+'/'+countyname[i]+'/'+countyname[i]+".csv",encoding='utf-8',index=False)
                except:
                    pass
            elif county_evident_B.size!=0:
                try:
                    B_save = county_evident_B.to_csv(path_main_list[1]+'/'+countyname[i]+'/'+countyname[i]+".csv",encoding='utf-8',index=False)
                except:
                    pass
            elif county_evident_C.size!=0:
                try:
                    C_save = county_evident_C.to_csv(path_main_list[2]+'/'+countyname[i]+'/'+countyname[i]+".csv",encoding='utf-8',index=False)
                except:
                    pass
#前期下載合併
def vol_combine(county_num):
    real_datalist_name=[]
    for i in range(len(real_datalist)):
        for j in range(len(real_datalist[0])):
            try:
                real_datalist_name.append(re.sub('.csv', '.xls', real_datalist[i][j]))
            except:
                pass
    #把資料夾內的csv檔案依照前季相同的邏輯，3類22縣市的結構組合
    
    real_datalist_name = np.reshape(real_datalist_name, (3,22))
    real_datalist_name = list(real_datalist_name)

    ############
    col_na_A = pd.read_csv(path_main_list[0]+'/'+newlist[0]+'/'+real_datalist[0][0],engine = "python")
    col_name_A = col_na_A.columns.tolist()
    pre_combine_df_A = pd.DataFrame([],columns=col_name_A)

    col_na_B = pd.read_csv(path_main_list[1]+'/'+newlist[0]+'/'+real_datalist[1][0],engine = "python")
    col_name_B = col_na_B.columns.tolist()
    pre_combine_df_B = pd.DataFrame([],columns=col_name_B)

    col_na_C = pd.read_csv(path_main_list[2]+'/'+newlist[0]+'/'+real_datalist[2][0],engine = "python")
    col_name_C = col_na_C.columns.tolist()
    pre_combine_df_C = pd.DataFrame([],columns=col_name_C)
    for i in range(len(real_datalist_name)):#3
            for j in range(len(pre_vol_filename)):#期數
                if i==0:#買賣 #(pip install xlrd)
                    try:
                        pre_vol_county = pd.read_excel(pre_vol_path+'/'+pre_vol_filename[j]+'/'+real_datalist_name[i][county_num],sheet_name='不動產買賣')
                        pre_combine_df_A = pre_combine_df_A.append(pre_vol_county)
                        pre_combine_df_A = pre_combine_df_A[~pre_combine_df_A['鄉鎮市區'].isin(['The villages and towns urban district'])]
                        pre_combine_df_A.index = range(1, len(pre_combine_df_A)+1)
                    except:
                        pass
                elif i==1:#預售屋
                    try:
                        pre_vol_county = pd.read_excel(pre_vol_path+'/'+pre_vol_filename[j]+'/'+real_datalist_name[i][county_num],sheet_name='預售屋買賣')
                        pre_combine_df_B = pre_combine_df_B.append(pre_vol_county)
                        pre_combine_df_B = pre_combine_df_B[~pre_combine_df_B['鄉鎮市區'].isin(['The villages and towns urban district'])]
                        pre_combine_df_B.index = range(1, len(pre_combine_df_B)+1)
                    except:
                        pass
                elif i==2:#租賃
                    try:
                        pre_vol_county = pd.read_excel(pre_vol_path+'/'+pre_vol_filename[j]+'/'+real_datalist_name[i][county_num],sheet_name='不動產租賃')
                        pre_combine_df_C = pre_combine_df_C.append(pre_vol_county)
                        pre_combine_df_C = pre_combine_df_C[~pre_combine_df_C['鄉鎮市區'].isin(['The villages and towns urban district'])]
                        pre_combine_df_C.index = range(1, len(pre_combine_df_C)+1)
                    except:
                        pass
    return pre_combine_df_A,pre_combine_df_B,pre_combine_df_C
#前期存檔
if pre_vol_filename == []:
    pass
else:
    for i in range(len(countyname)):
        pre_buysell,pre_newsell,pre_rent =vol_combine(i)
        pre_buysell.to_csv(path+'\\'+'1.不動產買賣\\'+countyname[i]+'\\'+countyname[i]+'.csv',encoding='utf-8',index=False,header=False,mode='a+')
        pre_newsell.to_csv(path+'\\'+'2.預售屋買賣\\'+countyname[i]+'\\'+countyname[i]+'.csv',encoding='utf-8',index=False,header=False,mode='a+')
        pre_rent.to_csv(path+'\\'+'3.不動產租賃\\'+countyname[i]+'\\'+countyname[i]+'.csv',encoding='utf-8',index=False,header=False,mode='a+')

# 重新讀取合併後檔案，依照交易/租賃日期篩選新->舊
# for i in range(len(countyname)):
#     com_A = pd.read_csv(path+'\\'+'1.不動產買賣\\'+countyname[i]+'\\'+countyname[i]+'.csv',encoding='utf-8')
#     com_A = com_A.sort_values(by=['交易年月日'],ignore_index=False)
#     com_A.to_csv(path+'\\'+'1.不動產買賣\\'+countyname[i]+'\\'+countyname[i]+'.csv',encoding='utf-8',index=False)
#     com_B = pd.read_csv(path+'\\'+'2.預售屋買賣\\'+countyname[i]+'\\'+countyname[i]+'.csv',encoding='utf-8')
#     com_B = com_B.sort_values(by=['交易年月日'],ignore_index=False)
#     com_B.to_csv(path+'\\'+'2.預售屋買賣\\'+countyname[i]+'\\'+countyname[i]+'.csv',encoding='utf-8',index=False)
#     com_C = pd.read_csv(path+'\\'+'3.不動產租賃\\'+countyname[i]+'\\'+countyname[i]+'.csv',encoding='utf-8')
#     com_C = com_C.sort_values(by=['租賃年月日'],ignore_index=True)
#     com_C.to_csv(path+'\\'+'3.不動產租賃\\'+countyname[i]+'\\'+countyname[i]+'.csv',encoding='utf-8',index=False)

#取得刪除檔案名稱，刪除下載下來的壓縮檔以及解壓縮後的資料夾
extracted_file = next(os.walk(path_main_list[1]))[1]+next(os.walk(pre_vol_path))[1]
zip_data = next(os.walk(path_main_list[0]))[2]+next(os.walk(pre_vol_path))[2]
extracted_file_df = pd.DataFrame(extracted_file,columns=['Name'])
extracted_zip_df = pd.DataFrame(zip_data,columns=['Name'])
delete_list_zip = extracted_zip_df[extracted_zip_df['Name'].str.contains('zip')]
delete_list_zip = list(delete_list_zip['Name'])
delete_list_file = extracted_file_df[extracted_file_df['Name'].str.contains('download|opendata')]#用"|"表示"or"                                
delete_list_file = list(delete_list_file['Name'])
evident_type =next(os.walk(path))[1]
all_file_path = []
for i in range(len(evident_type)):
    all_file_path.append(path+'/'+evident_type[i])

##迴圈刪除
for j in range(len(all_file_path)):
    for i in range(len(delete_list_zip)):
        dirPath_zip = all_file_path[j]+'/'+delete_list_zip[i]
        try:
            os.remove(dirPath_zip)
        except OSError as e:
            print(e)
        else:
            pass
            
for j in range(len(all_file_path)):
    for i in range(len(delete_list_file)):
        dirPath_file = all_file_path[j]+'/'+delete_list_file[i]
        try:
            shutil.rmtree(dirPath_file)
        except OSError as e:
            print(e)
        else:
            pass