import pandas as pd
import re
import os

#----------副函式----------

#尋找標題函式
def get_title(title_re,title_list):
    t_re = re.compile(title_re)
    for i in title_list:
        if t_re.findall(i) != []:
            return i
            break


#SW零件表與ERP-BOM比對函式(主函式)

def bom_check(folder,sw_file,erp_file,part_path="",out="out"):
    print(part_path)
    sw_path = f"{folder}\\{sw_file}.xlsx"
    df_sw1 = pd.read_excel(sw_path,sheet_name='工作表1',header=0)

    if erp_file != None:
        erp_path = f"{folder}\\{erp_file}.xlsx"
        df_erp1 = pd.read_excel(erp_path,sheet_name=0,header=1)
    else:
        df_erp1 = pd.DataFrame({'子件品號':['None'],'品名規格':['None'],'標準用量':['None'],'選用':['']})

    #排除市購件

    
    if os.path.isfile(part_path):
        print('有找到')
        df_part = pd.read_csv(part_path,header=None)
        df_part['使用'] = True
        df_part.columns = ['零件編號','使用']
        df_sw_part = pd.merge(df_sw1,df_part,how='outer')
        df_sw1 = df_sw_part[df_sw_part['使用'].isnull()].reset_index(drop=True)
    else:
        print(f'{part_path}\\part.csv 未找到市構件參考檔')



    #SW 資料更換欄位名

    try:            
        df_sw = df_sw1[(['零件編號','SW-檔案名稱(File Name)','Description','數量'])]
        df_sw = df_sw.rename(columns={'SW-檔案名稱(File Name)':'檔案名稱','數量':'SW-數量'})
    except Exception as er:
        
        raise Exception('SolidWorks 零件表標題錯誤')
        
    #過濾ERP不選用(留下Nan = True)
    df_erp1 = df_erp1[df_erp1['選用'].isnull()].reset_index(drop=True)
    
    #ERP資料更換欄位名
    
    title_list = df_erp1.columns
    item_no = get_title(r'子\s*件\s*品\s*號',title_list) #子件品號
    item_name = get_title(r'品\s*名\s*規\s*格',title_list) #品名規格
    item_vol = get_title(r'標\s*準\s*用\s*量',title_list) #標準用量
    
    erp_title = [item_no,item_name,item_vol]
    if None in erp_title:
        raise Exception('ERP BOM標題錯誤')

    df_erp = df_erp1[(erp_title)]
    df_erp = df_erp.rename(columns = {item_no:'料號',item_name:'品名',item_vol:'ERP用量'})


    #中文檔名判定函式
    def find_chinese(string):
        out = False
        for s in string:
            if ord(s) > 256 or ord(s) <= 32:
                out = True
                break
        return out

    df_sw_check = df_sw

    #SW檔案名稱檢查
    df_sw_check.loc[:,'檔案名稱檢查'] = df_sw['零件編號'] != df_sw['檔案名稱']


    #零件編號內含中文檢查
    for i in range(df_sw['檔案名稱'].size):
        
        try:
            df_sw_check.loc[i,'中文檔名檢查'] = find_chinese(df_sw['檔案名稱'][i])
        except:
            print(df_sw)




    #聚合ERP及SW(聯集)
    df_sw = df_sw.rename(columns={'零件編號':'料號'})
    df_sw_erp = pd.merge(df_erp,df_sw,left_on='料號',right_on='料號',how='outer')
    df_sw_erp = df_sw_erp.loc[:,['料號','品名','Description','ERP用量','SW-數量']]

    #僅存在於SW資料
    only_sw = df_sw_erp[df_sw_erp['ERP用量'].isnull()].loc[:,['料號','Description','SW-數量']].reset_index(drop=True)
    only_sw.sort_values(by='料號',inplace=True)
    #僅存在於ERP資料
    only_erp = df_sw_erp[df_sw_erp['SW-數量'].isnull()].loc[:,['料號','品名','ERP用量']].reset_index(drop=True)
    only_erp.sort_values(by='料號',inplace=True)
    #聚合ERP及SW(交集)
    df_sw_erp_in = pd.merge(df_erp,df_sw,left_on='料號',right_on='料號',how='inner')
    df_sw_erp_in = df_sw_erp_in.loc[:,['料號','品名','ERP用量','Description','SW-數量']]

    df_sw_erp_in
    #ERP及SW 品名檢查
    df_sw_erp_in.loc[:,'零件品名檢查'] = df_sw_erp_in['品名'] != df_sw_erp_in['Description']
    df_sw_erp_in

    #ERP及SW 數量檢查
    df_sw_erp_in.loc[:,'零件數量檢查'] = df_sw_erp_in['ERP用量'] != df_sw_erp_in['SW-數量']


    try:
        with pd.ExcelWriter(f'{folder}\\{out}.xlsx') as writer:

            #輸出SW及ERP聚合(聯集)
            df_sw_erp.to_excel(writer,sheet_name='SW×ERP',index=False)
            
            #輸出檔案名稱與零件編號不符
            df_sw_check[df_sw_check['檔案名稱檢查']==True].loc[:,['零件編號','檔案名稱','Description']].\
            to_excel(writer,sheet_name='零件編號檢查',index=False)

            #輸出檔案名稱內含中文或空白
            df_sw_check[df_sw_check['中文檔名檢查']==True].loc[:,['零件編號','檔案名稱','Description']].\
            to_excel(writer,sheet_name='中文檔名檢查',index=False)

            #輸出僅出現於ERP及SW
            only_sw.to_excel(writer,sheet_name='Only_Sw & ERP',index=False,header=['SW-零件編號','Description','SW-數量'])
            only_erp.to_excel(writer,sheet_name='Only_Sw & ERP',index=False,header=['ERP-料號','零件名稱','ERP數量'],startcol=4)

            #輸出SW與ERP品名不符
            df_sw_erp_in[df_sw_erp_in['零件品名檢查']==True].loc[:,['料號','品名','Description']].\
            to_excel(writer,sheet_name='零件品名檢查',index=False,header=['零件編號','ERP-零件名稱','SW-Description'])

            #輸出SW與ERP數目不符
            df_sw_erp_in[df_sw_erp_in['零件數量檢查']==True].loc[:,['料號','品名','SW-數量','ERP用量']].\
            to_excel(writer,sheet_name='零件數量檢查',index=False,header=['零件編號','品名','SW-數量','ERP用量'])
    except:
        raise Exception(f'輸出檔產生錯誤，可能開啟中。請關閉{sw_file[:-3]}..._chcek.xlsx檔')
