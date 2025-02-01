import bom_check as bc
import os 
import re

folder = os.getcwd()


def find_list(folder):
    li = os.walk(folder)
    f_li = list(li)[0][2]
    
    file = re.compile('(.*)_SW.xlsx')
    
    result = []
    for i in f_li:
        if file.findall(i) != []:
            result.append(file.findall(i)[0])
    return result

print('程式開啟完成')
input('按下Enter開始執行')

a = True
while a:
    print('-開始執行-')
    print(f'搜尋路徑{folder}中,Solidwork零件表xlsx檔...')

    li = find_list(folder)


    for f in li:
        try:
            if os.path.exists(f'{folder}\\{f}_ERP.xlsx') == True:
                bc.bom_check(folder,f'{f}_SW',f'{f}_ERP',f'{f}_check')
                print(f'{f} Solidwork 零件表 與 ERP-BOM表 比對完成')
                print(f'\t-->{f}_check.xlsx\n')
            elif os.path.exists(f'{folder}\\{f}_SW.xlsx') == True:
                bc.bom_check(folder,f'{f}_SW',None,f'{f}_part_name_check')
                print(f'{f} Solidwork 零件編號 檢查完成 ')
                print(f'\t{f}_part_name_check.xlsx\n')
            else:
                print(f'{f} 無檢查 檔案讀取錯誤\n')
        except Exception as err:
                print(f'{f} {err}\n')
    print('-執行完成-')
    print('可打開產生之check.xlsx確認檢查結果')
    b = input('修正完畢後，如需再次檢查請按下Enter。結束請輸入q')
    if b == 'q':
        a = False
    

    

   