import tkinter as tk
from tkinter.filedialog import askdirectory
import bom_check as bc
import re,os

#選擇資料夾
def select_filepath(*arg):
    folder = askdirectory()
    
    filepath.set(folder)

    li_list.delete(0,tk.END)
    for name in find_list(folder):
        li_list.insert(tk.END,name)

#尋找檔案名稱及加入清單
def find_list(folder):
    li = os.walk(folder)
    f_li = list(li)[0][2]
    
    file = re.compile('(.*)_SW.xlsx')
    
    result = []
    for i in f_li:
        if file.findall(i) != []:
            result.append(file.findall(i)[0])
    return result

#進行檢查
def start_check():
    li = li_list.get(0,tk.END)
    folder = filepath.get()
    part_path = f'{os.getcwd()}\\part.csv'
    print(part_path)
    li_out.delete(0,tk.END)
    
    if len(li)==0:
        li_out.insert(tk.END,'(無檢查項目)')
    else:
        li_out.insert(tk.END,'-----檢查開始-----')
    
    for f in li:
        
        try:
            if os.path.exists(f'{folder}\\{f}_ERP.xlsx') == True:
                bc.bom_check(folder,f'{f}_SW',f'{f}_ERP',part_path,f'{f}_check')
                li_out.insert(tk.END,(f'{f} BOM表 比對完成'))
                li_out.insert(tk.END,(f'    -->{f}_check.xlsx\n'))
            elif os.path.exists(f'{folder}\\{f}_SW.xlsx') == True:
                bc.bom_check(folder,f'{f}_SW',None,part_path,f'{f}_part_name_check')
                li_out.insert(tk.END,(f'{f}  零件編號 檢查完成 '))
                li_out.insert(tk.END,(f'    -->{f}_part_name_check.xlsx\n'))
            else:
                li_out.insert(tk.END,(f'{f} 無檢查 檔案讀取錯誤\n'))
        except Exception as err:
                print(err)
                li_out.insert(tk.END,(f'{f} {err}\n'))
    else:    
        li_out.insert(tk.END,'-----檢查結束-----')

#清單交換
def change_list(*arg):
    indexs = li_list.curselection()
    if len(indexs) != 0:
        xli_list.insert(tk.END,li_list.get(indexs))
        li_list.delete(indexs)

def change_xlist(*arg):
    indexs = xli_list.curselection()
    if len(indexs) != 0:
        li_list.insert(tk.END,xli_list.get(indexs))
        xli_list.delete(indexs)

#------GUI開頭-----

root = tk.Tk()
root.title('Bom Check 1.1')
root.resizable(0,0)
root.iconbitmap('D:\\python_code\\bom_check\\2_061.ico')

font_style = '標楷體 14'
filepath = tk.StringVar()

la_filepath = tk.Label(root,textvariable=filepath,font = font_style,width=40,anchor='e')
la_check = tk.Label(root,text='檢查清單',font=font_style,anchor='w')
la_xcheck = tk.Label(root,text='不檢查清單',font=font_style,anchor='w')

bu_filepath = tk.Button(root,text='選擇資料夾',command=select_filepath,font = font_style,width=15)

#組件清單、不檢查組件清單
li_list = tk.Listbox(root,font=font_style,width=28,height=10)
xli_list = tk.Listbox(root,font=font_style,width=28,height=10)

bu_check = tk.Button(root,text='檢查',command=start_check,font=font_style,width=10)
li_out = tk.Listbox(root,relief='flat',font=font_style,width=60,height=10)

li_list.bind("<Double-Button-1>",change_list)
xli_list.bind("<Double-Button-1>",change_xlist)

#GUI 輸出
la_filepath.grid(row=0,column=0,padx=5,pady=5,columnspan=3,sticky=tk.W)
bu_filepath.grid(row=0,column=3,padx=5,pady=5)
la_check.grid(row=1,column=0,padx=5,pady=5,columnspan=2)
la_xcheck.grid(row=1,column=2,padx=5,pady=5,columnspan=2)
li_list.grid(row=2,column=0,padx=5,pady=5,columnspan=2)
xli_list.grid(row=2,column=2,padx=5,pady=5,columnspan=2)
bu_check.grid(row=3,column=0,padx=5,pady=5,columnspan=2)
li_out.grid(row=4,column=0,padx=5,pady=5,columnspan=4)

root.mainloop()