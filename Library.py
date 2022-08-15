import pandas as pd
import numpy as np
import json
import xlwt 
from xlwt import Workbook 
import xlsxwriter

workbook = xlsxwriter.Workbook('library_dynamic.xlsx') 


f = open ('TAG_PARAM.json', "r") 
  
# Reading from file 
data = json.loads(f.read())

#print (data)

col_list=["test_steps"]
#data = pd.read_excel(r'.\datasheet.xlsx')
data1=pd.read_csv(r'.\new_test_refactored_test.csv',usecols=col_list)
#df = pd.DataFrame(data)
df1=pd.DataFrame(data1)
#n0 = df.shape
n1= df1.shape
#print(n1)
#print(df1["test_steps"])
#print(n0) 
arr1 = df1.to_numpy()
#print(arr1[0])
pillers=[[]]
for i in range (0,len(arr1),1):
    #print(type(arr1[i]))
    #print(i)
    #print("------------------------******************-------------")
    try:
        str_arr=str(arr1[i])
    
        #print(str_arr)
        str_arr1=str(str_arr[2:-2])
        #print((str_arr1))
        '''print("@@@@##$")
        #print(str_arr)
        #str_arr1=json.dumps(str_arr1)
        #dic = eval(str_arr)
        try:
            dic=json.loads(str_arr1)
        except Exception as err:
            print(err)
       
        #step=json.dumps(dic)
        print(type(dic))
        #print(type(step))
        print(type(str_arr1))
        for steps in dic:
            print(steps["tag"])'''
        #json_str=str(json.dumps(str_arr1))
        test_step_obj = eval(str_arr1)
        
        n=1
        piller_row=[]
        #pillers.append(piller_row)
        for steps in test_step_obj:
            exist=0
            flag=0
            print("Name => " , steps["tag"])
            #print(len(pillers))
            #print("233333333333333323232323232#@#@#2323232\n")
            for row in range(0,len(pillers),1):
                for col in range(0,len(pillers[row]),1):
                    if steps["tag"]==pillers[row][col]:
                        exist=1
                        #print("exist-----------")
            if exist==0:
                n=len(pillers)
                for row in range(0,len(pillers),1):
                    if (str(steps["tag"])[:5]=="START"):
                        pillers[0].append(steps["tag"])
                        flag=1
                        break
                    elif (str(steps["tag"])[:4]=="TEAR"):
                        pillers[n-1].append(steps["tag"])
                        flag=1
                        break
                        
                    elif(str(steps["tag"])[:3]==str(pillers[row][0])[:3] and not row==n-1):
                        pillers[row].append(steps["tag"])
                        flag=1
                        break
                        
                        
                if flag==0:
                    new_piller=[]
                    new_piller.append(steps["tag"])
                    pillers.append(new_piller)
                    flag=0
            
        
            
            
            
            
    except Exception as err:
        print(err)
        #print(exc_info())


# Workbook is created 
wb = Workbook() 
  
# add_sheet is used to create sheet. 
sheet2 = workbook.add_worksheet('Sheet1')
sheet1 = workbook.add_worksheet('Sheet2') 
name=[[]]
max_col,max_colm=1,1
for row in range(0,len(pillers),1):
    name_ar=[]
    name.append(name_ar)
    if(max_colm<len(pillers[row])):
        max_colm=len(pillers[row])
    #print(len(pillers[row]))
    for col in range(0,len(pillers[row]),1):
        sheet1.write(col+1,row,pillers[row][col])
        flag=0
        for i in data['Sheet1']:
            #print(i["tag_name"][:-4])
            if str(i["tag_name"][:-5])==pillers[row][col]:
                flag=1
                name[row].append(i["tag_short_unique__name"])
                sheet1.write(0,row,i["tag_group"])
                sheet2.write(0,row,i["tag_group"])
                sheet2.write(col+1,row,i["tag_short_unique__name"])
                break
        if flag==0:
            name[row].append(pillers[row][col])
            sheet1.write(0,row,pillers[row][col])
            sheet2.write(0,row,pillers[row][col])
            sheet2.write(col+1,row,pillers[row][col])
                #print(i["tag_short_unique__name"])
    #print (pillers[row])
    #print("\n")
print (max_colm)
for row in range(0,len(name),1):
    for col in range(len(name[row]),max_colm,1):
        sheet1.write(col+1,row,"N")
        sheet2.write(col+1,row,"N")

#wb.save('library_new.xlsx')

workbook.close()

print(name)
    
print((len(pillers)))


