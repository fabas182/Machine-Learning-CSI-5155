import os
import numpy as np
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile

import tkinter as tk
from tkinter import filedialog
import pandas as pd

root= tk.Tk()

canvas1 = tk.Canvas(root, width = 300, height = 300, bg = 'lightsteelblue')
canvas1.pack()

def getExcel ():
    global df
    filePath = filedialog.askopenfilename()
    cols=[0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26]
    df = pd.read_excel(filePath,usecols=cols,header=0,dtype={0: float,1 : float, 2: float, 3:float, 4:float, 5: float, 6: float})
    return df

browseButton_Excel = tk.Button(text='Import Excel File', command=getExcel, bg='green', fg='white', font=('helvetica', 12, 'bold'))
canvas1.create_window(150, 150, window=browseButton_Excel)
root.mainloop()
'replace empty cells with NaN'

df.replace(' ', np.nan, inplace=True)

r, c = df.shape    
print (df)  


'''creating a super Matrix to contain all the data"'''
superMatrix=np.zeros((r,c))
for i in range (0,r):
    for j in range (0,c):
        if j==0:
            superMatrix[i][j]=df.loc[i,"a"]            
        if j==1:
            superMatrix[i][j]=df.loc[i,"b"]
        if j==2:
            superMatrix[i][j]=df.loc[i,"c"]
        if j==3:
            superMatrix[i][j]=df.loc[i,"d"]
        if j==4:
            superMatrix[i][j]=df.loc[i,"e"]
        if j==5:
            superMatrix[i][j]=df.loc[i,"f"]
        if j==6:
            superMatrix[i][j]=df.loc[i,"g"]
        if j==7:
            superMatrix[i][j]=df.loc[i,"h"]
        if j==8:
            superMatrix[i][j]=df.loc[i,"i"]
        if j==9:
            superMatrix[i][j]=df.loc[i,"j"]
        if j==10:
            superMatrix[i][j]=df.loc[i,"k"]
        if j==11:
            superMatrix[i][j]=df.loc[i,"l"]
        if j==12:
            superMatrix[i][j]=df.loc[i,"m"]
        if j==13:
            superMatrix[i][j]=df.loc[i,"n"]
        if j==14:
            superMatrix[i][j]=df.loc[i,"o"]
        if j==15:
            superMatrix[i][j]=df.loc[i,"p"]
        if j==16:
            superMatrix[i][j]=df.loc[i,"q"]
        if j==17:
            superMatrix[i][j]=df.loc[i,"r"]
        if j==18:
            superMatrix[i][j]=df.loc[i,"s"]
        if j==19:
            superMatrix[i][j]=df.loc[i,"t"]
        if j==20:
            superMatrix[i][j]=df.loc[i,"u"]
        if j==21:
            superMatrix[i][j]=df.loc[i,"v"]
        if j==22:
            superMatrix[i][j]=df.loc[i,"w"]
        if j==23:
            superMatrix[i][j]=df.loc[i,"x"]
        if j==24:
            superMatrix[i][j]=df.loc[i,"y"]
        if j==25:
            superMatrix[i][j]=df.loc[i,"z"]
        if j==26:
            superMatrix[i][j]=df.loc[i,"zz"]

superMatrix.astype(float)    
np.round_(superMatrix,decimals=2,out=None)

totalSetCycles=superMatrix[r-1][1]
print("\n",totalSetCycles)

for i in range (r-1,1,-1):
   
    if superMatrix[i][0]==superMatrix[i-1][0]:
        superMatrix[i][26]=totalSetCycles-superMatrix[i][1]  
    else:
        superMatrix[i][26]=totalSetCycles-superMatrix[i][1]
        totalSetCycles=superMatrix[i-1][1]

superMatrix[1][26]=superMatrix[2][26]+1
superMatrix[0][26]=superMatrix[1][26]+1

for i in range (0,r):
    if superMatrix[i][26]>50:
        superMatrix[i][26]=-1
    elif superMatrix[i][26]<=50 and superMatrix[i][26]>5:
        superMatrix[i][26]=0
    elif superMatrix[i][26]<=5:
        superMatrix[i][26]=1


df=pd.DataFrame(superMatrix)
filepath='dataset4with3classes.xlsx'
df.to_excel(filepath,index=False)
