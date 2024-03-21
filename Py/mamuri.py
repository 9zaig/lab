import pandas as pd
import os

loc=''
files=os.listdir(loc)
file_names=[]
for file in files:
    file_path = os.path.join(loc, file)
    if os.path.isfile(file_path):
        file_names.append(int(file[:-7]))

file_names.sort()
mamuri=pd.DataFrame()
for file_name in file_names:
    file_name=str(file_name)+'mW.xlsx'
    a=file_name[:-7]
    df = pd.read_excel(loc + '/'+file_name)
    int=df.iloc[1,1]
    ans = pd.DataFrame([a,int])

    dan=df.iloc[5:,1]
    dan=dan/int
    concat1=pd.concat([ans,dan])

    mamuri=pd.concat([mamuri,concat1],axis=1)

mamuri.to_excel('mamuri.xlsx')