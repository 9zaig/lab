import pandas as pd
import openpyxl
from openpyxl.utils.dataframe import dataframe_to_rows
import numpy as np
input='./In/specdata.xlsx'
import warnings
warnings.filterwarnings("ignore")
# output_file ="specout.xlsx"
doc=pd.read_excel(input)
# print(doc)
cheapsize=doc.iloc[5,1]
# print(cheapsize)
seca=doc.iloc[:,6:17]
# print(seca)
secb=doc.iloc[:,24:]
integration_time=secb.iloc[1,:]
# print(integration_time)
secba=secb.drop(1)

new_header=secba.iloc[0]
secba=secba[1:]
secba.columns=new_header
# print(secba)
wave=secba.iloc[:,0]
numwave=len(wave)
# print(numwave)
photonE=1240/wave
# print(photonE)
photondf=pd.DataFrame(photonE).reset_index()
# print(wave)
data=secba.iloc[:,1:]
print(data)
# print(data.columns)
# print(data.idxmax())

ans=pd.DataFrame()
# print(ans)
order=1
# print(meanlist)
for column in data.columns:
    # print(order)
    num=int(data[column].idxmax())-2
    peak=wave.iloc[num]
    # print(peak)
    hu=1240/peak
    # print(hu)
    peak_intensity=0
    if peak > 400:
        peak_intensity=data[column].max()
    # print(data[column].iloc[256])
    integ=''
    if data[column].iloc[256]>0:
        integ=integration_time[order]
        # print(integ)
    if integ>0:
        current=data.columns.to_list()[order-1]
        # print(current)
        currentdensity=current/cheapsize
    df=pd.DataFrame([current,currentdensity,0,integ,hu,peak,peak_intensity]).T
    # print(df)
    ans=ans.append(df)
    # pd.concat([ans,df],axis=0)
    # print(hu)
    # print(peak)
    # print(peak_intensity)
    order = order + 1

# print(photondf)
# print(ans)
# print(ans.info)
workbook = openpyxl.load_workbook(input)

# 워크시트 선택 (예: 첫 번째 워크시트 선택)
worksheet = workbook['Sheet1']
for r_idx, row in enumerate(ans.values, 6):  # 6행부터 시작 (G6 셀)
    for c_idx, value in enumerate(row, 7):  # G열부터 시작
        cell = worksheet.cell(row=r_idx, column=c_idx, value=value)

# 변경 사항을 엑셀 파일에 저장
workbook.save('./out/out.xlsx')

# 엑셀 파일 닫기
workbook.close()
