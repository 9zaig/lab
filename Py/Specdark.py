import pandas as pd
import numpy as np
doc=pd.read_excel('./Input/통합 문서2.xlsx')
# print(doc)

####################################
#lowest peak intensity
filtered_values = doc.iloc[:, 12].apply(pd.to_numeric, errors='coerce').dropna()
filtered_values = filtered_values[filtered_values != 0]
min_value = np.min(filtered_values)
point=min_value
####################################
df1=pd.DataFrame(doc.iloc[:,12].dropna().iloc[1:]) #50*1
# print(df1)
data_org=pd.DataFrame(doc.iloc[:,19:]) #3195*50
print(data_org)
data = pd.DataFrame()

################################
#dark 위치 찾기
count_zero = (df1 == 0).sum()
# print("0의 개수:", count_zero)
dark=data_org.iloc[:,count_zero-1]
# print(dark)
################################

for col in doc.columns[25:75]:

    column_data= doc[col]


    # column_data = df.iloc[2:,col]
    # print(column_data)
    max_value = column_data.max()


    if max_value > point:
        data[col] = column_data
#max 최대값 넘는지 안넘는지
###########################################

# print(result_df)
counts = data.iloc[1].value_counts()
counts = counts.sort_index(ascending=False)
# print(counts)
counts_all=np.cumsum(counts)
# print(counts_all)
# print(counts.iloc[0])
df=pd.DataFrame()
wave=pd.DataFrame(doc.iloc[2:,24])
df['Wavelength']=wave

# 결과 출력
# # print(counts.index)
#두번째 숫자들 모음
###########################################
for column in data.columns:
    if data[column].iloc[1]==counts.index[0]:
        # print(data[column])
        # print(dark.iloc[2:,0])
        datt=data[column].iloc[2:]-dark.iloc[2:,0]
        df[data[column].iloc[0]]=datt

    elif data[column].iloc[1]==counts.index[1]:
        # print(data.iloc[:,counts_all.iloc[0]])
        # print(df.iloc[:,counts_all.iloc[0]-1])
        datt_21=data[column].iloc[2:]-df.iloc[:,counts_all.iloc[0]-1]
        # print(datt_2)
        mean_values = datt_21[:50].mean()
        # print(mean_values)
        datt_22=data[column].iloc[2:]-mean_values
        # print(datt_22)
        df[data[column].iloc[0]]=datt_22
    for n in range(2,len(counts)):
        # print(n)
        if data[column].iloc[1]==counts.index[n]:
            # print(data.iloc[:,counts_all.iloc[n-1]-1])
            # print(data[column].iloc[2:])
            # print(counts_all.iloc[n - 1]-1)
            # print(df.iloc[:,counts_all.iloc[n - 1]-1])
            datt_31 = data[column].iloc[2:] - df.iloc[:, counts_all.iloc[n-1] - 1]
            # print(datt_31)
            mean_values = datt_31[:50].mean()
            # print(mean_values)
            datt_32 = data[column].iloc[2:] - mean_values
            # print(datt_32)
            df[data[column].iloc[0]] = datt_32

# print(df)

df.iloc[:, :] = df.iloc[:, :].applymap(lambda x: 0.5 if x <= 0 else x)

print(df)
df.to_excel('final.xlsx')