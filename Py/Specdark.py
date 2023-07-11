import pandas as pd
import numpy as np
input='통합 문서2.xlsx'
input_file = "./Input/"+input
output_file = "./Output/"+input+"_out.xlsx"
doc=pd.read_excel(input_file)
# print(doc)

####################################
#lowest peak intensity
filtered_values = doc.iloc[:, 12].apply(pd.to_numeric, errors='coerce').dropna()
filtered_values = filtered_values[filtered_values != 0]
min_value = np.min(filtered_values)
point=min_value
# print(point)
####################################
df1=pd.DataFrame(doc.iloc[:,12].dropna().iloc[1:]) #50*1
# print(df1)
data_org=pd.DataFrame(doc.iloc[:,19:]) #3195*50
# print(data_org)
data = pd.DataFrame()

#dark 위치 찾기
count_zero = (df1 == 0).sum()
# print("0의 개수:", count_zero)
dark=data_org.iloc[:,count_zero-1]
# print(dark)
################################
for col in doc.columns[19:]:
    # print(doc[col])

    column_data= doc[col]


    # column_data = df.iloc[2:,col]
    # print(column_data)
    max_value = column_data.max()


    if max_value > point:
        data[col] = column_data
# max 최대값 넘는지 안넘는지
###########################################
counts = data.iloc[1].value_counts()
counts = counts.sort_index(ascending=False)
# print(counts)
counts_all=np.cumsum(counts)
df=pd.DataFrame()
wave=pd.DataFrame(doc.iloc[2:,18])
df['Wavelength']=wave
# print(wave)

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
        # print(datt_21)
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

df.iloc[:, :] = df.iloc[:, :].applymap(lambda x: 0.5 if x <= 0 else x)

# print(df)
df.to_excel(output_file)