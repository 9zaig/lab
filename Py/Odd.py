import pandas as pd

input_file = "./Input/5um.xlsx"
output_file = "./Output/5um_R.xlsx"
df = pd.read_excel(input_file)
first_column = df.iloc[:, 0]
selected_columns = df.iloc[:, 1::2]
new_df = pd.concat([first_column, selected_columns], axis=1)
new_df.to_excel(output_file, index=False)
