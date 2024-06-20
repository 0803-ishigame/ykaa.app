import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO

 

def to_excel(df):
    output = BytesIO()
    df.to_excel(output, index = False, sheet_name='Sheet1')
    processed_data = output.getvalue()

    return processed_data


data_width = 1500


house_list = ["指定なし"]
where_list = ["指定なし"]
df = pd.read_excel("./data/仕上げDB.xlsx")
df = df.fillna("ー")
df.index = df.index + 1
h_ar = df['使用物件'].unique()
w_ar = df['使用箇所'].unique()
for i in range(len(h_ar)):
    house_list.append(h_ar[i])
for i in range(len(w_ar)):
    where_list.append(w_ar[i])

st.set_page_config(page_title="YKAA", layout="wide")

st.title("使用仕上げ一覧")
col1, col2 = st.columns(2)
h_select = col1.selectbox('使用物件', house_list)
w_select = col2.selectbox('使用箇所', where_list)
if h_select == "指定なし" and w_select == "指定なし":
    st.dataframe(df, width=data_width)
    select_df = df
elif w_select == "指定なし":
    select_df = df[df["使用物件"] == h_select]
    select_df.index = np.arange(1, len(select_df)+1)
    st.dataframe(select_df, width=data_width)
elif h_select == "指定なし":
    select_df = df[df["使用箇所"] == w_select]
    select_df.index = np.arange(1, len(select_df)+1)
    st.dataframe(select_df, width=data_width)
else:
    select_df = df[(df["使用物件"] == h_select) & (df["使用箇所"] == w_select)]
    select_df.index = np.arange(1, len(select_df)+1)
    st.dataframe(select_df, width=data_width)
df_xlsx = to_excel(select_df)
st.download_button("EXCEL保存", df_xlsx, file_name="出力.xlsx")
 




