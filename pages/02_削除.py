import streamlit as st
import pandas as pd
import numpy as np

data_width = 1000

st.set_page_config(page_title="YKAA", layout="wide")

df = pd.read_excel("仕上げDB.xlsx")
house_list = []
where_list = []
where2_list = []
df.index = df.index + 1
df = df.fillna("ー")
h_ar = df['使用物件'].unique()
w_ar = df['使用箇所'].unique()
w2_ar = df['使用部位'].unique()
for i in range(len(h_ar)):
    house_list.append(h_ar[i])
for i in range(len(w_ar)):
    where_list.append(w_ar[i])
for i in range(len(w2_ar)):
    where2_list.append(w2_ar[i])
st.title("削除")
col1, col2, col3, col4 = st.columns(4)
h_select = col1.selectbox('使用物件', house_list)
w_select = col2.selectbox('使用箇所', where_list)
w_2_select = col3.selectbox('使用部位', where2_list)
select_df = df[(df["使用物件"] == h_select) & (df["使用箇所"] == w_select) & (df["使用部位"] == w_2_select)]
select_df.index = np.arange(1, len(select_df)+1)
if select_df.empty == True:st.error("データがありません")
else : st.dataframe(select_df, width=data_width)

if st.button("削除"):
    df.drop(df[(df["使用物件"] == h_select) & (df["使用箇所"] == w_select) & (df["使用部位"] == w_2_select)].index, inplace=False)
    df.to_excel("仕上げDB.xlsx", index=False)
    st.write("success")