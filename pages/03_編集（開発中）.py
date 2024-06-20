import streamlit as st
import pandas as pd
import numpy as np

def update(df, house, where, where2, company, goods_name, goods_number, goods_color, etc):
    df_new = pd.DataFrame({"使用物件": [house], "使用箇所":[where], "使用部位":[where2], "メーカー":[company], "商品名":[goods_name], "型番":[goods_number], "色番号":[goods_color], "備考":[etc]})
    df = pd.concat([df, df_new])
    df.to_excel("./data/仕上げDB.xlsx",index=False)
    print(df)


st.set_page_config(page_title="YKAA", layout="wide")

df = pd.read_excel("./data/仕上げDB.xlsx")
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
st.title("編集")
col1, col2, col3, col4 = st.columns(4)
h_select = col1.selectbox('使用物件', house_list)
w_select = col2.selectbox('使用箇所', where_list)
w_2_select = col3.selectbox('使用部位', where2_list)
if st.button("検索"):
    df_select = df[(df['使用物件'] == h_select) & (df['使用箇所'] == w_select) & (df['使用部位'] == w_2_select)]
    df = df[(df['使用物件'] != h_select) | (df['使用箇所'] != w_select) | (df['使用部位'] != w_2_select)]
    if df_select.empty == True:st.error("データがありません")
        #after
    else:
        st.header("After")
        house = st.text_input("使用物件", value=df_select.iat[0, 0])
        where = st.text_input("使用箇所", value=df_select.iat[0, 1])
        where2 = st.text_input("使用部位", value=df_select.iat[0, 2])
        company = st.text_input("メーカー", value=df_select.iat[0, 3])
        goods_name = st.text_input("商品名", value=df_select.iat[0, 4])
        goods_number = st.text_input("型番", value=df_select.iat[0, 5])
        goods_color = st.text_input("色番号", value=df_select.iat[0, 6])
        etc = st.text_input("備考", value=df_select.iat[0, 7])
st.button("編集", on_click= update(df, house, where, where2, company, goods_name, goods_number, goods_color, etc))
           
                
if st.button("キャンセル"):
    df = pd.concat([df, df_select])
    df.to_excel("./data/仕上げDB.xlsx",index=False)
    print(df)