import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
# suc_state = 1
# log_state = 0
# password = str(123)
# if (log_state == 0) & (suc_state == 1):
#     input_pass = st.text_input("password")
#     if st.button("login"):
#         if input_pass == password:
#             suc_state = 0
#         else: suc_state = 1
# if suc_state == 0:
#     log_state =3

##実行
menu = ["一覧", "新規作成", "削除"]
state = st.sidebar.selectbox("menu", menu)
if state == "一覧":
    def to_excel(df):
        output = BytesIO()
        df.to_excel(output, index = False, sheet_name='Sheet1')
        processed_data = output.getvalue()
        return processed_data

    data_width = 1500


    house_list = ["指定なし"]
    where_list = ["指定なし"]
    df = pd.read_excel("仕上げDB.xlsx")
    df = df.fillna("ー")
    df.index = df.index + 1
    h_ar = df['使用物件'].unique()
    w_ar = df['使用箇所'].unique()
    for i in range(len(h_ar)):
        house_list.append(h_ar[i])
    for i in range(len(w_ar)):
        where_list.append(w_ar[i])

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
if state == "新規作成":
    df = pd.read_excel("仕上げDB.xlsx")
    st.title("新規作成")
    house = st.text_input("使用物件")
    col1, col2 = st.columns(2)
    where = col1.text_input("使用箇所")
    where2 = col2.text_input("使用部位")
    col_1, col_2, col_3 = st.columns(3)
    company = col_1.text_input("メーカー")
    goods_name = col_2.text_input("商品名")
    goods_namber = col_3.text_input("型番")
    col__1, col__2 = st.columns(2)
    goods_color = col__1.text_input("色番号")
    etc = col__2.text_input("備考")
    if st.button("新規作成"):
        df_new = pd.DataFrame({"使用物件": [house], "使用箇所":[where], "使用部位":[where2], "メーカー":[company], "商品名":[goods_name], "型番":[goods_namber], "色番号":[goods_color], "備考":[etc]})
        df = pd.concat([df, df_new])
        df.to_excel("仕上げDB.xlsx",index=False)
        st.success("成功しました")
if state == "削除":
    data_width = 1000
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
        df = df.drop(df[(df["使用物件"] == h_select) & (df["使用箇所"] == w_select) & (df["使用部位"] == w_2_select)].index, inplace=False)
        df.to_excel("仕上げDB.xlsx", index=False)





