import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(
    page_title="新規作成",
)

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