# File     : sale_app.py
# Title    : Streamlit銷售案例應用
# Date     : 2025.3.15
# Author   : Ming-Chang Lee
# YouTube  : https://www.youtube.com/@alan9956
# RWEPA    : http://rwepa.blogspot.tw/
# GitHub   : https://github.com/rwepa
# Email    : alan9956@gmail.com

# pip install plotly-express
# pip install streamlit

# 參考資料: https://github.com/Sven-Bo/streamlit-sales-dashboard
# Excel下載: https://github.com/rwepa/DataDemo/blob/master/superstore_tw.xlsx

import pandas as pd
import plotly.express as px  
import streamlit as st  

st.set_page_config(page_title="銷售儀表板", page_icon=":bar_chart:", layout="wide")

# 快取資料裝飾子
@st.cache_data
def get_data_from_excel():
    df = pd.read_excel(
        io="data/superstore_tw.xlsx",
        engine="openpyxl",
        sheet_name="訂單",
    )
    return df

df = get_data_from_excel()

# ---- SIDEBAR ----

st.sidebar.image("data/rwepa_logo.png")

st.sidebar.header("資料篩選")

area = st.sidebar.multiselect(
    "選取區域:",
    options=df["區域"].unique(),
    default=df["區域"].unique()
)

customer_type = st.sidebar.multiselect(
    "選取客戶型態:",
    options=df["細分"].unique(),
    default=df["細分"].unique(),
)

mailing = st.sidebar.multiselect(
    "選取郵寄方式:",
    options=df["郵寄方式"].unique(),
    default=df["郵寄方式"].unique()
)

df_selection = df.query(
    "區域 == @area & 細分 == @customer_type & 郵寄方式 == @mailing"
)

# Check if the dataframe is empty:
if df_selection.empty:
    st.warning("目前沒有選取任何資料!")
    st.stop() # This will halt the app from further execution.

# ---- MAINPAGE ----
st.title(":bar_chart: 銷售儀表板")
st.markdown("##")

# TOP KPI's
total_sales = int(df_selection["銷售額"].sum())

average_sales = round(df_selection["銷售額"].mean(), 1)

average_profits = round(df_selection["利潤"].mean(), 1)

net_profit_margin = round((df_selection["利潤"].sum()/df_selection["銷售額"].sum())*100, 2)

left_column, middle_column, right_column, right_last_column = st.columns(4)

with left_column:
    st.subheader("總銷售額:")
    st.subheader(f"{total_sales}")
    
with middle_column:
    st.subheader("平均銷售額:")
    st.subheader(f"{average_sales}")
    
with right_column:
    st.subheader("平均獲利:")
    st.subheader(f"{average_profits}")
    
with right_last_column:
    st.subheader("淨利率(%):")
    st.subheader(f"{net_profit_margin}")

st.markdown("""---""")

# SALES BY AREA [BAR CHART]
sales_by_area = df_selection.groupby(by=["區域"])[["銷售額"]].sum()

fig_area_sales = px.bar(
    sales_by_area,
    x=sales_by_area.index,
    y="銷售額",
    title="<b>區域銷售統計圖</b>",
    color_discrete_sequence=["#0083B8"] * len(sales_by_area),
    template="plotly_white",
)

fig_area_sales.update_layout(
    xaxis=dict(tickmode="linear"),
    plot_bgcolor="rgba(0,0,0,0)",
    yaxis=(dict(showgrid=False)),
)

# SALES BY PRODUCT TYPE [BAR CHART]
sales_by_product_type = df_selection.groupby(by=["類別"])[["銷售額"]].sum().sort_values(by="銷售額")

fig_product_sales = px.bar(
    sales_by_product_type,
    x="銷售額",
    y=sales_by_product_type.index,
    orientation="h",
    title="<b>類別銷售統計圖</b>",
    color_discrete_sequence=["#0083B8"] * len(sales_by_product_type),
    template="plotly_white",
)

fig_product_sales.update_layout(
    plot_bgcolor="rgba(0,0,0,0)",
    xaxis=(dict(showgrid=False))
)

left_column, right_column = st.columns(2)

left_column.plotly_chart(fig_area_sales, use_container_width=True)

right_column.plotly_chart(fig_product_sales, use_container_width=True)

# ---- HIDE STREAMLIT STYLE ----
hide_st_style = """
            <style>
            #MainMenu {visibility: hidden;}
            footer {visibility: hidden;}
            header {visibility: hidden;}
            </style>
            """
st.markdown(hide_st_style, unsafe_allow_html=True)
