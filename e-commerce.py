import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import streamlit  as st
import openpyxl
from babel.numbers import format_currency

def create_monthly_order(df):
    monthly_order = df.resample(rule="ME", on="order_purchase_timestamp").agg({
    "order_id": "nunique",
    "payment_value": "sum"
    })
    monthly_order = monthly_order.reset_index()
    monthly_order.rename(columns={
    "order_id": "order_count",
    "payment_value": "order_value"
    }, inplace=True)
    return monthly_order

def create_sum_order_producttype(df):
    sum_product_type = df.groupby(by="product_category_name_english").order_id.nunique().sort_values(ascending=False).reset_index()
    return sum_product_type

def create_seller_bycity(df):
    seller_bycity = df.groupby(by="seller_city").seller_id.nunique().sort_values(ascending=False).reset_index()
    return seller_bycity

def create_seller_bystate(df):
    seller_bystate = df.groupby(by="seller_state").seller_id.nunique().sort_values(ascending=False).reset_index()
    return seller_bystate

def create_topseller_byorder(df):
    top_seller_byorder = df.groupby(by="seller_id").agg({
    "order_id": "nunique",
    "payment_value": "sum",
    "review_score": "mean"
    }).sort_values(by="order_id", ascending=False).reset_index()
    top_seller_byorder.rename(columns={
    "order_id": "order_count",
    "payment_value": "revenue"
    }, inplace=True)
    return top_seller_byorder

def create_topseller_byrevenue(df):
    top_seller_byrevenue = df.groupby(by="seller_id").agg({
    "order_id": "nunique",
    "payment_value": "sum",
    "review_score": "mean"
    }).sort_values(by="payment_value", ascending=False).reset_index()
    top_seller_byrevenue.rename(columns={
    "order_id": "order_count",
    "payment_value": "revenue"
    }, inplace=True)
    return top_seller_byrevenue

def create_customer_bycity(df):
    customer_bycity = df.groupby(by="customer_city").customer_id.nunique().sort_values(ascending=False).reset_index()
    return customer_bycity

def create_customer_bystate(df):
    customer_bystate = df.groupby(by="customer_state").customer_id.nunique().sort_values(ascending=False).reset_index()
    return customer_bystate

def create_rfm(df):
    rfm = df.groupby(by="customer_id", as_index=False).agg({
        "order_purchase_timestamp": "max",
        "order_id": "nunique",
        "payment_value": "sum"
    })
    rfm.columns = ["customer_id", "max_order_timestamp", "frequency","monetary"]

    rfm["max_order_timestamp"] = rfm["max_order_timestamp"].dt.date
    recent_date = df["order_purchase_timestamp"].dt.date.max()
    rfm["recency"] = rfm["max_order_timestamp"].apply(lambda x: (recent_date - x).days)
    rfm.drop("max_order_timestamp", axis=1, inplace=True)
    return rfm

all_df = pd.read_excel("all_ecommerce.xlsx")

datetime_columns = ["order_purchase_timestamp"]
all_df.sort_values(by="order_purchase_timestamp", inplace=True)
all_df.reset_index(inplace=True)

for column in datetime_columns:
    all_df[column] = pd.to_datetime(all_df[column])

min_date = all_df["order_purchase_timestamp"].min()
max_date = all_df["order_purchase_timestamp"].max()

with st.sidebar:
    start_date, end_date = st.date_input(
        label="Rentang Waktu",
        min_value=min_date,
        max_value=max_date,
        value=[min_date, max_date]
    )

kondisi_lebih = all_df["order_purchase_timestamp"] >= pd.to_datetime(start_date)
kondisi_kurang = all_df["order_purchase_timestamp"] <= pd.to_datetime(end_date)
main_df = all_df[kondisi_lebih & kondisi_kurang]

monthly_order = create_monthly_order(main_df)
sum_product_type = create_sum_order_producttype(main_df)
seller_bycity = create_seller_bycity(main_df)
seller_bystate = create_seller_bystate(main_df)
topseller_byorder = create_topseller_byorder(main_df)
topseller_byrevenue = create_topseller_byrevenue(main_df)
customer_bycity = create_customer_bycity(main_df)
customer_bystate = create_customer_bystate(main_df)
rfm = create_rfm(all_df)

st.header("Ecommerce-Public Dashboard :sparkles:")

st.subheader("Monthly Orders")

col1, col2 = st.columns(2)

with col1:
    total_order = monthly_order.order_count.sum()
    st.metric("Total Orders", value=total_order)

with col2:
    total_order_value = format_currency(monthly_order.order_value.sum(), "BRL", locale="es_CO")
    st.metric("Total Order Value", value=total_order_value)

fig, ax = plt.subplots(figsize=(16,8))

ax.plot(
    monthly_order["order_purchase_timestamp"],
    monthly_order["order_count"],
    marker="o",
    linewidth=2,
    color="red"
)
ax.tick_params(axis="y", labelsize=20)
ax.tick_params(axis="x", labelsize=15)

st.pyplot(fig)

st.subheader("Best and Worst Performing Product Type")

fig, ax = plt.subplots(nrows=1, ncols=2, figsize=(35,15))
colors = ["blue", "gray", "gray", "gray","gray"]

sns.barplot(y="product_category_name_english", x="order_id", data=sum_product_type.head(), palette=colors, ax=ax[0])
ax[0].set_ylabel(None)
ax[0].set_xlabel(None)
ax[0].set_title("Best Performing Product Type", loc="center", fontsize=30)
ax[0].tick_params(axis="y", labelsize=35)
ax[0].tick_params(axis="x", labelsize=30)

sns.barplot(y="product_category_name_english", x="order_id", data=sum_product_type.sort_values(by="order_id", ascending=True).head(), palette=colors, ax=ax[1])
ax[1].set_ylabel(None)
ax[1].set_xlabel(None)
ax[1].invert_xaxis()
ax[1].yaxis.set_label_position("right")
ax[1].yaxis.tick_right()
ax[1].set_title("Worst Performing Product Type", loc="center", fontsize=30)
ax[1].tick_params(axis="y", labelsize=35)
ax[1].tick_params(axis="x", labelsize=30)

st.pyplot(fig)

st.subheader("The Highest Number of Seller by Region")

col1, col2 = st.columns(2)

with col1:
    fig, ax= plt.subplots(figsize=(20,10))
    colors = ["blue", "gray", "gray","gray","gray"]

    sns.barplot(x="seller_city", y="seller_id", data=seller_bycity.head(5), palette=colors)
    plt.ylabel(None)
    plt.xlabel(None)
    plt.title("By City", loc="center", fontsize=50)
    plt.tick_params(axis="x", rotation=25,labelsize=35)
    plt.tick_params(axis="y", labelsize=30)
    st.pyplot(fig)

with col2:
    fig, ax = plt.subplots(figsize=(20,10))
    colors = ["blue", "gray", "gray","gray","gray"]

    sns.barplot(x="seller_state", y="seller_id", data=seller_bystate.head(5), palette=colors)
    plt.ylabel(None)
    plt.xlabel(None)
    plt.title("By State", loc="center", fontsize=50)
    plt.tick_params(axis="x", rotation=25, labelsize=35)
    plt.tick_params(axis="y", labelsize=30)
    st.pyplot(fig)

st.subheader("Best Seller Based on Order Count")

fig, ax = plt.subplots(figsize=(20,10))
colors = ["blue", "gray", "gray", "gray","gray"]

sns.barplot(y="seller_id", x="order_count", data=topseller_byorder.head(), palette=colors, ax=ax)
ax.set_ylabel(None)
ax.set_xlabel(None)
ax.set_title("Best Seller by Order Count", loc="center", fontsize=30)
ax.tick_params(axis="y", labelsize=20)
ax.tick_params(axis="x", labelsize=15)
st.pyplot(fig)

col1, col2 = st.columns(2)

with col1:
    fig, ax = plt.subplots(figsize=(20,10))
    colors = ["blue", "gray", "gray", "gray","gray"]

    sns.barplot(x="seller_id", y="review_score", data=topseller_byorder.head(), palette=colors)
    plt.ylabel(None)
    plt.xlabel(None)
    plt.title("The Average Rating Score", loc="center", fontsize=50)
    plt.tick_params(axis="y", labelsize=35)
    plt.tick_params(axis="x", rotation=85, labelsize=30)
    st.pyplot(fig)

with col2:
    fig, ax = plt.subplots(figsize=(20,10))
    colors = ["blue", "gray", "gray", "gray","gray"]

    sns.barplot(x="seller_id", y="revenue", data=topseller_byorder.head(), palette=colors)
    plt.ylabel(None)
    plt.xlabel(None)
    plt.title("The Revenue", loc="center", fontsize=50)
    plt.tick_params(axis="y", labelsize=35)
    plt.tick_params(axis="x", rotation=85, labelsize=30)
    st.pyplot(fig)

st.subheader("Best Seller Based on Revenue")

fig, ax = plt.subplots(figsize=(20,10))
colors = ["blue", "gray", "gray", "gray","gray"]

sns.barplot(y="seller_id", x="revenue", data=topseller_byrevenue.head(), palette=colors, ax=ax)
ax.set_ylabel(None)
ax.set_xlabel(None)
ax.set_title("Best Seller by Revenue", loc="center", fontsize=30)
ax.tick_params(axis="y", labelsize=20)
ax.tick_params(axis="x", labelsize=15)
st.pyplot(fig)

col1, col2 = st.columns(2)

with col1:
    fig, ax = plt.subplots(figsize=(20,10))
    colors = ["blue", "gray", "gray", "gray","gray"]

    sns.barplot(x="seller_id", y="review_score", data=topseller_byrevenue.head(), palette=colors)
    plt.ylabel(None)
    plt.xlabel(None)
    plt.title("The Average Rating Score", loc="center", fontsize=50)
    plt.tick_params(axis="y", labelsize=35)
    plt.tick_params(axis="x", rotation=85, labelsize=30)
    st.pyplot(fig)

with col2:
    fig, ax = plt.subplots(figsize=(20,10))
    colors = ["blue", "gray", "gray", "gray","gray"]

    sns.barplot(x="seller_id", y="order_count", data=topseller_byrevenue.head(), palette=colors)
    plt.ylabel(None)
    plt.xlabel(None)
    plt.title("The Order Count", loc="center", fontsize=50)
    plt.tick_params(axis="y", labelsize=35)
    plt.tick_params(axis="x", rotation=85, labelsize=30)
    st.pyplot(fig)

st.subheader("The Highest Number of Customer by Region")

col1, col2 = st.columns(2)

with col1:
    fig, ax= plt.subplots(figsize=(20,10))
    colors = ["blue", "gray", "gray","gray","gray"]

    sns.barplot(x="customer_city", y="customer_id", data=customer_bycity.head(5), palette=colors)
    plt.ylabel(None)
    plt.xlabel(None)
    plt.title("By City", loc="center", fontsize=50)
    plt.tick_params(axis="x", rotation=25, labelsize=35)
    plt.tick_params(axis="y", labelsize=30)
    st.pyplot(fig)

with col2:
    fig, ax = plt.subplots(figsize=(20,10))
    colors = ["blue", "gray", "gray","gray","gray"]

    sns.barplot(x="customer_state", y="customer_id", data=customer_bystate.head(5), palette=colors)
    plt.ylabel(None)
    plt.xlabel(None)
    plt.title("By State", loc="center", fontsize=50)
    plt.tick_params(axis="x", rotation=25, labelsize=35)
    plt.tick_params(axis="y", labelsize=30)
    st.pyplot(fig)

st.subheader("Best Customer Based on RFM Parameters")

col1, col2, col3 = st.columns(3)

with col1:
    avg_recency = round(rfm.recency.mean(),1)
    st.metric("Average Recency (days)", value=avg_recency)

with col2:
    avg_frequency = round(rfm.frequency.mean(), 2)
    st.metric("Average Frequency", value=avg_frequency)

with col3:
    avg_monetary = format_currency(rfm.monetary.mean(), "BRL", locale='es_CO')
    st.metric("Average Monetary", value=avg_monetary)

fig, ax = plt.subplots(nrows=1, ncols=3, figsize=(40,15))

colors = ["blue", "blue", "blue", "blue", "blue"]

sns.barplot(y="recency", x="customer_id", data=rfm.sort_values(by="recency", ascending=True).head(5), palette=colors, ax=ax[0])
ax[0].set_ylabel(None)
ax[0].set_xlabel("customer_id", fontsize=30)
ax[0].set_title("By Recency (days)", loc="center", fontsize=50)
ax[0].tick_params(axis="y", labelsize=30)
ax[0].tick_params(axis="x", rotation=85, labelsize=35)

sns.barplot(y="frequency", x="customer_id", data=rfm.sort_values(by="frequency", ascending=False).head(5), palette=colors, ax=ax[1])
ax[1].set_ylabel(None)
ax[1].set_xlabel("customer_id", fontsize=30)
ax[1].set_title("By Frequency", loc="center", fontsize=50)
ax[1].tick_params(axis="y", labelsize=30)
ax[1].tick_params(axis="x", rotation=85, labelsize=35)

sns.barplot(y="monetary", x="customer_id", data=rfm.sort_values(by="monetary", ascending=False).head(5), palette=colors, ax=ax[2])
ax[2].set_ylabel(None)
ax[2].set_xlabel("customer_id", fontsize=30)
ax[2].set_title("By Monetary", loc="center", fontsize=50)
ax[2].tick_params(axis="y", labelsize=30)
ax[2].tick_params(axis="x", rotation=85, labelsize=35)
st.pyplot(fig)


st.caption("Copyright (c) Pangestu 2024")
