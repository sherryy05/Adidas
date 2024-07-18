import streamlit as st 
import plotly.express as px
import pandas as pd
import os
import warnings
import matplotlib.pyplot as plt
import plotly.figure_factory as ff
import plotly.graph_objects as go

warnings.filterwarnings('ignore')

st.set_page_config(page_title="Dashboard Streamlit", page_icon=":bar_chart:", layout="wide")

st.title(" :bar_chart: Dashboard Interactive Adidas Store")
st.markdown('<style>div.block-container{padding-top:1rem;}</style>', unsafe_allow_html=True)

# Directly read the Excel file
os.chdir(r"C:\Sherry Andiyani\Streamlit")
df = pd.read_excel("Adidas.xlsx")

col1, col2 = st.columns(2)
df["InvoiceDate"] = pd.to_datetime(df["InvoiceDate"])

# Getting the min and max date
startDate = pd.to_datetime(df["InvoiceDate"]).min()
endDate = pd.to_datetime(df["InvoiceDate"]).max()

with col1:
    date1 = pd.to_datetime(st.date_input("Start Date", startDate))

with col2:
    date2 = pd.to_datetime(st.date_input("End Date", endDate))

df = df[(df["InvoiceDate"] >= date1) & (df["InvoiceDate"] <= date2)].copy()

st.sidebar.header("Choose your filter: ")

# Create for Region
region = st.sidebar.multiselect("Pick your Region", df["Region"].unique())
if not region:
    df2 = df.copy()
else:
    df2 = df[df["Region"].isin(region)]

# Create for State
state = st.sidebar.multiselect("Pick the State", df2["State"].unique())
if not state:
    df3 = df2.copy()
else:
    df3 = df2[df2["State"].isin(state)]

# Create for City 
city = st.sidebar.multiselect("Pick the City", df3["City"].unique())

# Filter the data based on Region, State and City 
if not region and not state and not city:
    filtered_df = df
elif not state and not city:
    filtered_df = df[df["Region"].isin(region)]
elif not region and not city:
    filtered_df = df[df["State"].isin(state)]
elif state and city:
    filtered_df = df3[df["State"].isin(state) & df3["City"].isin(city)]
elif region and city:
    filtered_df = df3[df["Region"].isin(region) & df3["City"].isin(city)]
elif region and state:
    filtered_df = df3[df["Region"].isin(region) & df3["State"].isin(state)]
elif city:
    filtered_df = df3[df3["City"].isin(city)]
else:
    filtered_df = df3[df3["Region"].isin(region) & df3["State"].isin(state) & df3["City"].isin(city)]

# Group by Retailer and sum TotalSales
retailer_df = filtered_df.groupby(by=["Retailer"], as_index=False)["TotalSales"].sum()

# Format TotalSales as USD
retailer_df["TotalSales"] = retailer_df["TotalSales"].apply(lambda x: '${:,.2f}'.format(x))

# Sort retailer_df by TotalSales descending
retailer_df = retailer_df.sort_values(by="TotalSales", ascending=True)

# Group by Region and sum TotalSales
region_df = filtered_df.groupby(by=["Region"], as_index=False)["TotalSales"].sum()

# Format TotalSales as USD
region_df["TotalSales"] = region_df["TotalSales"].apply(lambda x: '${:,.2f}'.format(x))

# Sort Region by TotalSales descending
region_df = region_df.sort_values(by="TotalSales", ascending=True)

with col1:
    st.subheader("Retailer wise Sales")
    fig = px.bar(retailer_df, x="Retailer", y="TotalSales", 
                 labels={"Retailer": "Retailer", "TotalSales": "Total Sales (USD)"},
                 template="seaborn")
    fig.update_layout(xaxis_categoryorder='total descending')  # Mengatur urutan kategori batang dari terbesar ke terkecil
    st.plotly_chart(fig, use_container_width=True, height=400)
    
with col2:
    st.subheader("Region wise Sales")
    fig = px.pie(filtered_df, values="TotalSales", names="Region", hole=0.5,
                 labels={"Region": "Region", "TotalSales": "Total Sales (USD)"})
    st.plotly_chart(fig, use_container_width=True)

cl1, cl2 = st.columns(2)
with cl1:
    with st.expander("Data Retailer"):
        st.write(retailer_df.style.background_gradient(cmap="Blues"))
        csv = retailer_df.to_csv(index=False).encode('utf-8')
        st.download_button("Download Data", data=csv, file_name="Retailer.csv", mime="text/csv",
                           help='Click here to download the data as a CSV file')

with cl2:
    with st.expander("Data Region"):
        st.write(region_df.style.background_gradient(cmap="Oranges"))
        csv = region_df.to_csv(index=False).encode('utf-8')
        st.download_button("Download Data", data=csv, file_name="Region.csv", mime="text/csv",
                           help='Click here to download the data as a CSV file')

filtered_df["month_year"] = filtered_df["InvoiceDate"].dt.to_period("M")
st.subheader('Time Series Analysis')

linechart = pd.DataFrame(filtered_df.groupby(filtered_df["month_year"].dt.strftime("%b - %Y"))["TotalSales"].sum()).reset_index()

# Format TotalSales as USD
linechart["TotalSales"] = linechart["TotalSales"].apply(lambda x: '${:,.2f}'.format(x))

fig2 = px.line(linechart, x="month_year", y="TotalSales", 
               labels={"month_year": "Month-Year", "TotalSales": "Total Sales (USD)"}, 
               height=500, width=1000, template="gridon")
fig2.update_traces(textposition='top center')
st.plotly_chart(fig2, use_container_width=True)

with st.expander("View Data of TimeSeries: "):
    st.write(linechart.T.style.background_gradient(cmap="Blues"))
    csv = linechart.to_csv(index=False).encode("utf-8")
    st.download_button('Download Data', data=csv, file_name="TimeSeries.csv", mime='text/csv')

st.subheader("Hierarchical view of Sales using TreeMap")
fig3 = px.treemap(filtered_df, path=["Region", "Retailer", "Product"], values="TotalSales", hover_data=["TotalSales"],
                  color="Product")

fig3.update_traces(textinfo='label+value', hovertemplate='<b>%{label}</b><br>Total Sales: $%{value:,.2f}')
fig3.update_layout(width=800, height=650)
st.plotly_chart(fig3, use_container_width=True)

chart1, chart2 = st.columns((2))
with chart1:
    st.subheader("Sales Method wise Sales")
    fig = px.pie(filtered_df, values="TotalSales", names="SalesMethod", template="plotly_dark")
    fig.update_traces(textinfo='label+value', hovertemplate='<b>%{label}</b><br>Total Sales: $%{value:,.2f}', textposition="inside")
    st.plotly_chart(fig, use_container_width=True)

with chart2:
    st.subheader('Retailer wise Sales')
    fig = px.pie(filtered_df, values="TotalSales", names="Retailer", template="plotly_dark")
    fig.update_traces(textinfo='label+value', hovertemplate='<b>%{label}</b><br>Total Sales: $%{value:,.2f}', textposition="inside")
    st.plotly_chart(fig, use_container_width=True)

# Ubah format TotalSales dan OperatingProfit menjadi numeric
filtered_df["TotalSales"] = pd.to_numeric(filtered_df["TotalSales"].replace({'\$': '', ',': ''}, regex=True))
filtered_df["OperatingProfit"] = pd.to_numeric(filtered_df["OperatingProfit"].replace({'\$': '', ',': ''}, regex=True))

st.subheader(":point_right: Month wise Product Sales Summary")
with st.expander("Summary_Table"):
    # Ambil 5 baris pertama dan pilih kolom yang diperlukan
    df_sample = df.head(5)[["Region", "State", "City", "Retailer", "UnitsSold", "TotalSales", "OperatingProfit"]]
    
    # Format kolom TotalSales dan OperatingProfit menjadi dollar
    df_sample["TotalSales"] = df_sample["TotalSales"].apply(lambda x: '${:,.2f}'.format(x))
    df_sample["OperatingProfit"] = df_sample["OperatingProfit"].apply(lambda x: '${:,.2f}'.format(x))
    
    # Buat tabel interaktif menggunakan Plotly
    fig = ff.create_table(df_sample, colorscale="Cividis")
    st.plotly_chart(fig, use_container_width=True)

    st.markdown("Month wise Product Table")
    filtered_df["month"] = filtered_df["InvoiceDate"].dt.month_name()
    product_Year = pd.pivot_table(data=filtered_df, values="TotalSales", index=["Product"], columns="month", aggfunc='sum', fill_value=0)
    
    # Format kolom TotalSales menjadi dollar
    product_Year = product_Year.applymap(lambda x: '${:,.2f}'.format(x))
    
    st.write(product_Year.style.background_gradient(cmap="Blues"))
    
# Create a scatter plot
data1 = px.scatter(filtered_df, x="TotalSales", y="OperatingProfit", size="UnitsSold")
data1['layout'].update(title="Relationship between Sales and Provit using Scatter Plot.",
                       titlefont=dict(size=20), xaxis=dict(title="Sales", titlefont=dict(size=19)),
                       yaxis=dict(title="OperatingProfit", titlefont=dict(size=19)))

st.plotly_chart(data1, use_container_width=True)

# Hitung operating margin
filtered_df["OperatingMargin"] = (filtered_df["OperatingProfit"] / filtered_df["TotalSales"]) * 100

# Tampilkan data dengan operating margin dan format dollar
with st.expander("View Data with Operating Margin"):
    # Buat salinan untuk tampilan agar tidak mengubah data asli
    filtered_df_display = filtered_df.copy()
    
    # Ubah format PricePerUnit, TotalSales, dan OperatingProfit menjadi dollar
    filtered_df_display["PriceperUnit"] = filtered_df_display["PriceperUnit"].apply(lambda x: '${:,.2f}'.format(x))
    filtered_df_display["TotalSales"] = filtered_df_display["TotalSales"].apply(lambda x: '${:,.2f}'.format(x))
    filtered_df_display["OperatingProfit"] = filtered_df_display["OperatingProfit"].apply(lambda x: '${:,.2f}'.format(x))
    
    # Tampilkan data dengan background gradient
    st.write(filtered_df_display.iloc[:500, 1:20:2].style.background_gradient(cmap="Oranges"))

# Group by month_year and sum TotalSales and UnitsSold
linechart = pd.DataFrame(filtered_df.groupby("month_year")[["TotalSales", "UnitsSold"]].sum()).reset_index()

# Convert month_year to string for JSON serialization
linechart["month_year"] = linechart["month_year"].astype(str)

# Create the combo chart
fig_combo = go.Figure()

# Add bar trace for TotalSales
fig_combo.add_trace(go.Bar(x=linechart["month_year"], y=linechart["TotalSales"], name="Total Sales", yaxis='y1'))

# Add line trace for UnitsSold
fig_combo.add_trace(go.Scatter(x=linechart["month_year"], y=linechart["UnitsSold"], name="Units Sold", yaxis='y2', mode='lines+markers'))

fig_combo.update_layout(
    xaxis=dict(title="Month-Year"),
    yaxis=dict(title="Total Sales (USD)", tickformat="$,.2f"),
    yaxis2=dict(title="Units Sold", overlaying='y', side='right'),
    title="Combo Chart of Total Sales and Units Sold over Time",
    template="plotly_dark"
)

# Display the combo chart
st.plotly_chart(fig_combo, use_container_width=True)

# Download original Dataset
csv = df.to_csv(index=False).encode('utf-8')
st.download_button('Download Data', data=csv, file_name="Data.csv", mime="text/csv")
