import streamlit as st
import pandas as pd
import plotly.express as px
import webbrowser as wb
import openpyxl
import os
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
import io
from datetime import datetime

today_date = datetime.now().strftime('%Y-%m-%d')

st.set_page_config(page_title="Dashboard", page_icon=":bar_chart:", layout="wide")

#st.header("File Upload")
#data_file = st.file_uploader(".xlsx file",type=['xlsx'])
#st.markdown("#")

st.title(":bar_chart: Dashboard ")

df1 = pd.read_excel("Example Datas.xlsx",sheet_name="Web")
df2 = pd.read_excel("Example Datas.xlsx",sheet_name="Shopee")
df3 = pd.read_excel("Example Datas.xlsx",sheet_name="Lazada")
df4 = pd.read_excel("Example Datas.xlsx",sheet_name="TikTok")

df1.drop(df1.index[:4], inplace=True)
df1.columns = df1.iloc[0]
df1 = df1[1:]
df1.reset_index(drop=True, inplace=True)
#df1
df2.drop(df2.index[:4], inplace=True)
df2.columns = df2.iloc[0]
df2 = df2[1:]
df2.reset_index(drop=True, inplace=True)
#df2
df3.drop(df3.index[:4], inplace=True)
df3.columns = df3.iloc[0]
df3 = df3[1:]
df3.reset_index(drop=True, inplace=True)
#df3
df4.drop(df4.index[:4], inplace=True)
df4.columns = df4.iloc[0]
df4 = df4[1:]
df4.reset_index(drop=True, inplace=True)
#df4

df = pd.concat([df1, df2, df3, df4])
#df

# Sample DataFrame with 'state', 'latitude', and 'longitude' for locations in Malaysia
states_coordinates = pd.DataFrame({
    'Shipping State': ['Johor', 'Kedah', 'Kelantan', 'Malacca', 'Negeri Sembilan', 'Pahang', 'Penang', 'Perak', 'Perlis', 'Sabah', 'Sarawak', 'Selangor', 'Terengganu'],
    'latitude': [1.4854, 6.1184, 6.1254, 2.1896, 2.7258, 3.8126, 5.4164, 4.5921, 6.4449, 5.9804, 1.5533, 3.0738, 5.3117],
    'longitude': [103.7618, 100.3682, 102.2386, 102.2501, 101.9424, 103.3256, 100.3308, 101.0901, 100.2048, 116.0753, 110.3441, 101.5183, 103.1324]
})

df_state = df.merge(states_coordinates, on='Shipping State', how='left')
df_state = df_state.dropna(subset=['latitude'])
#df_state

# Now try to render the map again
#st.map(df_state)

# Step 1: Count the occurrences of each 'Shipping State'
state_counts = df_state['Shipping State'].value_counts().rename_axis('Shipping State').reset_index(name='counts')
#state_counts
# Step 2: Merge the counts with the original DataFrame
# Ensure your original DataFrame has unique entries for each 'Shipping State's coordinates
df_unique_coords = df_state.drop_duplicates(subset=['Shipping State'])
#df_unique_coords
df_with_counts = pd.merge(state_counts, df_unique_coords, on='Shipping State')
#df_with_counts

df_with_counts.rename(columns={
    'latitude': 'lat',
    'longitude': 'lon',
    # Add more columns as needed
}, inplace=True)

df['Unit Total'] = df['Unit Total'].replace('-', 0)
total_sales = df['Unit Total'].sum()
total_sales = total_sales/1000
total_sales = int(total_sales)

df['Margin Per Item'] = df['Margin Per Item'].replace('-', 0)
total_profit = df['Margin Per Item'].sum()
total_profit = total_profit/1000
total_profit = int(total_profit)

total_qty = df['Quantity'].sum()

average_margin = round(df['Margin Per Item'].mean(),2)

profit_per= int(total_profit/total_sales *100)

df_marketplace = df['Order Source'].value_counts()
df_payment = df['Payment Method'].value_counts()
df_courier = df['Courier'].value_counts()

product_sales = df.groupby('Model')['Unit Total'].sum().reset_index()
###########################################################################################################################################

# Custom CSS to increase the font size of the metric value
st.markdown("""
<style>
[data-testid="stMetricValue"] { font-size: 36px; }
[data-testid="stMetricLabel"] { font-size: 18px; }
</style>
""", unsafe_allow_html=True)

# Your columns with metrics
#st.markdown('### Metrics')
col1, col2, col3, col4, col5 = st.columns(5)
with col1.container():
    col1.metric("Sales", f"RM {total_sales}k", "RM 67k")
with col2.container():
    col2.metric("Profit", f"RM{total_profit}k", "-RM 1k")
with col3.container():
    col3.metric("Total Item Sold", f"{total_qty} units", "271 units")
with col4.container():
    col4.metric("Average Margin per item", f"{average_margin}", "2.46")
with col5.container():
    col5.metric("Profit Percentage", f"{profit_per}%", "-3%")



colB1, colB2 = st.columns([2.2,3])

mapfig = px.scatter_geo(df_with_counts,
                     lat='lat',
                     lon='lon',
                     size='counts',
                     color='Shipping State',  # Set the color based on the state
                     hover_name='Shipping State',
                     projection='natural earth')

# Ensure that the land and country borders are visible
mapfig.update_geos(
    visible=True,  # This ensures that the geographic layout is visible
    showcountries=True,  # This shows country borders
    countrycolor="Black"  # You can customize the country border color
)

# Update the layout to focus on Malaysia
mapfig.update_layout(
    geo=dict(
        scope='asia',  # Set the scope to 'asia'
        center={'lat': 4.2105, 'lon': 108.5},  # Center the map on Malaysia
        projection_scale=6,
        showland=True,  # Ensure the land is shown
        landcolor='rgb(217, 217, 217)',  # Set the land color
        countrywidth=0.5  # Set the country border width
    )
)

with colB1:
    mapfig.update_layout(width=630, height=400, title='Total Order by State')
    st.plotly_chart(mapfig)

# Sample data
months = ['January', 'February', 'March', 'April', 'May', 'June',
          'July', 'August', 'September', 'October', 'November', 'December']
profits = [20000, 15000, 17000, 25000, 22000, 19000, 30000, 28000, 24000, 35000, 33000, 36000]
sales = [80000, 70000, 90000, 85000, 88000, 83000, 100000, 99000, 93000, 110000, 108000, 112000]

# Create subplots and mention that we want to share the x-axis
barline_fig = make_subplots(specs=[[{"secondary_y": True}]])

# Add bar chart for profits
barline_fig.add_trace(
    go.Bar(x=months, y=profits, name='Monthly Profits'),
    secondary_y=False,
)

# Add line chart for sales
barline_fig.add_trace(
    go.Scatter(x=months, y=sales, name='Monthly Sales', mode='lines+markers'),
    secondary_y=True,
)

# Add titles and labels
barline_fig.update_layout(
    title_text='Monthly Profits and Sales',
    xaxis_title='Month',
    yaxis_title='Profit ($)'
)

# Set y-axes titles
barline_fig.update_yaxes(title_text="Profit ($)", secondary_y=False)
barline_fig.update_yaxes(title_text="Sales ($)", secondary_y=True)

with colB2:
    barline_fig.update_layout(width=900, height=400)
    st.plotly_chart(barline_fig)

colC1, colC2, colC3, colC4 = st.columns([2,1,1,1])
with colC1:
    state_sales = df.groupby(['Shipping State', 'Category'])[['Unit Total']].sum().reset_index()
    # Step 2: Pivot the DataFrame
    pivot_df = state_sales.pivot(index='Shipping State', columns='Category', values='Unit Total')
    # Step 3: Create the stacked bar chart
    stackbar_fig = px.bar(pivot_df, barmode='stack')
    # Show the figure
    stackbar_fig.update_layout(width=500,height=350, title='Sales by State and Category', showlegend=False)
    st.plotly_chart(stackbar_fig)
with colC2:
    # Create the pie chart
    pie_fig1 = go.Figure(data=[go.Pie(labels=df_marketplace.index, values=df_marketplace.values, hole=0.5)])

    pie_fig1.update_layout(width=300,height=350, title='Order by MarketPlace')
    st.plotly_chart(pie_fig1)
with colC3:
    # Create the pie chart
    pie_fig2 = go.Figure(data=[go.Pie(labels=df_payment.index, values=df_payment.values, hole=0.7, textinfo="none")])
    pie_fig2.update_layout(width=250,height=350, showlegend=False, title='Preferred Payment Methods')
    st.plotly_chart(pie_fig2)
with colC4:
    # Create the pie chart
    pie_fig3 = go.Figure(data=[go.Pie(labels=df_courier.index, values=df_courier.values)])
    pie_fig3.update_layout(width=250,height=350, showlegend=False, title='Shipping Courier')
    st.plotly_chart(pie_fig3)

colD1, colD2 = st.columns(2)
with colD1:
    # Sort the DataFrame by the 'Sales' column in descending order and select the top 5
    top_products =  product_sales.sort_values('Unit Total', ascending=True).tail(5)
    # Create the bar graph with the updated variable name
    ybar_fig = px.bar(top_products, x='Unit Total', y='Model', orientation='h', title='Top 5 Products by Sales', color_discrete_sequence=['green'])
    # Show the figure
    ybar_fig.update_layout(width=600,height=350)
    st.plotly_chart(ybar_fig)
with colD2:
    # Sort the DataFrame by the 'Sales' column in descending order and select the top 5
    bot_products = product_sales[product_sales['Unit Total'] != 0]
    bot_products = bot_products.sort_values('Unit Total', ascending=False).tail(5)
    # Create the bar graph with the updated variable name
    ybar_fig = px.bar(bot_products, x='Unit Total', y='Model', orientation='h', title='Bottom 5 Products by Sales', color_discrete_sequence=['maroon'])
    # Show the figure
    ybar_fig.update_layout(width=600,height=350)
    st.plotly_chart(ybar_fig)
