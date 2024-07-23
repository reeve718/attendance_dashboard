import streamlit as st
import pandas as pd
import plotly.express as px
import base64

sample_file = 'data_example.xlsx'
monthly_list = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
full_month_list = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
st.set_page_config(layout="wide", page_title="Dashboard", page_icon="📈")

def display_dashboard(dataFile: str = sample_file):
    
    st.header('Products Trend', divider='rainbow')
    excel_file = pd.ExcelFile(dataFile)

    df_list = []
    # Iterate over each sheet in the Excel file
    for sheet_name in excel_file.sheet_names:
        # Read the sheet into a DataFrame
        df = excel_file.parse(sheet_name)
        # Append the DataFrame to the list
        df_list.append(df)

    df_description_list = []
    for df in df_list:
        grouped = df.groupby('Product')
        description_per_product = grouped['Quantity'].describe()
        df_description_list.append(description_per_product)

    df_grouped_list = []
    for df in df_list:
        df_grouped = df.groupby("Product").sum()
        df_grouped_list.append(df_grouped)
    concatenated_df = pd.concat(df_grouped_list, ignore_index=False, axis=1)
    concatenated_df.columns = [str(i+1) for i in range(12)]

    concatenated_df_total = concatenated_df.copy()
    sum_dict = {}
    for i in range(len(df_grouped_list)):
        sum_dict[str(i+1)] = df_grouped_list[i]['Quantity'].sum()
    concatenated_df_total.loc['Total'] = sum_dict
    concatenated_df_total['Total'] = concatenated_df_total.sum(axis=1)

    concatenated_df = concatenated_df.reset_index()
    df_melted = concatenated_df.melt(id_vars='Product', var_name='Month', value_name='Sales')

    with st.container():
        col1, col2 = st.columns(2)
        with col1:      
            # Create a line chart using Plotly Express
            fig = px.line(df_melted, x='Month', y='Sales', color='Product')

        # Set the x and y-axis labels
            fig.update_layout(xaxis_title='Month', yaxis_title='Sales')

        # Display the chart using Streamlit
            st.plotly_chart(fig)
        with col2:
            fig2 = px.bar(df_melted, x="Month", y="Sales", color="Product")
            fig2.update_layout(xaxis_title='Month', yaxis_title='Sales')
            st.plotly_chart(fig2)

# Display the DataFrame as a table with centering style
    st.header('Monthly Summary', divider='rainbow')
    with st.container():
        st.dataframe(concatenated_df_total, use_container_width=True)

    st.header('Monthly Stats', divider='rainbow')
    # Create tabs based on the monthly_list
    for index, tab in enumerate(st.tabs(monthly_list), start=1):
        with tab:
            st.header(full_month_list[index-1])  # Display the full month name
            with st.container():
                st.dataframe(df_description_list[index-1], use_container_width=True)
        
uploaded_file = st.sidebar.file_uploader("Upload an Excel file", type=["xlsx", "xls"])
if st.sidebar.button("Download Sample Data"):
    with open(sample_file, "rb") as f:
        excel_data = f.read()
    b64 = base64.b64encode(excel_data).decode()
    href = f'<a href="data:application/octet-stream;base64,{b64}" download="sample_data.xlsx">Download Sample Data</a>'
    st.sidebar.markdown(href, unsafe_allow_html=True)

# Main content
if uploaded_file is not None:
    if uploaded_file.type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or uploaded_file.type == 'application/vnd.ms-excel':
        st.title("Dashboard")
        display_dashboard(uploaded_file)
    else:
        st.write("Please upload a valid Excel file.")
else:
    st.title("Dashboard (Demo)")
    display_dashboard()
    st.write("Upload a data file in Excel format by using the sidebar on the left.")