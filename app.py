import streamlit as st
import pandas as pd
import plotly.express as px
import requests
import re

sample_file = 'data_example.xlsx'
default_url = "https://script.googleusercontent.com/macros/echo?user_content_key=CS1xGtawxafbsz_VXLceDZ8YhZJ3WLRGo3cTurMdyLFpRky5qzeXoqZBrYuJ0ElreSkNapo1i9TRXgBoQpou4xhHPHjfNUj4m5_BxDlH2jW0nuo2oDemN9CCS2h10ox_1xSncGQajx_ryfhECjZEnNGpsruz596piGbFJMaeWeZjFnhN0a9d4HwmOCyefi3ex5PPcpNsPlxGCUpGq7rwIA3gVkC6Pkbl8Md7CuyY-jH4-8QW02051A&lib=McNoStEX3P6w1EymfWUjOXqNKe58B53cM"
template_list = [ "Last 3 Months", "Last 6 Months", "Last 9 Months", "Last 12 Months","Overall"]
st.set_page_config(layout="wide", page_title="Dashboard", page_icon="📈")

def getDataJson(url):
    resp = requests.get(url)
    resp_json = resp.json()
    return resp_json

def transformData(df):
    strings = df.columns
    pattern = r"\[(.*?)\]"
    result = []
    for string in strings:
        matches = re.findall(pattern, string)
        result.extend(matches)
    df.columns = list(df.columns[:2]) + result
    df['date'] = pd.to_datetime(df['日期']).dt.strftime('%Y-%m-%d')
    cols = df.columns.tolist()[2:]
    cols = [cols[-1]] + cols[:-1]
    output_df = df[cols].copy()
    return output_df

def getDf_from_file(filePath):
    df = pd.read_excel(filePath)
    return transformData(df)

def getDf_from_google_sheet_json(resp_json):
    df = pd.DataFrame(resp_json['data'][1:], columns=resp_json['data'][0])
    return transformData(df)

def display_dashboard(dataFile: str = None):
    if dataFile:
        df = getDf_from_file(dataFile)
        df.fillna("", inplace=True)
    else:
        print("it is none")
        df = getDf_from_google_sheet_json(getDataJson(default_url))
        

    st.table(df)
    # df_description_list = []
    # for df in df_list:
    #     grouped = df.groupby('Product')
    #     description_per_product = grouped['Quantity'].describe()
    #     df_description_list.append(description_per_product)

    # df_grouped_list = []
    # for df in df_list:
    #     df_grouped = df.groupby("Product").sum()
    #     df_grouped_list.append(df_grouped)
    # concatenated_df = pd.concat(df_grouped_list, ignore_index=False, axis=1)
    # concatenated_df.columns = [str(i+1) for i in range(12)]

    # concatenated_df_total = concatenated_df.copy()
    # sum_dict = {}
    # for i in range(len(df_grouped_list)):
    #     sum_dict[str(i+1)] = df_grouped_list[i]['Quantity'].sum()
    # concatenated_df_total.loc['Total'] = sum_dict
    # concatenated_df_total['Total'] = concatenated_df_total.sum(axis=1)

    # concatenated_df = concatenated_df.reset_index()
    # df_melted = concatenated_df.melt(id_vars='Product', var_name='Month', value_name='Sales')

    # with st.container():
    #     col1, col2 = st.columns(2)
    #     with col1:      
    #         # Create a line chart using Plotly Express
    #         fig = px.line(df_melted, x='Month', y='Sales', color='Product')

    #     # Set the x and y-axis labels
    #         fig.update_layout(xaxis_title='Month', yaxis_title='Sales')

    #     # Display the chart using Streamlit
    #         st.plotly_chart(fig)
    #     with col2:
    #         fig2 = px.bar(df_melted, x="Month", y="Sales", color="Product")
    #         fig2.update_layout(xaxis_title='Month', yaxis_title='Sales')
    #         st.plotly_chart(fig2)

# Display the DataFrame as a table with centering style
    st.header('Present Rate Stats', divider='rainbow')
    # for index, tab in enumerate(st.tabs(template_list)):
    #     with tab:
    #         st.header(template_list[index])
    #         with st.container():
    #             st.dataframe(df_description_list[index], use_container_width=True)

    st.header('Cumulative Sum', divider='rainbow')
    # with st.container():
    #     #st.dataframe(concatenated_df_total, use_container_width=True)
    #     fig2 = px.bar(df_melted, x="Month", y="Sales", color="Product")
    #     fig2.update_layout(xaxis_title='Month', yaxis_title='Sales')
    #     st.plotly_chart(fig2)

        
uploaded_file = st.sidebar.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

# Main content
if uploaded_file is None:
    st.title("Attendance (Demo)")
    display_dashboard()
else:
    if uploaded_file.type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or uploaded_file.type == 'application/vnd.ms-excel':
        st.title("Attendance")
        display_dashboard(dataFile = uploaded_file)
    else:
        st.write("Please upload a valid Excel file.")