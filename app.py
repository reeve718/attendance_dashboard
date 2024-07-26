import streamlit as st
import pandas as pd
import plotly.express as px
import requests
import re
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(layout="wide", page_title="Dashboard", page_icon="📈")
default_url = "https://script.google.com/macros/s/AKfycby0GYKbkWWVtkaBVd8qZYGLPrsXCfAPIOvezbcqxlmeUffzp5ThspGPPUDNQWt7Nrr_/exec"
template_list = [ "Last 3 Months", "Last 6 Months", "Last 9 Months", "Last 12 Months","Overall"]
present = "出席"
absent = "缺席"
leave = "請假"

@st.cache_data
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
@st.cache_data
def getDf_from_file(filePath):
    df = pd.read_excel(filePath)
    return transformData(df)

def getDf_from_google_sheet_json(resp_json):
    df = pd.DataFrame(resp_json['data'][1:], columns=resp_json['data'][0])
    return transformData(df)

def filter_last_x_months(df, x):
    if x == 0:
        return df
    today = datetime.today()
    x_months_ago = today - relativedelta(months=x)
    temp_df = df.copy()
    temp_df['date'] = pd.to_datetime(temp_df['date'])
    return df[temp_df['date'] >= x_months_ago]

def getDf_present_rate(df, prenset, absent, leave):
    result = pd.DataFrame()
    for column in df.columns[1:]:
        present_count = df[column].value_counts().get(prenset, 0)
        absent_count = df[column].value_counts().get(absent, 0)
        leave_count = df[column].value_counts().get(leave, 0)
        total_count = present_count + absent_count
        present_rate = (present_count / total_count * 100) if total_count > 0 else 0
        result.loc[column, 'Present Count'] = present_count
        result.loc[column, 'Absent Count'] = absent_count
        result.loc[column, 'Leave Count'] = leave_count
        result.loc[column, 'Present Rate (%)'] = round(present_rate,2)
    result['Present Count'] = result['Present Count'].astype(int)
    result['Absent Count'] = result['Absent Count'].astype(int)
    result['Leave Count'] = result['Leave Count'].astype(int)
    result = result.sort_index()
    return result

def display_dashboard(dataFile: str = None):
    if dataFile:
        df = getDf_from_file(dataFile)
        df.fillna("", inplace=True)
    else:
        df = getDf_from_google_sheet_json(getDataJson(default_url))

    x_months = [i for i in range(0,13,3)]
    x_months =  x_months[1:] + [x_months[0]]
    df_description_list = []
    for x in x_months:
        df_description_list.append(getDf_present_rate(filter_last_x_months(df,x), present, absent, leave))

    # Display the DataFrame as a table with centering style
    st.header('Present Rate', divider='rainbow')
    for index, tab in enumerate(st.tabs(template_list)):
        with tab:
            st.header(template_list[index])
            with st.container():
                height = (len(df_description_list[index]) + 1) * 35 + 3
                st.dataframe(df_description_list[index], use_container_width=True, height=height)

            st.header('Cumulative Sum', divider='rainbow')
            with st.container():
                plot_df = df_description_list[index].copy()
                plot_df.reset_index(inplace=True)
                plot_df.columns = ['Name', 'Present Count', 'Absent Count', 'Leave Count', 'Present Rate (%)']
                fig = px.histogram(plot_df, x = "Name", y = ["Present Count", "Absent Count", "Leave Count"])
                st.plotly_chart(fig)

st.title("Attendance")
uploaded_file = st.sidebar.file_uploader("Upload an Excel file", type=["xlsx", "xls"])

# Main content
if uploaded_file is None:
    
    display_dashboard()
else:
    if uploaded_file.type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' or uploaded_file.type == 'application/vnd.ms-excel':
        st.title("Attendance")
        display_dashboard(dataFile = uploaded_file)
    else:
        st.write("Please upload a valid Excel file.")