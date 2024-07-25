import streamlit as st
import pandas as pd
import plotly.express as px
import requests
import re
from datetime import datetime
from dateutil.relativedelta import relativedelta

st.set_page_config(layout="wide", page_title="Dashboard", page_icon="📈")
default_url = "https://script.googleusercontent.com/macros/echo?user_content_key=CS1xGtawxafbsz_VXLceDZ8YhZJ3WLRGo3cTurMdyLFpRky5qzeXoqZBrYuJ0ElreSkNapo1i9TRXgBoQpou4xhHPHjfNUj4m5_BxDlH2jW0nuo2oDemN9CCS2h10ox_1xSncGQajx_ryfhECjZEnNGpsruz596piGbFJMaeWeZjFnhN0a9d4HwmOCyefi3ex5PPcpNsPlxGCUpGq7rwIA3gVkC6Pkbl8Md7CuyY-jH4-8QW02051A&lib=McNoStEX3P6w1EymfWUjOXqNKe58B53cM"
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
    today = datetime.today()
    x_months_ago = today - relativedelta(months=x)
    temp_df = df.copy()
    temp_df['date'] = pd.to_datetime(temp_df['date'])
    return df[temp_df['date'] >= x_months_ago]

def getDf_present_rate(df, prenset, absent, leave):
    result = pd.DataFrame()
    for column in df.columns[2:]:
        present_count = df[column].value_counts().get(prenset, 0)
        absent_count = df[column].value_counts().get(absent, 0)
        leave_count = df[column].value_counts().get(leave, 0)
        total_count = present_count + absent_count + leave_count
        present_rate = (present_count / total_count * 100) if total_count > 0 else 0
        result.loc[column, 'Present Count'] = present_count
        result.loc[column, 'Absent Count'] = absent_count
        result.loc[column, 'Leave Count'] = leave_count
        result.loc[column, 'Present Rate (%)'] = round(present_rate,2)
    result['Present Count'] = result['Present Count'].astype(int)
    result['Absent Count'] = result['Absent Count'].astype(int)
    result['Leave Count'] = result['Leave Count'].astype(int)
    return result

def display_dashboard(dataFile: str = None):
    if dataFile:
        df = getDf_from_file(dataFile)
        df.fillna("", inplace=True)
    else:
        df = getDf_from_google_sheet_json(getDataJson(default_url))
        
    present_rate_df = getDf_present_rate(df, present, absent, leave)
    filtered_3_df = filter_last_x_months(df,3)
    present_rate_3_df = getDf_present_rate(filtered_3_df, present, absent, leave)
    filtered_6_df = filter_last_x_months(df,6)
    present_rate_6_df = getDf_present_rate(filtered_6_df, present, absent, leave)
    filtered_9_df = filter_last_x_months(df,9)
    present_rate_9_df = getDf_present_rate(filtered_9_df, present, absent, leave)
    filtered_12_df = filter_last_x_months(df,12)
    present_rate_12_df = getDf_present_rate(filtered_12_df, present, absent, leave)
    df_description_list = [present_rate_3_df, present_rate_6_df, present_rate_9_df, present_rate_12_df, present_rate_df]

# Display the DataFrame as a table with centering style
    st.header('Present Rate', divider='rainbow')
    for index, tab in enumerate(st.tabs(template_list)):
        with tab:
            st.header(template_list[index])
            with st.container():
                st.dataframe(df_description_list[index], use_container_width=True)

    st.header('Cumulative Sum', divider='rainbow')
    with st.container():
        plot_df = present_rate_df.copy()
        plot_df.reset_index(inplace=True)
        plot_df.columns = ['Name', 'Present Count', 'Absent Count', 'Leave Count', 'Present Rate (%)']
        fig = px.histogram(plot_df, x = "Name", y = ["Present Count", "Absent Count", "Leave Count"])
        st.plotly_chart(fig)

st.title("Attendance (Demo)")
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