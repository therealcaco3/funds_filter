import io
import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter
import pandas as pd
import numpy as np
import streamlit as st
from st_aggrid import AgGrid

# Convert and extract data from the selected worksheet
def data_organize(file_path, selected_sheet_name):
    # Check the selected sheet
    if selected_sheet_name == '境內(TWD計價) -  ':
        # Load the Domestic worksheet into a DataFrame, excluding columns F(5) and G(6) 
        df = pd.read_excel(file_path, sheet_name=selected_sheet_name, header=None,
                            skiprows=11,    # starts from A12, so skip the first 11 rows
                            usecols=lambda x: x not in [5, 6])   # column indices in pandas are zero-based, so F = 5, G = 6.
        
        # check how many columns are in the dataframe
        num_columns = df.shape[1]

        # Rename the columns
        column_names = ['SITCA Domestic', '理柏 ID', 'ISIN 代碼', '名稱', '基金貨幣', '1M', '1M排名', '3M',
                        '3M排名', '6M', '6M排名', '1Y', '1Y排名', '2Y', '2Y排名', '3Y', '3Y排名', '5Y', '5Y排名',
                        '10Y', '10Y排名', '波動度 1Y', '波動度 3Y', '波動度 4Y']
        
        column_names_20 = ['SITCA Domestic', '理柏 ID', 'ISIN 代碼', '名稱', '基金貨幣', '1M', '1M排名', '3M',
                        '3M排名', '6M', '6M排名', '1Y', '1Y排名', '2Y', '2Y排名', '3Y', '3Y排名',
                        '波動度 1Y', '波動度 3Y', '波動度 4Y']
        
        # if the dataframe doesn't have 24 columns, then use column_names_20
        if num_columns == 24:
            df.columns = column_names
        elif num_columns == 20:
            df.columns = column_names_20


    elif selected_sheet_name == '境外(USD計價) -  ':
        # Load the Overseas worksheet into a DataFrame, excluding columns A(0)
        df = pd.read_excel(file_path, sheet_name=selected_sheet_name, header=None,
                        skiprows=11,    # starts from A12, so skip the first 11 rows
                        usecols=lambda x: x not in [0])   # column indices in pandas are zero-based
        
        # check how many columns are in the dataframe
        num_columns = df.shape[1]
        
        # Rename the columns
        column_names = ['理柏環球分類', '理柏 ID', 'ISIN 代碼', '名稱', '基金貨幣', 'Aggregate Fund Value USD 日期',
                        'Aggregate Fund Value USD 數值', '1M', '1M排名', '3M', '3M排名', '6M', '6M排名', 
                        '1Y', '1Y排名', '2Y', '2Y排名', '3Y', '3Y排名', '5Y', '5Y排名',
                        '10Y', '10Y排名', '波動度 1Y', '波動度 3Y']
        
        column_names_22 = ['理柏環球分類', '理柏 ID', 'ISIN 代碼', '名稱', '基金貨幣', 'Aggregate Fund Value USD 日期',
                        'Aggregate Fund Value USD 數值', '1M', '1M排名', '3M', '3M排名', '6M', '6M排名', 
                        '1Y', '1Y排名', '2Y', '2Y排名', '3Y', '3Y排名', '波動度 1Y', '波動度 3Y', '理柏ID']

        # if the dataframe doesn't have 25 columns, then use column_names_22
        if num_columns == 25:
            df.columns = column_names
        elif num_columns == 22:
            df.columns = column_names_22

    # Drop rows where column B is blank
    df = df.dropna(subset=['理柏 ID'])

    # Reset the index of the DataFrame
    df = df.reset_index(drop=True)

    # Determine the last row based on the presence of data in column B
    last_row_index = df['理柏 ID'].last_valid_index()
    data = df.loc[:last_row_index]
    
    return data


# Define domestic filter
def funds_filter(sheet_name, data, classification, figures, thresholds):
    # Check whether the selected sheet is Domestic or Overseas
    if sheet_name == '境外(USD計價) -  ':
        filtered_domestics = data.loc[data["理柏環球分類"] == classification]
    elif sheet_name == '境內(TWD計價) -  ':
        filtered_domestics = data.loc[data["SITCA Domestic"] == classification]

    # Select the desired figures (columns)
    filtered_figures = filtered_domestics[figures]
    
    # Extract the "名稱" column
    names = filtered_domestics["名稱"]
    
    # Combine the "名稱" column with filtered results
    filtered_figures = pd.concat([names, filtered_figures], axis=1)

    # Convert thresholds from percentage to decimal values
    thresholds_divide = [float(t) / 100 for t in thresholds]

    # create another list of figures that users actually need
    sel_figures = [figures[0]]
    sel_thresholds = [thresholds_divide[0]]

    for n in range(1, len(figures)):
        if thresholds_divide[n] != 1:
            sel_figures.append(figures[n])
            sel_thresholds.append(thresholds_divide[n])

    # Apply the user-defined thresholds to filter the securities
    filtered_results = filtered_figures.loc[
                        (filtered_figures[sel_figures[0]].notna()) &
                        (filtered_figures[sel_figures[0]] <= filtered_figures[sel_figures[0]].quantile(sel_thresholds[0]))]
    for i in range(1, len(sel_figures)):
        filtered_results = filtered_results.loc[
                        (filtered_results[sel_figures[i]].notna()) &
                        (filtered_results[sel_figures[i]] <= filtered_results[sel_figures[i]].quantile(sel_thresholds[i]))]
    
    # converting '排名' columns from float to int
    rank_cols = [r for r in filtered_results.columns if '排名' in r]
    filtered_results[rank_cols] = filtered_results[rank_cols].apply(pd.to_numeric, downcast='integer', errors='ignore')

    return filtered_results


# Adjust appearance of the front-end website
st.set_page_config(
    page_title="基金篩選器",
    page_icon=":moneybag:",
    layout="centered",
    initial_sidebar_state="auto",
)

st.title('基金篩選器')


# User input using Streamlit
st.markdown("#### 請上傳檔案：")
uploaded_file = st.file_uploader("上傳 .xlsx 檔案", type=".xlsx")

if uploaded_file is not None:
    # Read the uploaded file
    file_content = uploaded_file.read()
    wb = pd.read_excel(io.BytesIO(file_content), sheet_name=None)
    sheet_names = wb.keys()
    sheet_input = st.selectbox("請選擇工作表：", list(sheet_names))
    
    # Perform data processing and filtering based on the selected sheet and other inputs
    classification = st.text_input("請選擇基金分類：")

    # Users enter thresholds for each rank
    st.header('請輸入各數據排名欲取的報酬率百分比，例：取報酬率排名前50%的基金，請輸入"50"；預設為100%')
    rank_1M = st.number_input(label = '請輸入1M報酬率排名百分比：', value = 100, max_value = 100, min_value = 0)
    rank_3M = st.number_input(label = '請輸入3M報酬率排名百分比：', value = 100, max_value = 100, min_value = 0)
    rank_6M = st.number_input(label = '請輸入6M報酬率排名百分比：', value = 100, max_value = 100, min_value = 0)
    rank_1Y = st.number_input(label = '請輸入1Y報酬率排名百分比：', value = 100, max_value = 100, min_value = 0)
    rank_2Y = st.number_input(label = '請輸入2Y報酬率排名百分比：', value = 100, max_value = 100, min_value = 0)
    rank_3Y = st.number_input(label = '請輸入3Y報酬率排名百分比：', value = 100, max_value = 100, min_value = 0)
    rank_5Y = st.number_input(label = '請輸入5Y報酬率排名百分比：', value = 100, max_value = 100, min_value = 0)
    rank_10Y = st.number_input(label = '請輸入10Y報酬率排名百分比：', value = 100, max_value = 100, min_value = 0)

    selected_sheet_name = '境外(USD計價) -  ' if sheet_input == '境外(USD計價) -  ' else '境內(TWD計價) -  '
    figures = ['1M排名', '3M排名', '6M排名', '1Y排名', '2Y排名', '3Y排名', '5Y排名', '10Y排名']
    thresholds = [rank_1M, rank_3M, rank_6M, rank_1Y, rank_2Y, rank_3Y, rank_5Y, rank_10Y]

    # Execute the data organizer
    data = data_organize(uploaded_file, selected_sheet_name)

    # Adjust figures and thresholds based on the available columns
    available_columns = data.columns.tolist()
    figures = [figure for figure in figures if figure in available_columns]
    thresholds = thresholds[:len(figures)]

    # Execute the filter functions
    result = funds_filter(selected_sheet_name, data, classification, figures, thresholds)

    # Display the filtered results
    st.markdown("#### 篩選結果及其排名(可按住Shift鍵+滾輪滑動表格)")
    # Display the merged dataframe with horizontal scrollbar
    AgGrid(result)

    # Merge the two dataframes based on the "名稱" column
    merged_data = pd.merge(result['名稱'], data, on='名稱')

    # Display the merged dataframe
    st.markdown("#### 篩選結果及其完整數據(可按住Shift鍵+滾輪滑動表格)")
    # Display the merged dataframe with horizontal scrollbar
    AgGrid(merged_data)

    # Export the filtered results to a new csv file (if button is clicked)
    if st.button("將篩選結果輸出至 CSV"):
        # Export the filtered results to a new csv file
        csv_data = merged_data.to_csv(encoding='utf_8_sig', index=False)
        csv_filename = f'篩選結果_{classification}.csv'
        # Display a download button
        st.download_button(label="Download CSV", data=csv_data, file_name=csv_filename, mime='text/csv')
