import streamlit as st
import pandas as pd
import datetime
import io
import subprocess
import sys

def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

try:
    import openpyxl
except ImportError:
    install('openpyxl')
    import openpyxl

def convert_revenue_to_float(revenue_str):
    return float(revenue_str.replace('USD ', '').replace(',', ''))

def extract_country_name(column_name):
    return column_name.split(':')[-1].strip()

def process_revenue_data(hok_df, country_region):
    hok_revenue_columns = [col for col in hok_df.columns if col not in ['Date', 'Notes']]
    for column in hok_revenue_columns:
        hok_df[column] = hok_df[column].astype(str).apply(convert_revenue_to_float)
    hok_long = hok_df.melt(id_vars=['Date'], value_vars=hok_revenue_columns,
                           var_name='Country', value_name='Revenue')
    hok_long['Country'] = hok_long['Country'].apply(extract_country_name)
    hok_long = hok_long.merge(country_region, left_on='Country', right_on='国家名称', how='left')
    not_found = hok_long[hok_long['所属区域'].isna()]['Country'].unique()
    if len(not_found) > 0:
        st.warning(f"未找到对应区域的国家: {', '.join(not_found)}")
    hok_grouped = hok_long.groupby(['Date', '所属区域']).agg({'Revenue': 'sum'}).reset_index()
    hok_grouped.columns = ['Date', 'Region', 'Gross daily revenue']
    return hok_grouped

def process_units_data(hok_units_df, country_region):
    hok_units_columns = [col for col in hok_units_df.columns if col not in ['Date', 'Notes']]
    for column in hok_units_columns:
        hok_units_df[column] = pd.to_numeric(hok_units_df[column].astype(str).str.replace(',', ''), errors='coerce').fillna(0).astype(int)
    hok_units_long = hok_units_df.melt(id_vars=['Date'], value_vars=hok_units_columns,
                                       var_name='Country', value_name='Units')
    hok_units_long['Country'] = hok_units_long['Country'].apply(extract_country_name)
    hok_units_long = hok_units_long.merge(country_region, left_on='Country', right_on='国家名称', how='left')
    not_found = hok_units_long[hok_units_long['所属区域'].isna()]['Country'].unique()
    if len(not_found) > 0:
        st.warning(f"未找到对应区域的国家: {', '.join(not_found)}")
    hok_units_grouped = hok_units_long.groupby(['Date', '所属区域']).agg({'Units': 'sum'}).reset_index()
    hok_units_grouped.columns = ['Date', 'Region', 'Units']
    return hok_units_grouped

def process_app_store_data(regions_df, sales_df_units, sales_df_revenue):
    # 确认列名是否匹配，并调整列名
    regions_df.columns = ['Country', 'Country Code', 'Region']
    sales_df_units.columns = ['Territory', 'Measure', 'Jan 2023', 'Feb 2023', 'Mar 2023', 'Apr 2023',
                              'May 2023', 'Jun 2023', 'Jul 2023', 'Aug 2023', 'Sep 2023', 'Oct 2023',
                              'Nov 2023', 'Dec 2023', 'Jan 2024', 'Feb 2024', 'Mar 2024', 'Apr 2024',
                              'May 2024', 'Total']
    sales_df_revenue.columns = sales_df_units.columns

    # 将地区对应表转换为字典
    region_dict = pd.Series(regions_df['Region'].values, index=regions_df['Country']).to_dict()

    # 将销售数据中的国家转换为地区
    sales_df_units['Region'] = sales_df_units['Territory'].map(region_dict)
    sales_df_revenue['Region'] = sales_df_revenue['Territory'].map(region_dict)

    # 转换数据格式，将每月的销售数据转换为行
    value_vars = [col for col in sales_df_units.columns if col not in ['Territory', 'Region', 'Measure', 'Total']]
    sales_melted_units = pd.melt(sales_df_units, id_vars=['Territory', 'Region'],
                                 value_vars=value_vars,
                                 var_name='Date', value_name='Units')

    sales_melted_revenue = pd.melt(sales_df_revenue, id_vars=['Territory', 'Region'],
                                   value_vars=value_vars,
                                   var_name='Date', value_name='Revenue')

    # 格式化日期
    sales_melted_units['Date'] = pd.to_datetime(sales_melted_units['Date'], format='%b %Y')
    sales_melted_units['Date'] = sales_melted_units['Date'].dt.strftime('%Y/%m/%d')

    sales_melted_revenue['Date'] = pd.to_datetime(sales_melted_revenue['Date'], format='%b %Y')
    sales_melted_revenue['Date'] = sales_melted_revenue['Date'].dt.strftime('%Y/%m/%d')

    # 确保Units和Revenue列的数据类型为浮点数
    sales_melted_units['Units'] = sales_melted_units['Units'].astype(float)
    sales_melted_revenue['Revenue'] = sales_melted_revenue['Revenue'].astype(float)

    # 按日期和地区汇总销售数据
    sales_summary_units = sales_melted_units.groupby(['Date', 'Region']).sum().reset_index()
    sales_summary_revenue = sales_melted_revenue.groupby(['Date', 'Region']).sum().reset_index()

    # 合并两个表格
    merged_summary = pd.merge(sales_summary_units, sales_summary_revenue, on=['Date', 'Region'], how='outer')

    # 填充缺失值为0
    merged_summary['Units'].fillna(0, inplace=True)
    merged_summary['Revenue'].fillna(0, inplace=True)

    # 添加平台和标题列
    merged_summary['Platform'] = 'IOS'
    merged_summary['Title'] = 'PUBGM'

    # 重新排列列顺序
    final_summary = merged_summary[['Date', 'Title', 'Platform', 'Region', 'Units', 'Revenue']]

    # 创建一个新的 DataFrame 并按日期排序
    sorted_summary = final_summary.sort_values(by=['Date', 'Region']).reset_index(drop=True)

    return sorted_summary

def main():
    st.title('游戏数据处理程序')
    
    # 选择平台
    platform = st.radio("选择平台", ('App Store', 'Google Play'))

    # 用户输入游戏名称
    game_name = st.text_input('请输入游戏名称', 'PUBGM' if platform == 'App Store' else 'HOK')

    # 根据选择的平台提供不同的日期格式提示
    if platform == 'App Store':
        date_format_hint = "请输入日期格式（App Store 通常为 %Y/%m/%d）"
        default_date_format = '%Y/%m/%d'
    else:  # Google Play
        date_format_hint = "请输入日期格式（Google Play 通常为 %b %d, %Y，例如 Jan 01, 2023）"
        default_date_format = '%b %d, %Y'

    # 用户输入日期格式
    date_format = st.text_input(date_format_hint, default_date_format)
    
    if platform == 'App Store':
        regions_file = st.file_uploader('上传国家地区对照表', type='xlsx')
        sales_file_units = st.file_uploader('上传Units销售数据', type='xls')
        sales_file_revenue = st.file_uploader('上传Revenue销售数据', type='xls')
        
        if st.button('处理数据'):
            if regions_file and sales_file_units and sales_file_revenue:
                try:
                    regions_df = pd.read_excel(regions_file, engine='openpyxl')
                    sales_df_units = pd.read_excel(sales_file_units, engine='xlrd')
                    sales_df_revenue = pd.read_excel(sales_file_revenue, engine='xlrd')
                    
                    final_summary = process_app_store_data(regions_df, sales_df_units, sales_df_revenue)
                    st.dataframe(final_summary)
                    
                    # 保存结果到Excel文件
                    output = io.BytesIO()
                    with pd.ExcelWriter(output, engine='openpyxl') as writer:
                        final_summary.to_excel(writer, index=False, float_format="%.10f")
                    output.seek(0)
                    
                    st.download_button(
                        label="下载Excel文件",
                        data=output,
                        file_name=f"{game_name}_汇总表.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                except Exception as e:
                    st.error(f"处理数据时发生错误: {e}")
            else:
                st.warning("请上传所有必要的文件")
    
    else:  # Google Play
        hok_users_files = st.file_uploader('上传 Users 文件', type='csv', accept_multiple_files=True)
        hok_revenue_files = st.file_uploader('上传 Revenue 文件', type='csv', accept_multiple_files=True)
        country_region_file = st.file_uploader('上传 国家地区对照表 文件', type='xlsx')
        
        if st.button('处理数据'):
            try:
                if not hok_users_files or not hok_revenue_files or not country_region_file:
                    st.error("请上传所有必要的文件")
                    return

                all_users_data = []
                all_revenue_data = []
                
                for file in hok_users_files:
                    df = pd.read_csv(file)
                    all_users_data.append(df)
                
                for file in hok_revenue_files:
                    df = pd.read_csv(file)
                    all_revenue_data.append(df)
                
                country_region = pd.read_excel(country_region_file)

                if not all_users_data or not all_revenue_data:
                    st.error("上传的文件中没有数据")
                    return

                # 合并所有用户数据
                hok_users = pd.concat(all_users_data, ignore_index=True)
                
                # 合并所有收入数据
                hok_revenue = pd.concat(all_revenue_data, ignore_index=True)

                # 处理收入和用户数据
                hok_revenue_grouped = process_revenue_data(hok_revenue, country_region)
                hok_units_grouped = process_units_data(hok_users, country_region)

                # 合并收入和用户获取数据
                final_grouped_df = pd.merge(hok_units_grouped, hok_revenue_grouped, on=['Date', 'Region'], how='outer').fillna(0)

                # 添加固定的 Title 和 Platform 列
                final_grouped_df['Title'] = game_name
                final_grouped_df['Platform'] = 'GP'

                # 将列重新排列为所需的顺序
                final_grouped_df = final_grouped_df[['Date', 'Title', 'Platform', 'Region', 'Units', 'Gross daily revenue']]

                # 将日期转换为 datetime 格式进行排序，然后再转换回用户指定的格式
                final_grouped_df['Date'] = pd.to_datetime(final_grouped_df['Date'], format=date_format, errors='coerce')
                final_grouped_df = final_grouped_df.sort_values(by='Date')
                final_grouped_df['Date'] = final_grouped_df['Date'].dt.strftime(date_format)

                # 显示最终数据框
                st.dataframe(final_grouped_df)

                # 提供下载选项
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    final_grouped_df.to_excel(writer, index=False, float_format="%.10f")
                output.seek(0)
                
                st.download_button(
                    label="下载Excel文件",
                    data=output,
                    file_name=f"{game_name}_最终汇总表.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            except Exception as e:
                st.error(f"处理数据时发生错误: {e}")

if __name__ == '__main__':
    main()
