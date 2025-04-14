import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import io
import uuid
from bao_cao_tuan import bao_cao_tuan_electric
from bao_cao_thang import bao_cao_thang_electric
from xuat_file_excel import export_excel_formatted_fixed

# Set page config
st.set_page_config(page_title="Quản Lý Hạ Tầng Kỹ Thuật")


# Function to display dataframes and create visualizations
def display_dataframes(dfs, titles, is_monthly=False):
    for i, df in enumerate(dfs):
        if i < len(titles):
            st.subheader(titles[i])
        else:
            st.subheader(f"Bảng dữ liệu {i + 1}")

        # Display the dataframe with formatting
        st.dataframe(df)

        # Create visualization if appropriate
        if not df.empty and ('Thanh toán (KWh)' in df.columns or not df.select_dtypes(include=[np.number]).empty):
            st.subheader("Biểu đồ")

            # Tạo UUID duy nhất cho mỗi dataframe
            df_uuid = str(uuid.uuid4())
            chart_counter = 0  # Bộ đếm cho các biểu đồ trong dataframe

            try:
                # Different chart types based on data structure
                if 'Sản lượng tuần mới (kWh)' in df.columns:
                    # For summary sheet
                    valid_data = df[~df['Sản lượng tuần mới (kWh)'].isna()].copy()
                    valid_data = valid_data[valid_data['STT'] < 21]  # Loại bỏ hàng Tổng
                    if not valid_data.empty:
                        summary_fig = px.bar(
                            valid_data,
                            y='Sản lượng tuần mới (kWh)',
                            x='Địa chỉ',
                            title='Sản lượng tiêu thụ điện theo địa điểm',
                            color='Địa chỉ',
                            color_discrete_sequence=px.colors.qualitative.Set1,
                            text='Sản lượng tuần mới (kWh)'
                        )
                        summary_fig.update_traces(texttemplate='%{text:.1f}', textposition='outside')
                        chart_key = f"summary_{df_uuid}_{chart_counter}"
                        chart_counter += 1
                        st.plotly_chart(summary_fig, use_container_width=True, key=chart_key)

                elif 'Tiêu thụ (KWh)' in df.columns:
                    # For consumption sheet
                    valid_data = df[df['STT'] < 4].copy()  # Loại bỏ hàng tổng
                    consumption_fig = px.pie(
                        valid_data,
                        values='Tiêu thụ (KWh)',
                        names='Địa điểm',
                        title='Tỷ lệ tiêu thụ điện',
                        color_discrete_sequence=px.colors.qualitative.Set1
                    )
                    chart_key = f"consumption_{df_uuid}_{chart_counter}"
                    chart_counter += 1
                    st.plotly_chart(consumption_fig, use_container_width=True, key=chart_key)

                elif 'Thanh toán (KWh)' in df.columns and 'Địa chỉ' in df.columns:
                    # For service sheets in both weekly and monthly reports
                    try:
                        # Lọc dữ liệu hợp lệ
                        filtered_df = df.copy()

                        # Loại bỏ hàng tổng
                        if 'STT' in filtered_df.columns:
                            filtered_df['STT'] = filtered_df['STT'].astype(str)
                            filtered_df = filtered_df[
                                ~filtered_df['STT'].str.contains('Tổng|tổng|^$', case=False, na=False, regex=True)
                            ]
                            # Thêm lọc cho các giá trị số
                            try:
                                filtered_df['STT'] = pd.to_numeric(filtered_df['STT'], errors='coerce')
                                filtered_df = filtered_df.dropna(subset=['STT'])
                            except:
                                pass

                        # Chuyển đổi sang số
                        filtered_df['Thanh toán (KWh)'] = pd.to_numeric(filtered_df['Thanh toán (KWh)'],
                                                                        errors='coerce')
                        filtered_df = filtered_df.dropna(subset=['Thanh toán (KWh)'])

                        if not filtered_df.empty:
                            # Sắp xếp dữ liệu để đảm bảo thứ tự nhất quán
                            filtered_df = filtered_df.sort_values('Địa chỉ')

                            # Biểu đồ cột
                            bar_fig = px.bar(
                                filtered_df,
                                y='Thanh toán (KWh)',
                                x='Địa chỉ',
                                title='Sản lượng thanh toán theo địa điểm',
                                color='Địa chỉ',
                                color_discrete_sequence=px.colors.qualitative.Set1,
                                text='Thanh toán (KWh)'
                            )
                            bar_fig.update_traces(texttemplate='%{text:.1f}', textposition='outside')

                            bar_key = f"bar_{df_uuid}_{chart_counter}"
                            chart_counter += 1
                            st.plotly_chart(bar_fig, use_container_width=True, key=bar_key)

                            # Thêm biểu đồ tròn cho báo cáo tháng nếu có nhiều giá trị
                            if len(filtered_df) > 1:
                                pie_fig = px.pie(
                                    filtered_df,
                                    values='Thanh toán (KWh)',
                                    names='Địa chỉ',
                                    title='Tỷ lệ tiêu thụ điện theo địa điểm',
                                    color_discrete_sequence=px.colors.qualitative.Set1,
                                    hole=0.4
                                )
                                pie_fig.update_traces(textinfo='percent+label')

                                pie_key = f"pie_{df_uuid}_{chart_counter}"
                                chart_counter += 1
                                st.plotly_chart(pie_fig, use_container_width=True, key=pie_key)
                    except Exception as e:
                        st.warning(f"Lỗi khi tạo biểu đồ cho dữ liệu thanh toán: {str(e)}")

                # Xử lý biểu đồ cho bảng trung bình
                elif isinstance(df.columns, pd.MultiIndex) and 'TB tuần' in df.columns.get_level_values(1):
                    try:
                        # Lấy các cột trung bình tuần
                        tb_tuan_cols = [col[0] for col in df.columns if col[1] == 'TB tuần']

                        if tb_tuan_cols:
                            # Chuẩn bị dữ liệu cho biểu đồ
                            chart_data = df.iloc[:-1].copy()  # Bỏ hàng tổng
                            chart_data = chart_data.xs('TB tuần', level=1, axis=1)
                            chart_data = chart_data.reset_index()

                            # Chuyển đổi dữ liệu
                            melted_data = pd.melt(
                                chart_data,
                                id_vars=['Tên công tơ'],
                                value_vars=tb_tuan_cols,
                                var_name='Thời gian',
                                value_name='Tiêu thụ (KWh)'
                            )

                            avg_fig = px.bar(
                                melted_data,
                                x='Tên công tơ',
                                y='Tiêu thụ (KWh)',
                                color='Thời gian',
                                color_discrete_sequence=px.colors.qualitative.Set1,
                                barmode='group',
                                title='Tiêu thụ trung bình theo tuần'
                            )

                            avg_key = f"avg_{df_uuid}_{chart_counter}"
                            chart_counter += 1
                            st.plotly_chart(avg_fig, use_container_width=True, key=avg_key)
                    except Exception as e:
                        st.warning(f"Lỗi khi tạo biểu đồ trung bình: {str(e)}")

                # Thêm biểu đồ tổng quát cho các bảng khác có dữ liệu số
                elif len(df) > 1:
                    try:
                        # Tìm các cột số
                        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()

                        if len(numeric_cols) > 0:
                            # Tìm một cột phù hợp cho trục x
                            x_col = None
                            for potential_col in ['Địa chỉ', 'Địa điểm', 'Tên công tơ']:
                                if potential_col in df.columns:
                                    x_col = potential_col
                                    break

                            # Không sử dụng STT làm trục x cho biểu đồ

                            if x_col:
                                # Lấy cột số đầu tiên để làm trục y
                                y_col = numeric_cols[0]

                                # Lọc dữ liệu hợp lệ
                                valid_data = df[~df[y_col].isna()].copy()

                                # Loại bỏ hàng tổng nếu có
                                if 'STT' in valid_data.columns:
                                    valid_data = valid_data[
                                        ~valid_data['STT'].astype(str).str.contains('Tổng|tổng|^$', case=False,
                                                                                    na=False, regex=True)
                                    ]

                                if len(valid_data) > 1:
                                    # Sắp xếp dữ liệu
                                    if x_col in valid_data.columns:
                                        valid_data = valid_data.sort_values(x_col)

                                    gen_fig = px.bar(
                                        valid_data,
                                        x=x_col,
                                        y=y_col,
                                        title=f'{y_col} theo {x_col}',
                                        color=x_col,
                                        color_discrete_sequence=px.colors.qualitative.Set1,
                                        text=y_col
                                    )
                                    gen_fig.update_traces(texttemplate='%{text:.1f}', textposition='outside')

                                    gen_key = f"gen_{df_uuid}_{chart_counter}"
                                    chart_counter += 1
                                    st.plotly_chart(gen_fig, use_container_width=True, key=gen_key)
                    except Exception as e:
                        st.warning(f"Lỗi khi tạo biểu đồ: {str(e)}")

            except Exception as e:
                st.warning(f"Không thể tạo biểu đồ: {str(e)}")

        st.markdown("---")


# Tiêu đề chính
st.title("Báo Cáo Tổng Hợp Số Điện TTXVN")

# Main tabs - luôn hiển thị 2 tab chính
tab1, tab2 = st.tabs(["Báo cáo tuần", "Báo cáo tháng"])

# Tab 1: Weekly report
with tab1:
    st.header("Báo cáo tuần")

    # File uploader for weekly report
    uploaded_file_weekly = st.file_uploader("Tải lên file Excel dữ liệu tuần", type=["xlsx", "xls"],
                                            key="weekly_report")

    if uploaded_file_weekly is not None:
        try:
            # Process the file
            sheet_df_dict, titles_dict = bao_cao_tuan_electric(uploaded_file_weekly)

            # Lưu dữ liệu vào session state để dùng khi tạo báo cáo
            st.session_state.weekly_sheet_df_dict = sheet_df_dict
            st.session_state.weekly_titles_dict = titles_dict

            # Create tabs for each sheet in the weekly report
            week_tabs = st.tabs(list(sheet_df_dict.keys()))

            # Display content in each tab
            for i, (sheet_name, tab) in enumerate(zip(sheet_df_dict.keys(), week_tabs)):
                with tab:
                    st.subheader(f"Báo cáo {sheet_name}")
                    display_dataframes(sheet_df_dict[sheet_name], titles_dict.get(sheet_name, []))

            # Thêm nút tạo báo cáo ở cuối
            if st.button("Tạo báo cáo tuần", key="create_weekly_report"):
                # Hiển thị thông báo thành công
                st.success("Đã tạo báo cáo thành công!")

                # Tạo file Excel để tải xuống
                output = io.BytesIO()
                export_excel_formatted_fixed(sheet_df_dict, titles_dict, output)
                output.seek(0)

                # Hiển thị nút tải xuống
                st.download_button(
                    label="Tải xuống báo cáo tuần",
                    data=output.getvalue(),
                    file_name="bao_cao_tuan.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

        except Exception as e:
            st.error(f"Lỗi khi xử lý báo cáo tuần: {str(e)}")

# Tab 2: Monthly report
with tab2:
    st.header("Báo cáo tháng")

    # File uploader for monthly report
    uploaded_file_monthly = st.file_uploader("Tải lên file Excel dữ liệu tháng", type=["xlsx", "xls"],
                                             key="monthly_report")

    if uploaded_file_monthly is not None:
        try:
            # Process the file for monthly report
            sheet_dv_dict, titles_dv_dict, sheet_vna8_dict, titles_vna8_dict = bao_cao_thang_electric(
                uploaded_file_monthly)

            # Lưu dữ liệu vào session state để dùng khi tạo báo cáo
            st.session_state.monthly_sheet_dv_dict = sheet_dv_dict
            st.session_state.monthly_titles_dv_dict = titles_dv_dict
            st.session_state.monthly_sheet_vna8_dict = sheet_vna8_dict
            st.session_state.monthly_titles_vna8_dict = titles_vna8_dict

            # Create container tabs for DV and VNA8
            month_tab1, month_tab2 = st.tabs(["Dịch vụ", "VNA8"])

            # DV tab
            with month_tab1:
                # Create tabs for each sheet in DV dict
                dv_tabs = st.tabs(list(sheet_dv_dict.keys()))

                for i, (sheet_name, tab) in enumerate(zip(sheet_dv_dict.keys(), dv_tabs)):
                    with tab:
                        st.subheader(f"Báo cáo {sheet_name}")
                        display_dataframes(sheet_dv_dict[sheet_name], titles_dv_dict.get(sheet_name, []),
                                           is_monthly=True)

                # Thêm nút tạo báo cáo ở cuối
                if st.button("Tạo báo cáo Dịch vụ", key="create_dv_report"):
                    # Hiển thị thông báo thành công
                    st.success("Đã tạo báo cáo thành công!")

                    # Tạo file Excel để tải xuống
                    dv_output = io.BytesIO()
                    export_excel_formatted_fixed(sheet_dv_dict, titles_dv_dict, dv_output)
                    dv_output.seek(0)

                    # Hiển thị nút tải xuống
                    st.download_button(
                        label="Tải xuống báo cáo Dịch vụ",
                        data=dv_output.getvalue(),
                        file_name="bao_cao_dich_vu.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

            # VNA8 tab
            with month_tab2:
                # Create tabs for each sheet in VNA8 dict
                vna8_tabs = st.tabs(list(sheet_vna8_dict.keys()))

                for i, (sheet_name, tab) in enumerate(zip(sheet_vna8_dict.keys(), vna8_tabs)):
                    with tab:
                        st.subheader(f"Báo cáo {sheet_name}")
                        display_dataframes(sheet_vna8_dict[sheet_name], titles_vna8_dict.get(sheet_name, []),
                                           is_monthly=True)

                # Thêm nút tạo báo cáo ở cuối
                if st.button("Tạo báo cáo VNA8", key="create_vna8_report"):
                    # Hiển thị thông báo thành công
                    st.success("Đã tạo báo cáo thành công!")

                    # Tạo file Excel để tải xuống
                    vna8_output = io.BytesIO()
                    export_excel_formatted_fixed(sheet_vna8_dict, titles_vna8_dict, vna8_output)
                    vna8_output.seek(0)

                    # Hiển thị nút tải xuống
                    st.download_button(
                        label="Tải xuống báo cáo VNA8",
                        data=vna8_output.getvalue(),
                        file_name="bao_cao_vna8.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

        except Exception as e:
            st.error(f"Lỗi khi xử lý báo cáo tháng: {str(e)}")
