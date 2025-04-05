import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
from quan_ly_ha_tang import quan_ly_ha_tang_electric
from quan_ly_ha_tang import export_excel_formatted_fixed


def main():
    st.title("Ứng dụng báo cáo số điện tuần")

    # Upload file Excel
    uploaded_file = st.file_uploader("Tải lên file Excel", type=["xlsx"])

    if uploaded_file is not None:
        # Lưu file tạm thời
        file_path = "temp_data.xlsx"
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Xử lý dữ liệu
        sheet_df_dict, titles_dict = quan_ly_ha_tang_electric(file_path)

        # Lấy các DataFrame từ sheet_df_dict
        df = sheet_df_dict['Dữ liệu nhập'][0]
        dich_vu_dfs = sheet_df_dict['Dịch vụ']
        tong_hop_df = sheet_df_dict['Tổng hợp'][0]
        trung_binh_df = sheet_df_dict['Trung bình'][0]
        tieu_thu_df = sheet_df_dict['Tiêu thụ'][0]

        # Hiển thị dữ liệu theo tab - thêm tab "Tiêu thụ"
        tab1, tab2, tab3, tab4, tab5 = st.tabs(["Dữ liệu chính", "Dịch vụ", "Tổng hợp", "Trung bình", "Tiêu thụ"])

        with tab1:
            st.header("Dữ liệu chính")
            st.dataframe(df)

            # Biểu đồ cho dữ liệu chính
            if not df.empty and df.shape[1] > 6:
                st.subheader("Biểu đồ tiêu thụ điện")
                fig = px.bar(df.iloc[:10], y=df.columns[6], title="Top 10 địa điểm tiêu thụ điện")
                st.plotly_chart(fig)

        with tab2:
            st.header("Dịch vụ")

            # Tạo selectbox để chọn dịch vụ
            service_options = titles_dict['Dịch vụ']
            selected_service = st.selectbox("Chọn dịch vụ:", service_options)

            # Tìm index của dịch vụ được chọn
            service_index = titles_dict['Dịch vụ'].index(selected_service)
            selected_df = dich_vu_dfs[service_index]

            st.dataframe(selected_df)

            # Biểu đồ cho dịch vụ
            try:
                st.subheader(f"Phân bổ tiêu thụ điện: {selected_service}")

                # Lọc dòng tổng để hiển thị trong biểu đồ
                service_data = selected_df[selected_df["Stt"] != "Tổng:"]
                if "Thanh toán (KWh)" in service_data.columns:
                    fig = px.pie(service_data, values="Thanh toán (KWh)", names="Địa chỉ",
                                 title=f"Phân bổ tiêu thụ điện: {selected_service}")
                    st.plotly_chart(fig)
            except Exception as e:
                st.write(f"Không thể tạo biểu đồ: {str(e)}")

        with tab3:
            st.header("Tổng hợp")
            st.dataframe(tong_hop_df)

            # Biểu đồ cho tổng hợp
            st.subheader("So sánh tiêu thụ điện theo địa điểm")
            try:
                # Lấy dữ liệu không bao gồm dòng tổng
                chart_data = tong_hop_df.iloc[:-1]
                fig = px.bar(chart_data, x="Địa chỉ", y="Sản lượng tuần mới (kWh)",
                             title="Tiêu thụ điện theo địa điểm")
                st.plotly_chart(fig)
            except Exception as e:
                st.write(f"Không thể tạo biểu đồ tổng hợp: {str(e)}")

        with tab4:
            st.header("Trung bình")
            st.dataframe(trung_binh_df)

            # Biểu đồ cho trung bình
            st.subheader("Tiêu thụ điện trung bình theo tuần")
            try:
                # Lấy dữ liệu TB tuần từ MultiIndex DataFrame
                tb_tuan_data = pd.DataFrame()
                for col in trung_binh_df.columns.levels[0]:
                    if 'TB tuần' in trung_binh_df[col].columns:
                        tb_tuan_data[col] = trung_binh_df[col]['TB tuần']

                fig = px.line(tb_tuan_data.iloc[:-1], title="Tiêu thụ điện trung bình theo tuần")
                st.plotly_chart(fig)
            except Exception as e:
                st.write(f"Không thể tạo biểu đồ trung bình: {str(e)}")

        with tab5:
            st.header("Tiêu thụ")
            st.dataframe(tieu_thu_df)

            # Biểu đồ cho tiêu thụ
            st.subheader("Phân bổ tiêu thụ điện")
            try:
                if "Tiêu thụ (KWh)" in tieu_thu_df.columns:
                    fig = px.pie(tieu_thu_df, values="Tiêu thụ (KWh)", names="Địa điểm",
                                 title="Phân bổ lượng điện tiêu thụ")
                    st.plotly_chart(fig)
            except Exception as e:
                st.write(f"Không thể tạo biểu đồ tiêu thụ: {str(e)}")

        # Tạo nút tải xuống báo cáo sử dụng hàm mới
        if st.button("Tạo báo cáo Excel"):
            output_path = "bao_cao_so_dien.xlsx"

            # Sử dụng hàm export_excel_formatted_fixed với cấu trúc mới
            export_excel_formatted_fixed(sheet_df_dict, titles_dict, output_path)

            st.success("Đã tạo báo cáo Excel thành công! File: bao_cao_so_dien.xlsx")

            with open(output_path, "rb") as file:
                st.download_button(
                    label="Tải xuống báo cáo Excel",
                    data=file,
                    file_name="bao_cao_so_dien.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )


if __name__ == "__main__":
    main()
