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
        # Pass the uploaded file directly to the function
        sheet_df_dict, titles_dict = quan_ly_ha_tang_electric(uploaded_file)

        # Lấy các DataFrame từ sheet_df_dict
        df = sheet_df_dict['Dữ liệu nhập'][0]
        dich_vu_dfs = sheet_df_dict['Dịch vụ']
        tong_hop_df = sheet_df_dict['Tổng hợp'][0]
        trung_binh_df = sheet_df_dict['Trung bình'][0]
        tieu_thu_df = sheet_df_dict['Tiêu thụ'][0]

        # Hiển thị dữ liệu theo tab
        tab1, tab2, tab3, tab4, tab5 = st.tabs(["Dữ liệu chính", "Dịch vụ", "Tổng hợp", "Trung bình", "Tiêu thụ"])

        with tab1:
            st.header("Dữ liệu chính")
            st.dataframe(df)

            # Biểu đồ cho dữ liệu chính
            if not df.empty and df.shape[1] > 6:
                st.subheader("Biểu đồ tiêu thụ điện")
                try:
                    fig = px.bar(df.iloc[:10], y="D", title="Top 10 địa điểm tiêu thụ điện")
                    st.plotly_chart(fig)
                except Exception as e:
                    st.write(f"Không thể tạo biểu đồ: {str(e)}")

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

                # Lọc dòng không phải dòng tổng
                service_data = selected_df[selected_df["Stt"] != "Tổng:"].copy()

                if "Thanh toán (KWh)" in service_data.columns:
                    # Chuyển đổi sang số để tránh lỗi
                    service_data["Thanh toán (KWh)"] = pd.to_numeric(service_data["Thanh toán (KWh)"], errors='coerce')
                    service_data = service_data.dropna(subset=["Thanh toán (KWh)"])

                    if not service_data.empty:
                        fig = px.pie(service_data, values="Thanh toán (KWh)", names="Địa chỉ",
                                     title=f"Phân bổ tiêu thụ điện: {selected_service}")
                        st.plotly_chart(fig)
                    else:
                        st.write("Không đủ dữ liệu để hiển thị biểu đồ")
                else:
                    st.write("Không tìm thấy cột 'Thanh toán (KWh)' trong dữ liệu")

            except Exception as e:
                st.write(f"Không thể tạo biểu đồ: {str(e)}")

        with tab3:
            st.header("Tổng hợp")
            st.dataframe(tong_hop_df)

            # Biểu đồ cho tổng hợp
            st.subheader("So sánh tiêu thụ điện theo địa điểm")
            try:
                # Lọc dữ liệu cho biểu đồ, bỏ dòng tổng
                chart_data = tong_hop_df[tong_hop_df["Stt"] < 21].copy()

                # Chuyển đổi sang số để tránh lỗi
                chart_data["Sản lượng tuần mới (kWh)"] = pd.to_numeric(chart_data["Sản lượng tuần mới (kWh)"],
                                                                       errors='coerce')
                chart_data = chart_data.dropna(subset=["Sản lượng tuần mới (kWh)"])

                if not chart_data.empty:
                    fig = px.bar(chart_data, x="Địa chỉ", y="Sản lượng tuần mới (kWh)",
                                 title="Tiêu thụ điện theo địa điểm")
                    st.plotly_chart(fig)
                else:
                    st.write("Không đủ dữ liệu để hiển thị biểu đồ")
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

                # Lấy tất cả cột cấp đầu tiên
                first_level_cols = trung_binh_df.columns.get_level_values(0).unique()

                for col in first_level_cols:
                    # Lấy giá trị TB tuần cho mỗi cột
                    if 'TB tuần' in trung_binh_df[col].columns:
                        # Bỏ hàng cuối cùng (Tổng cộng)
                        tb_tuan_data[col] = trung_binh_df[col]['TB tuần'].iloc[:-1]

                if not tb_tuan_data.empty:
                    fig = px.line(tb_tuan_data, title="Tiêu thụ điện trung bình theo tuần")
                    st.plotly_chart(fig)
                else:
                    st.write("Không có dữ liệu trung bình tuần")
            except Exception as e:
                st.write(f"Không thể tạo biểu đồ trung bình: {str(e)}")

        with tab5:
            st.header("Tiêu thụ")
            st.dataframe(tieu_thu_df)

            # Biểu đồ cho tiêu thụ
            st.subheader("Phân bổ tiêu thụ điện")
            try:
                # Chuyển đổi sang số để tránh lỗi
                tieuthu_data = tieu_thu_df.copy()
                tieuthu_data["Tiêu thụ (KWh)"] = pd.to_numeric(tieuthu_data["Tiêu thụ (KWh)"], errors='coerce')
                tieuthu_data = tieuthu_data.dropna(subset=["Tiêu thụ (KWh)"])

                if not tieuthu_data.empty:
                    fig = px.pie(tieuthu_data, values="Tiêu thụ (KWh)", names="Địa điểm",
                                 title="Phân bổ lượng điện tiêu thụ")
                    st.plotly_chart(fig)
                else:
                    st.write("Không đủ dữ liệu để hiển thị biểu đồ")
            except Exception as e:
                st.write(f"Không thể tạo biểu đồ tiêu thụ: {str(e)}")

        # Tạo nút tải xuống báo cáo
        if st.button("Tạo báo cáo Excel"):
            # Sử dụng tempfile để tạo file tạm thời
            import tempfile
            with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp:
                temp_path = tmp.name
                # Sử dụng hàm export với đường dẫn tạm
                export_excel_formatted_fixed(sheet_df_dict, titles_dict, temp_path)

                # Đọc file để tải xuống
                with open(temp_path, "rb") as f:
                    excel_data = f.read()

                st.download_button(
                    label="Tải xuống báo cáo Excel",
                    data=excel_data,
                    file_name="bao_cao_so_dien.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )


if __name__ == "__main__":
    main()
