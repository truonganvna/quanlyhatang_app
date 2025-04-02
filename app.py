import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import plotly.express as px
from quan_ly_ha_tang import quan_ly_ha_tang_electric


def main():
    st.title("Ứng dụng Quản lý Hạ tầng Điện")

    # Upload file Excel
    uploaded_file = st.file_uploader("Tải lên file Excel", type=["xlsx"])

    if uploaded_file is not None:
        # Lưu file tạm thời
        file_path = "temp_data.xlsx"
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())

        # Xử lý dữ liệu
        df, dich_vu_sheet_frame, tong_hop_sheet_frame, trung_binh_sheet_frame = quan_ly_ha_tang_electric(file_path)

        # Hiển thị dữ liệu theo tab
        tab1, tab2, tab3, tab4 = st.tabs(["Dữ liệu chính", "Dịch vụ", "Tổng hợp", "Trung bình"])

        with tab1:
            st.header("Dữ liệu chính")
            st.dataframe(df)

            # Biểu đồ cho dữ liệu chính
            st.subheader("Biểu đồ tiêu thụ điện")
            fig = px.bar(df.iloc[:10], y=df.columns[6], title="Top 10 địa điểm tiêu thụ điện")
            st.plotly_chart(fig)

        with tab2:
            st.header("Dịch vụ")
            st.dataframe(dich_vu_sheet_frame)

            # Biểu đồ cho dịch vụ
            st.subheader("Phân bổ tiêu thụ điện theo dịch vụ")
            service_data = dich_vu_sheet_frame[dich_vu_sheet_frame["Stt"] == "Tổng:"]
            fig = px.pie(service_data, values="Thanh toán (KWh)", names=service_data.index,
                         title="Phân bổ tiêu thụ điện theo dịch vụ")
            st.plotly_chart(fig)

        with tab3:
            st.header("Tổng hợp")
            st.dataframe(tong_hop_sheet_frame)

            # Biểu đồ cho tổng hợp
            st.subheader("So sánh tiêu thụ điện theo địa điểm")
            fig = px.bar(tong_hop_sheet_frame.iloc[1:21], x="Địa chỉ", y="Sản lượng tuần mới (kWh)",
                         title="Tiêu thụ điện theo địa điểm")
            st.plotly_chart(fig)

        with tab4:
            st.header("Trung bình")
            st.dataframe(trung_binh_sheet_frame)

            # Biểu đồ cho trung bình
            st.subheader("Tiêu thụ điện trung bình theo tuần")
            # Lấy dữ liệu TB tuần từ các cột
            try:
                tb_tuan_data = trung_binh_sheet_frame.xs('TB tuần', level=1, axis=1).iloc[:6]
                fig = px.line(tb_tuan_data, title="Tiêu thụ điện trung bình theo tuần")
                st.plotly_chart(fig)
            except:
                st.write("Không thể tạo biểu đồ cho dữ liệu trung bình")

        # Tạo nút tải xuống báo cáo
        from quan_ly_ha_tang import export_excel_file

        if st.button("Tạo báo cáo Excel"):
            export_excel_file(df, dich_vu_sheet_frame, tong_hop_sheet_frame, trung_binh_sheet_frame)
            st.success("Đã tạo báo cáo Excel thành công! File: bao_cao_so_dien.xlsx")

            with open("bao_cao_so_dien.xlsx", "rb") as file:
                st.download_button(
                    label="Tải xuống báo cáo Excel",
                    data=file,
                    file_name="bao_cao_so_dien.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )


if __name__ == "__main__":
    main()
