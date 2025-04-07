import pandas as pd

def bao_cao_thang_electric(file_input):
    df = pd.read_excel(file_input, skiprows=2).set_index('Địa chỉ')

    df_mb1 = df.loc[["ĐH-Outdoor T3 (MB)", "ĐH-Outdoor T2 (MB)", "ĐH-Indoor T3+T2 (MB)", "ĐH-Outdoor T1 (MB+TTX)", "ĐH-Indoor T1 (MB)", "ĐH-Indoor T1 (TTX)"]]
    df_mb1_frame = pd.DataFrame({
        "Stt": [1, 2, 3, 4, 5, 6, ""],
        "Số công tơ": [12068239, 12063634, 12063346, 12068229, 13011012, 13011297, "Tổng"],
        "Loại c tơ": ["3 pha", "3 pha", "3 pha", "3 pha", "1 pha", "1 pha", ""],
        "Địa chỉ": ["ĐH tầng 3 (Outdoor)", "ĐH tầng 2 (Outdoor)", "ĐH tầng 3+2(Indoor)", "ĐH tầng 1 (Outdoor)", "ĐH tầng1 (Indoor)-MB", "ĐH tầng 1 (Indoor)-TTX", ""],
        "CSCK": [""]*7,
        "CSĐK": [""]*7,
        "Hệ số": [1, 1, 1, 1, 1, 1, ""],
        "Tổng KWh": [""]*7,
        "%": [""]*7,
        "Thanh toán (KWh)": [""]*7
    })
    df_mb1_frame.loc[:5, "CSĐK"] = df_mb1["28/02"].values
    df_mb1_frame.loc[:5, "CSCK"] = df_mb1["31/03"].values
    df_mb1_frame.loc[:5, "Tổng KWh"] = [(df_mb1_frame.loc[i, "CSCK"] - df_mb1_frame.loc[i, "CSĐK"]) * df_mb1_frame.loc[i, "Hệ số"] for i in range(6)]
    df_mb1_frame.loc[3, "%"] = round(df_mb1_frame.loc[4, "Tổng KWh"]/(df_mb1_frame.loc[4, "Tổng KWh"] + df_mb1_frame.loc[5, "Tổng KWh"]), 2)
    df_mb1_frame.loc[:4, "Thanh toán (KWh)"] = df_mb1_frame.loc[:5, "Tổng KWh"]
    df_mb1_frame.loc[3, "Thanh toán (KWh)"] = df_mb1_frame.loc[3, "Tổng KWh"] * df_mb1_frame.loc[3, "%"]
    df_mb1_frame.loc[6, "Thanh toán (KWh)"] = sum(df_mb1_frame.loc[:4, "Thanh toán (KWh)"])


    df_mb2 = df.loc[["AS + ĐL T3+T2+T1 (MB)"]]
    df_mb2_frame = pd.DataFrame({
        "Stt": [1, ""],
        "Số công tơ": [403189, "Tổng"],
        "Loại c tơ": ["3 pha", ""],
        "Địa chỉ": ["AS + ĐL tầng 1,2,3", ""],
        "CSCK": [""]*2,
        "CSĐK": [""]*2,
        "Hệ số": [1, ""],
        "Tổng KWh": [""]*2,
        "%": [""]*2,
        "Thanh toán (KWh)": [""]*2
    })
    df_mb2_frame.loc[:0,'CSĐK'] = df_mb2['28/02'].values
    df_mb2_frame.loc[:0,'CSCK'] = df_mb2['31/03'].values
    df_mb2_frame.loc[:0,'Tổng KWh'] = (df_mb2_frame.loc[0,'CSCK'] - df_mb2_frame.loc[0,'CSĐK']) * df_mb2_frame.loc[0,'Hệ số']
    df_mb2_frame.loc[:1,'Thanh toán (KWh)'] = df_mb2_frame.loc[0,'Tổng KWh']


    df_giovani1 = df.loc[["Outdoor 1-T1", "Outdoor 2-T1", "Indoor-GIOVANI", "Indoor- TTX"]]
    df_giovani1_frame = pd.DataFrame({
        "Stt": [1, 2, "", 3, 4, ""],
        "Số công tơ": [14038150, 14038145, "", 15009568, 15012663, "Tổng"],
        "Loại c tơ": ["3 pha", "3 pha", "", "1 pha", "1 pha", ""],
        "Địa chỉ": [
            "ĐH-Outdoor M1", "ĐH-Outdoor M2", "",
            "ĐH-Indoor N1 (GVN)", "ĐH-Indoor N2 (TTX)", ""
        ],
        "CSCK": ["", "", "", "", "", ""],
        "CSĐK": ["", "", "", "", "", ""],
        "Hệ số": [1, 1, "", 1, 1, ""],
        "Tổng KWh": ["", "", "", "", "", ""],
        "%": ["", "", "", "", "", ""],
        "Thanh toán (KWh)": ["", "", "", "", "", ""]
    })
    df_giovani1_frame.loc[:1, "CSĐK"] = df_giovani1["28/02"].iloc[:2].values
    df_giovani1_frame.loc[3:4, "CSĐK"] = df_giovani1["28/02"].iloc[2:4].values

    df_giovani1_frame.loc[:1, "CSCK"] = df_giovani1["31/03"].iloc[:2].values
    df_giovani1_frame.loc[3:4, "CSCK"] = df_giovani1["31/03"].iloc[2:4].values
    df_giovani1_frame.loc[:1, "Tổng KWh"] = (df_giovani1_frame.loc[:1, "CSCK"] - df_giovani1_frame.loc[:1, "CSĐK"]) * df_giovani1_frame.loc[:1, "Hệ số"]
    df_giovani1_frame.loc[3:4, "Tổng KWh"] = (df_giovani1_frame.loc[3:4, "CSCK"] - df_giovani1_frame.loc[3:4, "CSĐK"]) * df_giovani1_frame.loc[3:4, "Hệ số"]
    df_giovani1_frame.loc[2, "%"] = round(df_giovani1_frame.loc[3, "Tổng KWh"]/(df_giovani1_frame.loc[3, "Tổng KWh"] + df_giovani1_frame.loc[4, "Tổng KWh"]), 2)
    df_giovani1_frame.loc[2, "Thanh toán (KWh)"] = round(sum(df_giovani1_frame.loc[:1, "Tổng KWh"]) * df_giovani1_frame.loc[2, "%"], 1)
    df_giovani1_frame.loc[3, "Thanh toán (KWh)"] = df_giovani1_frame.loc[3, "Tổng KWh"]
    df_giovani1_frame.loc[5, "Thanh toán (KWh)"] = pd.to_numeric(df_giovani1_frame.loc[:4, "Thanh toán (KWh)"], errors="coerce").sum()


    df_giovani2 = df.loc[["AS+ĐL - GIOVANI"]]
    df_giovani2_frame = pd.DataFrame({
        "Stt": [1, ""],
        "Số công tơ": [10511904, "Tổng"],
        "Loại c tơ": ["3 pha", ""],
        "Địa chỉ": ["AS + ĐL DVT1", ""],
        "CSCK": [""]*2,
        "CSĐK": [""]*2,
        "Hệ số": [1, ""],
        "Tổng KWh": [""]*2,
        "%": [""]*2,
        "Thanh toán (KWh)": [""]*2
    })
    df_giovani2_frame.loc[:0,'CSĐK'] = df_giovani2['28/02'].values
    df_giovani2_frame.loc[:0,'CSCK'] = df_giovani2['31/03'].values
    df_giovani2_frame.loc[:0,'Tổng KWh'] = (df_giovani2_frame.loc[0,'CSCK'] - df_giovani2_frame.loc[0,'CSĐK']) * df_giovani2_frame.loc[0,'Hệ số']
    df_giovani2_frame.loc[:1,'Thanh toán (KWh)'] = df_giovani2_frame.loc[0,'Tổng KWh']


    df_gme = df.loc[["AS + ĐL T4 (GME)", "ĐH-Outdoor T4 (GME +TTX)", "ĐH-Indoor T4 (GME)", "ĐH-Indoor T4 (TTX)"]]
    df_gme_frame = pd.DataFrame({
        "Stt": [4, 1, 3, 2, ""],
        "Số công tơ": [16698180, 16705013, 16702810, 16702656, "Tổng"],
        "Loại c tơ": ["1 pha", "3 pha", "1 pha", "1 pha", ""],
        "Địa chỉ": [
            "AS + ĐL T4 (GME)",
            "ĐH-Outdoor T4",
            "ĐH-Indoor T4 (GME)",
            "ĐH-Indoor T4 (TTX)",
            ""
        ],
        "CSCK": [""] * 5,
        "CSĐK": [""] * 5,
        "Hệ số": [1, 1, 1, 1, ""],
        "Tổng KWh": [""] * 5,
        "%": [""] * 5,
        "Thanh toán (KWh)": [""] * 5
    })
    df_gme_frame.loc[:3, "CSĐK"] = df_gme["28/02"].values
    df_gme_frame.loc[:3, "CSCK"] = df_gme["31/03"].values
    df_gme_frame.loc[:3, 'Tổng KWh'] = (df_gme_frame.loc[:3, 'CSCK'] - df_gme_frame.loc[:3, 'CSĐK']) * df_gme_frame.loc[:3, 'Hệ số']
    df_gme_frame.loc[1, "%"] = round(df_gme_frame.loc[2, "Tổng KWh"]/(df_gme_frame.loc[2, "Tổng KWh"] + df_gme_frame.loc[3, "Tổng KWh"]), 2)

    df_gme_frame.loc[0, "Thanh toán (KWh)"] = df_gme_frame.loc[0, "Tổng KWh"]
    df_gme_frame.loc[1, "Thanh toán (KWh)"] = df_gme_frame.loc[1, "Tổng KWh"] * df_gme_frame.loc[1, "%"]
    df_gme_frame.loc[2, "Thanh toán (KWh)"] = df_gme_frame.loc[2, "Tổng KWh"]

    df_gme_frame.loc[4, "Thanh toán (KWh)"] = sum(df_gme_frame.loc[:2, "Thanh toán (KWh)"])

    df_baoviet = df.loc[["Bảo Việt T5"]]
    df_baoviet_frame = pd.DataFrame({
        "Stt": [1, ""],
        "Số công tơ": [12068236, "Tổng"],
        "Loại c tơ": ["3 pha", ""],
        "Địa chỉ": ["ĐH-Outdoor T5", ""],
        "CSCK": ["", ""],
        "CSĐK": ["", ""],
        "Hệ số": [1, ""],
        "Tổng KWh": ["", ""],
        "%": ["", ""],
        "Thanh toán (KWh)": ["", ""]
    })

    df_baoviet_frame.loc[:0, "CSĐK"] = df_baoviet["28/02"].values
    df_baoviet_frame.loc[:0, "CSCK"] = df_baoviet["31/03"].values
    df_baoviet_frame.loc[:0, "Tổng KWh"] = (df_baoviet_frame.loc[0, "CSCK"] - df_baoviet_frame.loc[0, "CSĐK"]) * df_baoviet_frame.loc[0, "Hệ số"]
    df_baoviet_frame.loc[0, "Thanh toán (KWh)"] = df_baoviet_frame.loc[0, "Tổng KWh"]
    df_baoviet_frame.loc[1, "Thanh toán (KWh)"] = sum(df_baoviet_frame.loc[:0, "Thanh toán (KWh)"])

    df_fbs = df.loc[["AS+ĐL FSB T12", "ĐH-Outdoor T12 (Mới)", "ĐH-Indoor FSB (Mới)"]]
    df_fsb_frame = pd.DataFrame({
        "Stt": [1, 2, 3, ""],
        "Số công tơ": [12068236, 16719586, "09048549", "Tổng"],
        "Loại công tơ": ["3 pha", "1 pha", "1 pha", ""],
        "Địa chỉ": ["AS+ĐL FSB T12", "ĐH-Outdoor T12 (Mới)", "ĐH-Indoor FSB (Mới)", ""],
        "CSCK": [""]*4,
        "CSĐK": [""]*4,
        "Hệ số": [1, 1, 1, ""],
        "Tổng KWh": [""]*4,
        "%": ["", "", "", ""],
        "Thanh toán (KWh)": [""]*4
    })
    df_fsb_frame.loc[:2, "CSĐK"] = df_fbs["28/02"].values
    df_fsb_frame.loc[:2, "CSCK"] = df_fbs["31/03"].values
    df_fsb_frame.loc[:2, "Tổng KWh"] = (df_fsb_frame.loc[:2, "CSCK"] - df_fsb_frame.loc[:2, "CSĐK"]) * df_fsb_frame.loc[:2, "Hệ số"]
    df_fsb_frame.loc[:2, "Thanh toán (KWh)"] = df_fsb_frame.loc[:2, "Tổng KWh"]
    df_fsb_frame.loc[3, "Thanh toán (KWh)"] = sum(df_fsb_frame.loc[:2, "Thanh toán (KWh)"])

    df_fbs_old_frame = pd.DataFrame({
        "Stt": [1, 2, 3, ""],
        "Số công tơ": [14047859, 17736901, 14520204, "Tổng"],
        "Loại c tơ": ["3 pha", "1 pha", "1 pha", ""],
        "Địa chỉ": ["ĐH-Outdoor T12 (Cũ)", "ĐH-Outdoor T12 (Cũ)", "ĐH-Outdoor T12 (Cũ)", ""],
        "CSCK": [""]*4,
        "CSĐK": [""]*4,
        "Hệ số": [1, 1, 1, ""],
        "Tổng KWh": [""]*4,
        "%": [""]*4,
        "Thanh toán (KWh)": [""]*4
    })

    df_fbs_old = df.loc[["ĐH-Outdoor T12 (Cũ)", "ĐH-Indoor T12 (FSB-Cũ)", "ĐH-Indoor T12 (TTX-Cũ)"]]
    df_fbs_old_frame.loc[:2, "CSĐK"] = df_fbs_old["28/02"].values
    df_fbs_old_frame.loc[:2, "CSCK"] = df_fbs_old["31/03"].values
    df_fbs_old_frame.loc[:2, "Tổng KWh"] = (df_fbs_old_frame.loc[:2, "CSCK"] - df_fbs_old_frame.loc[:2, "CSĐK"]) * df_fbs_old_frame.loc[:2, "Hệ số"]
    df_fbs_old_frame.loc[1, "%"] = round(df_fbs_old_frame.loc[1, "Tổng KWh"]/(df_fbs_old_frame.loc[1, "Tổng KWh"] + df_fbs_old_frame.loc[2, "Tổng KWh"]), 2)
    df_fbs_old_frame.loc[0, "Thanh toán (KWh)"] = round(df_fbs_old_frame.loc[0, "Tổng KWh"] * df_fbs_old_frame.loc[1, "%"], 1)
    df_fbs_old_frame.loc[1, "Thanh toán (KWh)"] = round(df_fbs_old_frame.loc[1, "Tổng KWh"], 1)
    df_fbs_old_frame.loc[3, "Thanh toán (KWh)"] = sum(df_fbs_old_frame.loc[:1, "Thanh toán (KWh)"])

    # Bảng Trung tâm hợp tác Quốc tế Thông tấn phần văn phòng
    df_dv8_vp = df.loc[["AS+ĐL 8THĐ", "ĐH 8THĐ", "33LTT", "AS18THĐ", "ĐH18THĐ", "AS20THĐ", "ĐH20THĐ", "AS-T6-11THĐ", "ĐH-T6-11THĐ", "AS-P604-11THĐ", "ĐH-P604-11THĐ", "AS T5.11THĐ",
                        "ĐH T5.11THĐ", "ĐH-T7 P701-703-705", "AS-T7 P701-703-705", "ĐH-T7 P702-704-706", "AS-T7 P702-704-706", "ĐH-T7 P707", "AS-T7 P707", "ĐH-T7 P708", "AS-T7 P708"]]


    df_dv8_vp_frame = pd.DataFrame({
        "Stt": [1, "", "", "", 2, "2.1", "", "", "", "", "2.2", "", "", "", "", "", "", "2.3", "", "", "", "", "", "", "", "", 3, ""],
        "Số công tơ": ["", 14308090, 14308130, 372755, "CT mới", "", 15485, 15485, 15485, 15485, "", "", "", "", "", "", "", "", 30094, 37553, 30098, 30138, 13579, 52583, 13464, 277745, 99178786, "Tổng"],
        "Loại công tơ": ["", "3 pha", "3 pha", "3 pha", "3 pha", "", "3 pha", "3 pha", "3 pha", "3 pha", "", "3 pha", "3 pha", "1 pha", "1 pha", "3 pha", "3 pha", "", "3 pha", "3 pha", "3 pha", "3 pha", "1 pha", "1 pha", "1 pha", "1 pha", "3 pha", ""],
        "Địa chỉ": ["Khu vực số 8THĐ", "AS 8THĐ (từ 5LTK)", "ĐH 8THĐ (từ 5LTK)", "Số 33LTT (từ 8THĐ)", "BA - 630KVA", "18-20 THĐ", "AS18. THĐ", "ĐH18. THĐ", "AS20. THĐ", "ĐH20. THĐ",
                    "T6.11.THĐ Sử dụng", "AS T6.11THĐ", "ĐH T6.11THĐ", "AS P604.11THĐ", "ĐH P604.11THĐ", "AS T5.11THĐ", "ĐH T5.11THĐ", "T7.11.THĐ Sử dụng", "ĐH.P701-3-5", "AS.P701-3-5",
                    "ĐH.P702-4-6", "AS.P702-4-6", "ĐH. P707", "AS. P707", "ĐH. P708", "AS. P708", "Nguồn của Sphon", ""],
        "CSCK": [""] * 28,
        "CSĐK": [""] * 28,
        "Hệ số": [""] * 28,
        "Tổng KWh": [""] * 28,
        "%": [""] * 28,
        "Thanh toán (KWh)": [""] * 28
    })

    df_dv8_vp_frame.loc[4, "Hệ số"] = 200
    df_dv8_vp_frame.loc[26, "Hệ số"] = 20

    df_dv8_vp_frame.loc[1:3, "CSĐK"] = df_dv8_vp.loc["AS+ĐL 8THĐ":"33LTT", "28/02"].values
    df_dv8_vp_frame.loc[1:3, "CSCK"] = df_dv8_vp.loc["AS+ĐL 8THĐ":"33LTT", "31/03"].values
    df_dv8_vp_frame.loc[1:3, "Hệ số"] = df_dv8_vp.loc["AS+ĐL 8THĐ":"33LTT", "Hệ số"].values
    df_dv8_vp_frame.loc[1:3, "Tổng KWh"] = (df_dv8_vp_frame.loc[1:3, "CSCK"] - df_dv8_vp_frame.loc[1:3, "CSĐK"]) * df_dv8_vp_frame.loc[1:3, "Hệ số"]
    df_dv8_vp_frame.loc[1, "Thanh toán (KWh)"] = df_dv8_vp_frame.loc[1:3, "Tổng KWh"].sum()

    df_dv8_vp_frame.loc[4, "CSĐK"] = df.loc["TBA-11THĐ 630KVA-10/0.4KV", "28/02"]
    df_dv8_vp_frame.loc[4, "CSCK"] = df.loc["TBA-11THĐ 630KVA-10/0.4KV", "31/03"]
    df_dv8_vp_frame.loc[4, "Tổng KWh"] = (df_dv8_vp_frame.loc[4, "CSCK"] - df_dv8_vp_frame.loc[4, "CSĐK"]) * df_dv8_vp_frame.loc[4, "Hệ số"]

    df_dv8_vp_frame.loc[6:9, "CSĐK"] = df_dv8_vp.loc["AS18THĐ":"ĐH20THĐ", "28/02"].values
    df_dv8_vp_frame.loc[6:9, "CSCK"] = df_dv8_vp.loc["AS18THĐ":"ĐH20THĐ", "31/03"].values
    df_dv8_vp_frame.loc[6:9, "Hệ số"] = df_dv8_vp.loc["AS18THĐ":"ĐH20THĐ", "Hệ số"].values
    df_dv8_vp_frame.loc[6:9, "Tổng KWh"] = (df_dv8_vp_frame.loc[6:9, "CSCK"] - df_dv8_vp_frame.loc[6:9, "CSĐK"]) * df_dv8_vp_frame.loc[6:9, "Hệ số"]
    df_dv8_vp_frame.loc[6, "Thanh toán (KWh)"] = df_dv8_vp_frame.loc[6:9, "Tổng KWh"].sum()

    df_dv8_vp_frame.loc[11:16, "CSĐK"] = df_dv8_vp.loc["AS-T6-11THĐ":"ĐH T5.11THĐ", "28/02"].values
    df_dv8_vp_frame.loc[11:16, "CSCK"] = df_dv8_vp.loc["AS-T6-11THĐ":"ĐH T5.11THĐ", "31/03"].values
    df_dv8_vp_frame.loc[11:16, "Hệ số"] = df_dv8_vp.loc["AS-T6-11THĐ":"ĐH T5.11THĐ", "Hệ số"].values
    df_dv8_vp_frame.loc[11:16, "Tổng KWh"] = (df_dv8_vp_frame.loc[11:16, "CSCK"] - df_dv8_vp_frame.loc[11:16, "CSĐK"]) * df_dv8_vp_frame.loc[11:16, "Hệ số"]
    df_dv8_vp_frame.loc[11, "Thanh toán (KWh)"] = df_dv8_vp_frame.loc[11:16, "Tổng KWh"].sum()

    df_dv8_vp_frame.loc[18:25, "CSĐK"] = df_dv8_vp.loc["ĐH-T7 P701-703-705":"AS-T7 P708", "28/02"].values
    df_dv8_vp_frame.loc[18:25, "CSCK"] = df_dv8_vp.loc["ĐH-T7 P701-703-705":"AS-T7 P708", "31/03"].values
    df_dv8_vp_frame.loc[18:25, "Hệ số"] = df_dv8_vp.loc["ĐH-T7 P701-703-705":"AS-T7 P708", "Hệ số"].values
    df_dv8_vp_frame.loc[18:25, "Tổng KWh"] = (df_dv8_vp_frame.loc[18:25, "CSCK"] - df_dv8_vp_frame.loc[18:25, "CSĐK"]) * df_dv8_vp_frame.loc[18:25, "Hệ số"]
    df_dv8_vp_frame.loc[18, "Thanh toán (KWh)"] = df_dv8_vp_frame.loc[18:25, "Tổng KWh"].sum()


    df_dv8_vp_frame_total_KWh = df_dv8_vp_frame.loc[1:2, "Tổng KWh"].sum() + df_dv8_vp_frame.loc[6:9, "Tổng KWh"].sum() + df_dv8_vp_frame.loc[11:12, "Tổng KWh"].sum() + \
                                df_dv8_vp_frame.loc[18:25, "Tổng KWh"].sum() - df_dv8_vp_frame.loc[3, "Tổng KWh"] - df_dv8_vp_frame.loc[13:14, "Tổng KWh"].sum()
    df_dv8_vp_frame_total_thanhtoan = df_dv8_vp_frame.loc[1:3, "Tổng KWh"].sum() + df_dv8_vp_frame.loc[6:9, "Tổng KWh"].sum() + df_dv8_vp_frame.loc[11:16, "Tổng KWh"].sum() + df_dv8_vp_frame.loc[18:25, "Tổng KWh"].sum()

    df_dv8_vp_frame.loc[27, "Tổng KWh"] = df_dv8_vp_frame_total_KWh
    df_dv8_vp_frame.loc[27, "Thanh toán (KWh)"] = df_dv8_vp_frame_total_thanhtoan


    # Bảng Trung tâm hợp tác Quốc tế Thông tấn phần PG Bank
    df_dv8_pgb = df.loc[["ĐH  -T1 (11THĐ)", "AS - T1 (11THĐ)", "ĐH  -T2 (11THĐ)", "AS - T2 (11THĐ)"]]


    df_dv8_pgb_frame = pd.DataFrame({
        "Stt": [1, 2, 3, 4, ""],
        "Số công tơ": [12067352, 13012746, 13010295, 13010298, "Tổng"],
        "Loại công tơ": ["1 pha", "1 pha", "1 pha", "1 pha", ""],
        "Địa chỉ": ["ĐH -T1 (11THĐ)", "AS - T1 (11THĐ)", "ĐH -T2 (11THĐ)", "AS - T2 (11THĐ)", ""],
        "CSCK": [""]*5,
        "CSĐK": [""]*5,
        "Hệ số": [""]*5,
        "Tổng KWh": [""]*5,
        "%": [""]*5,
        "Thanh toán (KWh)": [""]*5,
    })

    df_dv8_pgb_frame.loc[:3, "CSĐK"] = df_dv8_pgb.loc["ĐH  -T1 (11THĐ)":"AS - T2 (11THĐ)", "28/02"].values
    df_dv8_pgb_frame.loc[:3, "CSCK"] = df_dv8_pgb.loc["ĐH  -T1 (11THĐ)":"AS - T2 (11THĐ)", "31/03"].values
    df_dv8_pgb_frame.loc[:3, "Hệ số"] = df_dv8_pgb.loc["ĐH  -T1 (11THĐ)":"AS - T2 (11THĐ)", "Hệ số"].values
    df_dv8_pgb_frame.loc[:3, "Tổng KWh"] = (df_dv8_pgb_frame.loc[:3, "CSCK"] - df_dv8_pgb_frame.loc[:3, "CSĐK"]) * df_dv8_pgb_frame.loc[:3, "Hệ số"]
    df_dv8_pgb_frame.loc[:3, "Thanh toán (KWh)"] = df_dv8_pgb_frame.loc[:3, "Tổng KWh"]
    df_dv8_pgb_frame.loc[4, "Thanh toán (KWh)"] = sum(df_dv8_pgb_frame.loc[:3, "Thanh toán (KWh)"])

    # Bảng Trung tâm hợp tác Quốc tế Thông tấn phần văn phòng
    df_vna8_vp = df.loc[["AS+ĐL 8THĐ", "ĐH 8THĐ", "33LTT", "AS18THĐ", "ĐH18THĐ", "AS20THĐ", "ĐH20THĐ", "AS-T6-11THĐ", "ĐH-T6-11THĐ", "AS-P604-11THĐ", "ĐH-P604-11THĐ", "AS T5.11THĐ",
                        "ĐH T5.11THĐ", "ĐH-T7 P701-703-705", "AS-T7 P701-703-705", "ĐH-T7 P702-704-706", "AS-T7 P702-704-706", "ĐH-T7 P707", "AS-T7 P707", "ĐH-T7 P708", "AS-T7 P708"]]

    df_vna8_vp_frame = pd.DataFrame({
        "Stt": [1, "", "", "", 2, "2.1", "", "", "", "", "2.2", "", "", "", "", "", "", "2.3", "", "", "", "", "", "", "", "", 3, ""],
        "Số công tơ": ["", 14308090, 14308130, 372755, "CT mới", "", 15485, 15485, 15485, 15485, "", "", "", "", "", "", "", "", 30094, 37553, 30098, 30138, 13579, 52583, 13464, 277745, 99178786, "Tổng"],
        "Loại công tơ": ["", "3 pha", "3 pha", "3 pha", "3 pha", "", "3 pha", "3 pha", "3 pha", "3 pha", "", "3 pha", "3 pha", "1 pha", "1 pha", "3 pha", "3 pha", "", "3 pha", "3 pha", "3 pha", "3 pha", "1 pha", "1 pha", "1 pha", "1 pha", "3 pha", ""],
        "Địa chỉ": ["Khu vực số 8THĐ", "AS 8THĐ (từ 5LTK)", "ĐH 8THĐ (từ 5LTK)", "Số 33LTT (từ 8THĐ)", "BA - 630KVA", "18-20 THĐ", "AS18. THĐ", "ĐH18. THĐ", "AS20. THĐ", "ĐH20. THĐ",
                    "T6.11.THĐ Sử dụng", "AS T6.11THĐ", "ĐH T6.11THĐ", "AS P604.11THĐ", "ĐH P604.11THĐ", "AS T5.11THĐ", "ĐH T5.11THĐ", "T7.11.THĐ Sử dụng", "ĐH.P701-3-5", "AS.P701-3-5",
                    "ĐH.P702-4-6", "AS.P702-4-6", "ĐH. P707", "AS. P707", "ĐH. P708", "AS. P708", "Nguồn của Sphon", ""],
        "CSCK": [""] * 28,
        "CSĐK": [""] * 28,
        "Hệ số": [""] * 28,
        "Tổng KWh": [""] * 28,
        "%": [""] * 28,
        "Thanh toán (KWh)": [""] * 28
    })

    df_vna8_vp_frame.loc[4, "Hệ số"] = 200
    df_vna8_vp_frame.loc[26, "Hệ số"] = 20

    df_vna8_vp_frame.loc[1:3, "CSĐK"] = df_vna8_vp.loc["AS+ĐL 8THĐ":"33LTT", "28/02"].values
    df_vna8_vp_frame.loc[1:3, "CSCK"] = df_vna8_vp.loc["AS+ĐL 8THĐ":"33LTT", "31/03"].values
    df_vna8_vp_frame.loc[1:3, "Hệ số"] = df_vna8_vp.loc["AS+ĐL 8THĐ":"33LTT", "Hệ số"].values
    df_vna8_vp_frame.loc[1:3, "Tổng KWh"] = (df_vna8_vp_frame.loc[1:3, "CSCK"] - df_vna8_vp_frame.loc[1:3, "CSĐK"]) * df_vna8_vp_frame.loc[1:3, "Hệ số"]
    df_vna8_vp_frame.loc[1, "Thanh toán (KWh)"] = df_vna8_vp_frame.loc[1:3, "Tổng KWh"].sum()

    df_vna8_vp_frame.loc[4, "CSĐK"] = df.loc["TBA-11THĐ 630KVA-10/0.4KV", "28/02"]
    df_vna8_vp_frame.loc[4, "CSCK"] = df.loc["TBA-11THĐ 630KVA-10/0.4KV", "31/03"]
    df_vna8_vp_frame.loc[4, "Tổng KWh"] = (df_vna8_vp_frame.loc[4, "CSCK"] - df_vna8_vp_frame.loc[4, "CSĐK"]) * df_vna8_vp_frame.loc[4, "Hệ số"]

    df_vna8_vp_frame.loc[6:9, "CSĐK"] = df_vna8_vp.loc["AS18THĐ":"ĐH20THĐ", "28/02"].values
    df_vna8_vp_frame.loc[6:9, "CSCK"] = df_vna8_vp.loc["AS18THĐ":"ĐH20THĐ", "31/03"].values
    df_vna8_vp_frame.loc[6:9, "Hệ số"] = df_vna8_vp.loc["AS18THĐ":"ĐH20THĐ", "Hệ số"].values
    df_vna8_vp_frame.loc[6:9, "Tổng KWh"] = (df_vna8_vp_frame.loc[6:9, "CSCK"] - df_vna8_vp_frame.loc[6:9, "CSĐK"]) * df_vna8_vp_frame.loc[6:9, "Hệ số"]
    df_vna8_vp_frame.loc[6, "Thanh toán (KWh)"] = df_vna8_vp_frame.loc[6:9, "Tổng KWh"].sum()

    df_vna8_vp_frame.loc[11:16, "CSĐK"] = df_vna8_vp.loc["AS-T6-11THĐ":"ĐH T5.11THĐ", "28/02"].values
    df_vna8_vp_frame.loc[11:16, "CSCK"] = df_vna8_vp.loc["AS-T6-11THĐ":"ĐH T5.11THĐ", "31/03"].values
    df_vna8_vp_frame.loc[11:16, "Hệ số"] = df_vna8_vp.loc["AS-T6-11THĐ":"ĐH T5.11THĐ", "Hệ số"].values
    df_vna8_vp_frame.loc[11:16, "Tổng KWh"] = (df_vna8_vp_frame.loc[11:16, "CSCK"] - df_vna8_vp_frame.loc[11:16, "CSĐK"]) * df_vna8_vp_frame.loc[11:16, "Hệ số"]

    df_vna8_vp_frame.loc[18:25, "CSĐK"] = df_vna8_vp.loc["ĐH-T7 P701-703-705":"AS-T7 P708", "28/02"].values
    df_vna8_vp_frame.loc[18:25, "CSCK"] = df_vna8_vp.loc["ĐH-T7 P701-703-705":"AS-T7 P708", "31/03"].values
    df_vna8_vp_frame.loc[18:25, "Hệ số"] = df_vna8_vp.loc["ĐH-T7 P701-703-705":"AS-T7 P708", "Hệ số"].values
    df_vna8_vp_frame.loc[18:25, "Tổng KWh"] = (df_vna8_vp_frame.loc[18:25, "CSCK"] - df_vna8_vp_frame.loc[18:25, "CSĐK"]) * df_vna8_vp_frame.loc[18:25, "Hệ số"]
    df_vna8_vp_frame.loc[18, "Thanh toán (KWh)"] = df_vna8_vp_frame.loc[18:25, "Tổng KWh"].sum()

    df_vna8_vp_frame.loc[11, "Thanh toán (KWh)"] = df_vna8_vp_frame.loc[11:16, "Tổng KWh"].sum() + df_vna8_vp_frame.loc[18:25, "Tổng KWh"].sum()


    df_vna8_vp_frame_total_KWh = df_vna8_vp_frame.loc[1:2, "Tổng KWh"].sum() + df_vna8_vp_frame.loc[4, "Tổng KWh"].sum()
    df_vna8_vp_frame_total_thanhtoan = df_vna8_vp_frame.loc[1:3, "Tổng KWh"].sum() + df_vna8_vp_frame.loc[6:9, "Tổng KWh"].sum() + df_vna8_vp_frame.loc[11:16, "Tổng KWh"].sum() + df_vna8_vp_frame.loc[18:25, "Tổng KWh"].sum()

    df_vna8_vp_frame.loc[27, "Tổng KWh"] = df_vna8_vp_frame_total_KWh
    df_vna8_vp_frame.loc[27, "Thanh toán (KWh)"] = df_vna8_vp_frame_total_thanhtoan

    # Bảng Trung tâm hợp tác Quốc tế Thông tấn phần PG Bank
    df_vna8_pgb = df.loc[["ĐH  -T1 (11THĐ)", "AS - T1 (11THĐ)", "ĐH  -T2 (11THĐ)", "AS - T2 (11THĐ)"]]
    df_vna8_pgb_frame = pd.DataFrame({
        "Stt": [1, 2, 3, 4, ""],
        "Số công tơ": [12067352, 13012746, 13010295, 13010298, "Tổng"],
        "Loại công tơ": ["1 pha", "1 pha", "1 pha", "1 pha", ""],
        "Địa chỉ": ["ĐH -T1 (11THĐ)", "AS - T1 (11THĐ)", "ĐH -T2 (11THĐ)", "AS - T2 (11THĐ)", ""],
        "CSCK": [""]*5,
        "CSĐK": [""]*5,
        "Hệ số": [""]*5,
        "Tổng KWh": [""]*5,
        "%": [""]*5,
        "Thanh toán (KWh)": [""]*5,
    })

    df_vna8_pgb_frame.loc[:3, "CSĐK"] = df_vna8_pgb.loc["ĐH  -T1 (11THĐ)":"AS - T2 (11THĐ)", "28/02"].values
    df_vna8_pgb_frame.loc[:3, "CSCK"] = df_vna8_pgb.loc["ĐH  -T1 (11THĐ)":"AS - T2 (11THĐ)", "31/03"].values
    df_vna8_pgb_frame.loc[:3, "Hệ số"] = df_vna8_pgb.loc["ĐH  -T1 (11THĐ)":"AS - T2 (11THĐ)", "Hệ số"].values
    df_vna8_pgb_frame.loc[:3, "Tổng KWh"] = (df_vna8_pgb_frame.loc[:3, "CSCK"] - df_vna8_pgb_frame.loc[:3, "CSĐK"]) * df_vna8_pgb_frame.loc[:3, "Hệ số"]
    df_vna8_pgb_frame.loc[:3, "Thanh toán (KWh)"] = df_vna8_pgb_frame.loc[:3, "Tổng KWh"]
    df_vna8_pgb_frame.loc[4, "Thanh toán (KWh)"] = sum(df_vna8_pgb_frame.loc[:3, "Thanh toán (KWh)"])

    vna8_8tdh = df_vna8_vp_frame.loc[1, "Thanh toán (KWh)"]
    vna8_18_20thd = df_vna8_vp_frame.loc[6, "Thanh toán (KWh)"]
    vn8_t6_t7_11thd = df_vna8_vp_frame.loc[11, "Thanh toán (KWh)"] + df_vna8_vp_frame.loc[18, "Thanh toán (KWh)"]
    vna8_pgb = df_vna8_pgb_frame.loc[4, "Thanh toán (KWh)"]

    vna_lst = [vna8_8tdh, vna8_18_20thd, vn8_t6_t7_11thd, vna8_pgb]

    df_vna8_frame = pd.DataFrame({
        "Stt": [1, 2, 3, 4, ""],  # Dòng "Tổng" không có số thứ tự
        "Địa chỉ": [
            "Khu vực số 8 THĐ",
            "Khu vực số 18-20 THĐ",
            "Khu vực 11 THĐ",
            "Ngân hàng PG Bank",
            "Tổng"
        ],
        "Tổng KWh": ["", "", "", "", ""],
        "Thanh toán (KWh)": ["", "", "", "", ""],
        "Ghi chú": ["", "", "", "", ""]
    })

    df_vna8_frame.loc[:3, "Tổng KWh"] = vna_lst
    df_vna8_frame.loc[4, "Tổng KWh"] = sum(vna_lst)
    df_vna8_frame.loc[:3, "Thanh toán (KWh)"] = [7590, 2412, 6411, 2352]
    df_vna8_frame.loc[4, "Thanh toán (KWh)"] = sum(df_vna8_frame.loc[:3, "Thanh toán (KWh)"])
    df_vna8_frame


    sheet_dv_dict = {
        "MB1": [df_mb1_frame],
        "MB2": [df_mb2_frame],
        "Giovani1": [df_giovani1_frame],
        "Giovani2": [df_giovani2_frame],
        "GME": [df_gme_frame],
        "Bảo Việt": [df_baoviet_frame],
        "FBS": [df_fsb_frame, df_fbs_old_frame]
    }


    titles_dv_dict = {
        "MB1": ["Ngân hàng Thương mại cổ phần Quân đội, chi nhánh Hoàn Kiếm"],
        "MB2": ["Ngân hàng Thương mại cổ phần Quân đội, chi nhánh Hoàn Kiếm"],
        "Giovani1": ["GIOVANI"],
        "Giovani2": ["GIOVANI"],
        "GME": ["Văn phòng Công ty GME tầng 4 - 79 LTK"],
        "Bảo Việt": ["Bảo Việt"],
        "FBS": ["FBS", "FBS (hệ cũ)"]
    }

    sheet_vna8_dict = {
        "DV8": [df_dv8_vp_frame, df_dv8_pgb_frame],
        "VNA1": [df_vna8_vp_frame, df_vna8_pgb_frame],
        "VNA2": [df_vna8_frame]
    }

    titles_vna8_dict = {
        "DV8": ["VNA8 khu vực Văn phòng", "VNA8 khu vực PG Bank"],
        "VNA1": ["VNA8 khu vực Văn phòng", "VNA8 khu vực PG Bank"],
        "VNA2": ["Trung tâm Hợp tác quốc tế Thông tấn"]
    }

    return sheet_dv_dict, titles_dv_dict, sheet_vna8_dict, titles_vna8_dict
