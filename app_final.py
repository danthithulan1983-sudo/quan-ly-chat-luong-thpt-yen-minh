import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import google.generativeai as genai
import io
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# Cấu hình trang
st.set_page_config(page_title="Quản trị KHTN - THPT Yên Minh", page_icon="🏆", layout="wide")

st.title("🏆 Quản trị Chất lượng 2025 - THPT Yên Minh")
st.markdown("**Hệ sinh thái phân tích điểm số khép kín: Nhập liệu -> Trí tuệ nhân tạo -> Xuất báo cáo Cloud.**")

# ==========================================
# 1. CÁC HÀM XỬ LÝ LÕI (DATA & CLOUD)
# ==========================================

@st.cache_data(ttl=10)
def load_and_transform_data(url):
    """Tải dữ liệu bảng ngang từ Google Sheets và xoay dọc để phân tích"""
    try:
        file_id = url.split("/d/")[1].split("/")[0]
        export_url = f"https://docs.google.com/spreadsheets/d/{file_id}/gviz/tq?tqx=out:csv"
        df_ngang = pd.read_csv(export_url)
        df_ngang.columns = [str(c).strip() for c in df_ngang.columns]
        
        fixed_cols = ['Ma_HS', 'Lop', 'Lan_Thi']
        mon_hoc_cols = [c for c in df_ngang.columns if c not in fixed_cols]
        
        # Unpivot dữ liệu
        df_doc = pd.melt(df_ngang, id_vars=fixed_cols, value_vars=mon_hoc_cols, var_name='Mon_Hoc', value_name='Diem_Thi')
        df_doc = df_doc.dropna(subset=['Diem_Thi'])
        df_doc['Diem_Thi'] = df_doc['Diem_Thi'].astype(str).str.replace(',', '.')
        df_doc['Diem_Thi'] = pd.to_numeric(df_doc['Diem_Thi'], errors='coerce')
        df_doc = df_doc.dropna(subset=['Diem_Thi'])
        
        return df_doc, None
    except Exception as e:
        return None, f"Lỗi kết nối hoặc sai cấu trúc: {e}"

def ghi_ket_qua_len_sheet(df_ket_qua, link_sheet, ten_sheet_dich="Bao_Cao_AI"):
    """Ghi báo cáo ngược lên Google Sheets dùng Tài khoản dịch vụ (Robot)"""
    try:
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        thu_muc_hien_tai = os.path.dirname(os.path.abspath(__file__))
        
        # Radar tìm file credentials
        cac_ten_co_the = ["credentials.json", "credentials", "credentials.json.json", "credentials.txt"]
        duong_dan_chuan = None
        for ten in cac_ten_co_the:
            thu_nghiem = os.path.join(thu_muc_hien_tai, ten)
            if os.path.exists(thu_nghiem):
                duong_dan_chuan = thu_nghiem
                break
                
        if duong_dan_chuan is None:
            return False, "❌ Không tìm thấy file 'credentials.json' bảo mật."

        creds = ServiceAccountCredentials.from_json_keyfile_name(duong_dan_chuan, scope)
        client = gspread.authorize(creds)
        
        sheet_file = client.open_by_url(link_sheet)
        try:
            worksheet = sheet_file.worksheet(ten_sheet_dich)
        except gspread.WorksheetNotFound:
            worksheet = sheet_file.add_worksheet(title=ten_sheet_dich, rows="100", cols="20")
            
        worksheet.clear()
        du_lieu_ghi = [df_ket_qua.columns.values.tolist()] + df_ket_qua.values.tolist()
        worksheet.update(du_lieu_ghi)
        
        return True, f"✅ Đã xuất báo cáo thành công sang Sheet: '{ten_sheet_dich}'!"
    except Exception as e:
        return False, f"❌ Lỗi ghi dữ liệu: {e}"

# ==========================================
# 2. GIAO DIỆN PHÂN QUYỀN (SIDEBAR)
# ==========================================

with st.sidebar:
    st.header("⚙️ Quản trị Hệ thống")
    
    # Giao diện đăng nhập
    admin_password = st.text_input("🔑 Mật khẩu Quản trị:", type="password", placeholder="Dành cho BGH/Tổ trưởng")
    
    # Kiểm tra quyền từ Streamlit Secrets
    is_admin = False
    try:
        if admin_password == st.secrets["ADMIN_PASSWORD"]:
            is_admin = True
            st.success("✅ Đã xác thực quyền Quản trị!")
        elif admin_password != "":
            st.error("❌ Sai mật khẩu!")
        else:
            st.info("👁️ Chế độ Khách (Chỉ xem báo cáo).")
    except Exception as e:
        st.error("⚠️ Lỗi cấu hình file bí mật (.streamlit/secrets.toml)")

    st.divider()
    gsheet_url = st.text_input("🔗 Dán link Google Sheet:", placeholder="File có cột: Ma_HS, Lop, Lan_Thi...")
    
# ==========================================
# 3. LUỒNG XỬ LÝ CHÍNH VÀ BIỂU ĐỒ
# ==========================================

if gsheet_url:
    df_doc, err = load_and_transform_data(gsheet_url)
    if err:
        st.error(err)
    else:
        ds_lan_thi = df_doc['Lan_Thi'].unique()
        ds_mon = df_doc['Mon_Hoc'].unique()
        
        st.markdown("### 🔍 Bộ lọc & Cài đặt Chỉ tiêu")
        col1, col2, col3 = st.columns(3)
        with col1:
            chon_lan = st.selectbox("Chọn Đợt thi:", sorted(ds_lan_thi))
        with col2:
            chon_mon = st.selectbox("Chọn Môn học để phân tích:", sorted(ds_mon))
        with col3:
            chi_tieu_mon = st.number_input(f"🎯 Chỉ tiêu Điểm TB môn {chon_mon}:", min_value=0.0, max_value=10.0, value=6.5, step=0.1)

        df_mon = df_doc[df_doc['Mon_Hoc'] == chon_mon]
        df_hien_tai = df_mon[df_mon['Lan_Thi'] == chon_lan]

        if df_hien_tai.empty:
            st.warning(f"Chưa có dữ liệu môn {chon_mon} trong đợt {chon_lan}.")
        else:
            # --- THỐNG KÊ TỔNG QUAN ---
            st.markdown(f"#### 📊 Thống kê Toàn khối môn {chon_mon}")
            diem_tb_khoi = df_hien_tai['Diem_Thi'].mean()
            chenh_lech = diem_tb_khoi - chi_tieu_mon
            so_hs_thi_mon = len(df_hien_tai)
            hs_dat = len(df_hien_tai[df_hien_tai['Diem_Thi'] >= chi_tieu_mon])
            ti_le_dat = (hs_dat / so_hs_thi_mon) * 100 if so_hs_thi_mon > 0 else 0

            k1, k2, k3 = st.columns(3)
            k1.metric("Trung bình Toàn khối", f"{diem_tb_khoi:.2f}", f"{chenh_lech:.2f} so với Chỉ tiêu")
            k2.metric("Số HS dự thi môn này", f"{so_hs_thi_mon} HS")
            k3.metric("Tỉ lệ học sinh Đạt chỉ tiêu", f"{ti_le_dat:.1f}%")

            # --- BIỂU ĐỒ TRỰC QUAN ---
            tab_chart1, tab_chart2 = st.tabs(["📊 Các lớp vs Chỉ tiêu", "📈 Quỹ đạo tiến bộ"])
            with tab_chart1:
                df_lop = df_hien_tai.groupby('Lop')['Diem_Thi'].mean().reset_index()
                fig_bar = go.Figure()
                fig_bar.add_trace(go.Bar(x=df_lop['Lop'], y=df_lop['Diem_Thi'], name='Điểm TB thực tế', marker_color='#3b82f6'))
                fig_bar.add_hline(y=chi_tieu_mon, line_dash="dash", line_color="red", annotation_text=f"Chỉ tiêu: {chi_tieu_mon}", annotation_position="top right")
                fig_bar.update_layout(title=f"Chất lượng các lớp môn {chon_mon} ({chon_lan})", yaxis_title="Điểm Trung Bình")
                st.plotly_chart(fig_bar, use_container_width=True)

            with tab_chart2:
                df_trend = df_mon.groupby(['Lan_Thi', 'Lop'])['Diem_Thi'].mean().reset_index()
                fig_line = px.line(df_trend, x="Lan_Thi", y="Diem_Thi", color="Lop", markers=True, title=f"Quỹ đạo điểm trung bình môn {chon_mon}")
                fig_line.add_hline(y=chi_tieu_mon, line_dash="dot", line_color="red", annotation_text="Chỉ tiêu")
                st.plotly_chart(fig_line, use_container_width=True)

            # ==========================================
            # 4. KHU VỰC QUẢN TRỊ (AI & EXPORT)
            # ==========================================
            
            st.divider()
            # TÍNH TOÁN BẢNG BÁO CÁO TỔNG HỢP
            bao_cao_list = []
            for lop in df_hien_tai['Lop'].unique():
                lop_data = df_hien_tai[df_hien_tai['Lop'] == lop]
                si_so = len(lop_data)
                diem_tb = lop_data['Diem_Thi'].mean()
                hs_dat_lop = len(lop_data[lop_data['Diem_Thi'] >= chi_tieu_mon])
                ti_le_lop = (hs_dat_lop / si_so) * 100 if si_so > 0 else 0
                
                bao_cao_list.append({
                    'Lớp': lop, 'Sĩ số thi': si_so, 'Chỉ tiêu Giao': chi_tieu_mon,
                    'Điểm TB Thực tế': round(diem_tb, 2), 'Chênh lệch': round(diem_tb - chi_tieu_mon, 2),
                    'Số HS Đạt': hs_dat_lop, 'Tỉ lệ Đạt (%)': round(ti_le_lop, 1)
                })
            df_bao_cao = pd.DataFrame(bao_cao_list)
            
            col_bcao, col_ai = st.columns([1.5, 1])
            
            with col_bcao:
                st.markdown(f"#### 📥 Bảng Báo cáo Tổng hợp")
                st.dataframe(df_bao_cao, use_container_width=True)
                
                # Nút tải Excel (Ai cũng tải được)
                buffer = io.BytesIO()
                with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                    df_bao_cao.to_excel(writer, sheet_name=f'Bao_Cao', index=False)
                st.download_button(label="💾 Tải file Excel về máy tính", data=buffer.getvalue(), file_name=f"Bao_Cao_{chon_mon}_{chon_lan}.xlsx", mime="application/vnd.ms-excel", use_container_width=True)
                
                # Nút Đẩy lên Cloud (Chỉ Admin)
                if st.button("🚀 XUẤT THẲNG LÊN GOOGLE SHEETS", type="primary", use_container_width=True):
                    if is_admin:
                        with st.spinner("Đang ra lệnh cho Robot ghi dữ liệu lên Cloud..."):
                            ten_sheet_moi = f"Báo Cáo {chon_mon} - {chon_lan}"
                            thanh_cong, thong_bao = ghi_ket_qua_len_sheet(df_bao_cao, gsheet_url, ten_sheet_moi)
                            if thanh_cong:
                                st.success(thong_bao)
                                st.balloons()
                            else:
                                st.error(thong_bao)
                    else:
                        st.warning("🔒 Chức năng đồng bộ Cloud chỉ dành cho Ban Giám Hiệu. Vui lòng đăng nhập!")

            with col_ai:
                st.markdown(f"#### 🤖 AI Tham mưu Chuyên môn")
                if st.button(f"Đánh giá & Đề xuất giải pháp bằng AI", use_container_width=True):
                    if is_admin:
                        with st.spinner("AI đang phân tích sự chênh lệch..."):
                            try:
                                genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
                                model = genai.GenerativeModel('gemini-3.1-flash-lite-preview')
                                
                                so_lieu_ai = df_lop.to_string(index=False)
                                prompt = f"""
                                Dưới tư cách là Ban Giám Hiệu, hãy phân tích bảng điểm trung bình môn {chon_mon} của khối 12:
                                {so_lieu_ai}
                                Chỉ tiêu nhà trường giao là {chi_tieu_mon} điểm.
                                Yêu cầu:
                                1. Chỉ ra lớp nào vượt/chưa đạt chỉ tiêu.
                                2. Đề xuất 2 giải pháp quản lý để nâng cao tỉ lệ đạt.
                                """
                                st.info(model.generate_content(prompt).text)
                            except Exception as e:
                                st.error(f"Lỗi API hoặc sai key: {e}")
                    else:
                        st.warning("🔒 Chức năng AI chỉ dành cho Quản trị viên. Vui lòng đăng nhập ở Menu bên trái!")