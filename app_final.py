import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import google.generativeai as genai
import io
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import docx
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
import re

# Cấu hình tab trình duyệt
st.set_page_config(page_title="Quản trị KHTN 2026 - THPT Yên Minh", page_icon="📝", layout="wide")

# ==========================================
# 0. GIAO DIỆN HEADER
# ==========================================
st.markdown("""
<div style="text-align: center; margin-top: 10px; margin-bottom: 0px;">
    <h1 style="color: #1A365D; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-weight: 900; font-size: 2.6rem; text-transform: uppercase; letter-spacing: 2px;">
        HỆ THỐNG QUẢN TRỊ CHẤT LƯỢNG ĐA CHIỀU
    </h1>
</div>
""", unsafe_allow_html=True)

col_trai, col_giua, col_phai = st.columns([1, 2, 1]) 
with col_giua:
    try:
        st.image("logo.png", use_container_width=True)
    except Exception as e:
        pass
        
st.markdown("<hr style='border: 0; height: 1px; background-image: linear-gradient(to right, rgba(0,0,0,0), rgba(0,0,0,0.1), rgba(0,0,0,0)); margin-bottom: 30px;'>", unsafe_allow_html=True)

# ==========================================
# 1. CÁC HÀM XỬ LÝ LÕI ĐỌC DỮ LIỆU
# ==========================================
@st.cache_data(ttl=10)
def load_and_transform_data(url):
    try:
        file_id = url.split("/d/")[1].split("/")[0]
        export_url = f"https://docs.google.com/spreadsheets/d/{file_id}/gviz/tq?tqx=out:csv"
        df_raw = pd.read_csv(export_url, header=None)
        
        # 1. TRUY TÌM CHÍNH XÁC DÒNG TIÊU ĐỀ THEO MẪU SỞ
        idx_mon = -1
        for i in range(min(10, len(df_raw))):
            row_str = " ".join([str(x).lower() for x in df_raw.iloc[i].values])
            if "toán" in row_str and ("văn" in row_str or "ngữ" in row_str) and ("lớp" in row_str or "lop" in row_str):
                idx_mon = i
                break
                
        if idx_mon == -1: idx_mon = 1
        idx_lan = idx_mon - 1 if idx_mon > 0 else 0
        
        row_lan = df_raw.iloc[idx_lan].copy()
        row_mon = df_raw.iloc[idx_mon].copy()
        
        # 2. RADAR NHẬN DIỆN LẦN THI ĐỘNG
        active_lan = "Lần 1"
        new_cols = []
        
        cot_co_dinh = ['tt', 'stt', 'sbd', 'họ tên', 'ngày sinh', 'lớp', 'trường', 'ghi chú', 'họ và tên', 'họ_và_tên', 'ngày_tháng_năm_sinh', 'lop', 'ten_hoc_sinh', 'phòng thi', 'phòng', 'mã hs']
        
        for i in range(len(row_mon)):
            c_lan = str(row_lan.iloc[i]).strip() if pd.notna(row_lan.iloc[i]) else ""
            c_mon = str(row_mon.iloc[i]).strip() if pd.notna(row_mon.iloc[i]) else ""
            
            match = re.search(r'(lần\s*\d+|đợt\s*\d+)', c_lan, re.IGNORECASE)
            if match:
                active_lan = match.group(1).title()
                
            if c_mon.lower() in ["nan", "none", "unnamed", ""]: 
                c_mon = ""
                
            c_mon_lower = c_mon.lower()
            
            if c_mon_lower in cot_co_dinh:
                new_cols.append(c_mon)
            elif c_mon != "":
                new_cols.append(f"{c_mon}|{active_lan}")
            else:
                new_cols.append("CỘT_RÁC")
                
        # 3. CẮT BỎ TIÊU ĐỀ CŨ VÀ DỌN SẠCH CỘT RÁC
        start_data_idx = idx_mon + 1
        df_ngang = df_raw.iloc[start_data_idx:].reset_index(drop=True)
        df_ngang.columns = new_cols
        
        mask_not_rac = [c != "CỘT_RÁC" for c in df_ngang.columns]
        df_ngang = df_ngang.loc[:, mask_not_rac]
        df_ngang = df_ngang.loc[:, ~df_ngang.columns.duplicated()]
        
        # 4. ÉP DỌC DỮ LIỆU CHUYÊN NGHIỆP
        cac_cot_thong_tin = [c for c in df_ngang.columns if '|' not in c]
        cac_cot_diem = [c for c in df_ngang.columns if '|' in c]
        
        df_doc = pd.melt(df_ngang, id_vars=cac_cot_thong_tin, value_vars=cac_cot_diem, 
                         var_name='_VAR_AI_', value_name='_VAL_AI_')
        
        split_cols = df_doc['_VAR_AI_'].str.split('|', n=1, expand=True)
        df_doc['Mon_Hoc'] = split_cols[0]
        df_doc['Lan_Thi'] = split_cols[1]
        df_doc['Diem_Thi'] = df_doc['_VAL_AI_']
            
        # 5. CHUẨN HÓA TÊN CỘT ĐỂ KHÔNG BAO GIỜ LỖI
        rename_dict = {}
        has_ten = False
        has_lop = False
        for col in df_doc.columns:
            cl = str(col).lower().replace("_", " ").strip()
            if not has_ten and ('tên' in cl or 'ten' in cl):
                rename_dict[col] = 'Ten_Hoc_Sinh'
                has_ten = True
            elif not has_lop and ('lớp' in cl or 'lop' in cl):
                rename_dict[col] = 'Lop'
                has_lop = True
        
        df_doc = df_doc.rename(columns=rename_dict)
        if 'Lop' not in df_doc.columns: df_doc['Lop'] = "Khối Chung"
        if 'Ten_Hoc_Sinh' not in df_doc.columns: df_doc['Ten_Hoc_Sinh'] = "Chưa rõ"
            
        # 6. RÚT TRÍCH VÀ LÀM SẠCH ĐIỂM SỐ
        df_clean = df_doc[['Ten_Hoc_Sinh', 'Lop', 'Mon_Hoc', 'Lan_Thi', 'Diem_Thi']].copy()
        
        # --- FIX TYPEERROR: Ép cột Lớp thành chuỗi Text và thay thế ô trống ---
        df_clean['Lop'] = df_clean['Lop'].fillna("Chưa rõ Lớp").astype(str)
        df_clean['Lop'] = df_clean['Lop'].replace('nan', 'Chưa rõ Lớp')
        df_clean['Lop'] = df_clean['Lop'].replace('', 'Chưa rõ Lớp')
        
        df_clean = df_clean.dropna(subset=['Diem_Thi', 'Mon_Hoc', 'Lan_Thi'])
        df_clean['Diem_Thi'] = df_clean['Diem_Thi'].astype(str).str.replace(',', '.')
        df_clean['Diem_Thi'] = pd.to_numeric(df_clean['Diem_Thi'], errors='coerce')
        df_clean = df_clean.dropna(subset=['Diem_Thi'])
        
        return df_clean, None
    except Exception as e:
        return None, f"🛑 Lỗi đọc dữ liệu: {e}"

# --- HÀM XUẤT GOOGLE SHEETS ---
def ghi_ket_qua_len_sheet(df_ket_qua, link_sheet, ten_sheet_dich="Bao_Cao_AI"):
    try:
        import json
        scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
        try:
            creds_dict = json.loads(st.secrets["GOOGLE_CREDENTIALS"], strict=False)
            creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
        except Exception as e:
            return False, f"❌ Chưa cấu hình Két sắt (Secrets) hoặc sai định dạng: {e}"

        client = gspread.authorize(creds)
        sheet_file = client.open_by_url(link_sheet)
        
        try: worksheet = sheet_file.worksheet(ten_sheet_dich)
        except: worksheet = sheet_file.add_worksheet(title=ten_sheet_dich, rows="100", cols="20")
            
        worksheet.clear()
        du_lieu_ghi = [df_ket_qua.columns.values.tolist()] + df_ket_qua.values.tolist()
        worksheet.update(du_lieu_ghi)
        return True, f"✅ Đã xuất báo cáo thành công sang Sheet: '{ten_sheet_dich}'!"
    except Exception as e:
        return False, f"❌ Lỗi ghi dữ liệu: {e}"

# ==========================================
# 2. GIAO DIỆN HỆ THỐNG
# ==========================================
with st.sidebar:
    st.header("⚙️ Quản trị Hệ thống")
    admin_password = st.text_input("🔑 Mật khẩu Quản trị:", type="password")
    is_admin = False
    if admin_password == st.secrets.get("ADMIN_PASSWORD", ""):
        is_admin = True
        st.success("✅ Đã xác thực quyền!")
    
    st.divider()
    gsheet_url = st.text_input("🔗 Dán link Google Sheet:")

# ==========================================
# 3. LUỒNG PHÂN TÍCH CHÍNH & TAB BÁO CÁO
# ==========================================
if gsheet_url:
    df_doc, err = load_and_transform_data(gsheet_url)
    if err:
        st.error(err)
    elif df_doc.empty:
        st.warning("Tệp dữ liệu đang trống hoặc cấu trúc chưa chuẩn.")
    else:
        ds_lan_thi = sorted(df_doc['Lan_Thi'].unique())
        ds_mon = sorted(df_doc['Mon_Hoc'].unique())
        
        st.markdown("### 🎯 THIẾT LẬP CHỈ TIÊU KỲ VỌNG")
        c1, c2 = st.columns(2)
        with c1: chi_tieu_chung = st.number_input("📈 Chỉ tiêu Điểm TB Chung (Toàn khối):", value=6.0, step=0.1)
        with c2: chi_tieu_mon = st.number_input("🎯 Chỉ tiêu Điểm TB Bộ môn:", value=6.5, step=0.1)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        # --- 4 TAB PHÂN TÍCH CHUYÊN SÂU ---
        tab1, tab2, tab3, tab4 = st.tabs(["📊 1. TỔNG QUAN CHUNG", "📈 2. TIẾN TRÌNH 1 MÔN", "🔎 3. PHÂN TÍCH SÂU 1 MÔN", "🏆 4. BẢNG TỔNG HỢP TOÀN DIỆN"])
        
        # ---------------------------------------------------------------------
        # TAB 1: TỔNG QUAN TOÀN KHỐI
        # ---------------------------------------------------------------------
        with tab1:
            st.markdown("#### 🌟 Biến động Điểm Trung bình Chung qua các Đợt thi")
            tb_khoi_cac_lan = df_doc.groupby('Lan_Thi')['Diem_Thi'].mean().reset_index()
            tb_khoi_cac_lan['Điểm TB Chung'] = tb_khoi_cac_lan['Diem_Thi'].round(2)
            tb_khoi_cac_lan['Chỉ tiêu Giao'] = chi_tieu_chung
            tb_khoi_cac_lan['Chênh lệch'] = (tb_khoi_cac_lan['Điểm TB Chung'] - chi_tieu_chung).round(2)
            
            fig_chung = go.Figure()
            fig_chung.add_trace(go.Scatter(x=tb_khoi_cac_lan['Lan_Thi'], y=tb_khoi_cac_lan['Điểm TB Chung'],
                                           mode='lines+markers+text', name='Thực tế',
                                           text=tb_khoi_cac_lan['Điểm TB Chung'], textposition='top center',
                                           line=dict(color='blue', width=3), marker=dict(size=10)))
            fig_chung.add_trace(go.Scatter(x=tb_khoi_cac_lan['Lan_Thi'], y=tb_khoi_cac_lan['Chỉ tiêu Giao'],
                                           mode='lines', name='Chỉ tiêu',
                                           line=dict(color='red', width=2, dash='dash')))
            fig_chung.update_layout(title="Tiến trình chất lượng Toàn khối", yaxis_title="Điểm TB Chung")
            st.plotly_chart(fig_chung, use_container_width=True)
            
            st.dataframe(tb_khoi_cac_lan[['Lan_Thi', 'Điểm TB Chung', 'Chỉ tiêu Giao', 'Chênh lệch']], use_container_width=True, hide_index=True)

        # ---------------------------------------------------------------------
        # TAB 2: TIẾN TRÌNH BỘ MÔN
        # ---------------------------------------------------------------------
        with tab2:
            chon_mon_tab2 = st.selectbox("🔍 Chọn Môn học để xem tiến trình:", ds_mon, key='mon_tab2')
            df_mon_tien_trinh = df_doc[df_doc['Mon_Hoc'] == chon_mon_tab2]
            tb_mon_cac_lan = df_mon_tien_trinh.groupby('Lan_Thi')['Diem_Thi'].mean().reset_index()
            tb_mon_cac_lan['Điểm TB Môn'] = tb_mon_cac_lan['Diem_Thi'].round(2)
            
            c_bieudo, c_bang = st.columns([2, 1])
            with c_bieudo:
                fig_mon = px.bar(tb_mon_cac_lan, x='Lan_Thi', y='Điểm TB Môn', text='Điểm TB Môn',
                                 title=f"Biểu đồ điểm TB môn {chon_mon_tab2} qua các lần thi",
                                 color='Lan_Thi', color_discrete_sequence=px.colors.qualitative.Set2)
                fig_mon.add_hline(y=chi_tieu_mon, line_dash="dash", line_color="red", annotation_text="Chỉ tiêu môn")
                fig_mon.update_traces(textposition='outside')
                st.plotly_chart(fig_mon, use_container_width=True)
            
            with c_bang:
                st.markdown("<br><br>", unsafe_allow_html=True)
                tb_mon_cac_lan['Chỉ tiêu Môn'] = chi_tieu_mon
                tb_mon_cac_lan['Chênh lệch'] = (tb_mon_cac_lan['Điểm TB Môn'] - chi_tieu_mon).round(2)
                st.dataframe(tb_mon_cac_lan[['Lan_Thi', 'Điểm TB Môn', 'Chênh lệch']], hide_index=True)

        # ---------------------------------------------------------------------
        # TAB 3: PHÂN TÍCH SÂU 1 MÔN (XẾP HẠNG & PHỔ ĐIỂM)
        # ---------------------------------------------------------------------
        with tab3:
            cc1, cc2 = st.columns(2)
            with cc1: chon_mon = st.selectbox("Chọn Môn học:", ds_mon, key='mon_tab3')
            with cc2: chon_lan = st.selectbox("Chọn Đợt thi:", ds_lan_thi, key='lan_tab3')
            
            df_hien_tai = df_doc[(df_doc['Mon_Hoc'] == chon_mon) & (df_doc['Lan_Thi'] == chon_lan)].copy()
            
            if df_hien_tai.empty:
                st.warning(f"Chưa có dữ liệu cho môn {chon_mon} đợt {chon_lan}.")
            else:
                bins = [-1, 3.499, 4.999, 6.999, 7.999, 10.1]
                labels = ['< 3.5', '3.5 - < 5.0', '5.0 - < 7.0', '7.0 - < 8.0', '8.0 - 10']
                df_hien_tai['Pho_Diem'] = pd.cut(df_hien_tai['Diem_Thi'], bins=bins, labels=labels, right=False)
                
                bang_pho_diem = pd.crosstab(df_hien_tai['Lop'], df_hien_tai['Pho_Diem']).reindex(columns=labels, fill_value=0)
                dong_toan_khoi_pd = pd.DataFrame(bang_pho_diem.sum()).T
                dong_toan_khoi_pd.index = ['⭐ TOÀN KHỐI']
                bang_pho_diem = pd.concat([bang_pho_diem, dong_toan_khoi_pd])
                
                bao_cao_list = []
                # Đã an toàn tuyệt đối nhờ ép kiểu ở hàm load_data
                for lop in sorted(df_hien_tai['Lop'].unique()):
                    lop_data = df_hien_tai[df_hien_tai['Lop'] == lop]
                    bao_cao_list.append({
                        'Lớp': lop, 'Sĩ số': len(lop_data), 
                        'Điểm TB': round(lop_data['Diem_Thi'].mean(), 2), 
                        'Chênh lệch CT': round(lop_data['Diem_Thi'].mean() - chi_tieu_mon, 2)
                    })
                
                df_bao_cao = pd.DataFrame(bao_cao_list).sort_values(by='Điểm TB', ascending=False).reset_index(drop=True)
                df_bao_cao.insert(0, 'Xếp hạng', range(1, len(df_bao_cao) + 1))
                df_tong_hop = pd.merge(df_bao_cao, bang_pho_diem.reset_index(), left_on='Lớp', right_on='index', how='left').drop(columns=['index'])
                
                tb_khoi = df_hien_tai['Diem_Thi'].mean()
                d_toan_khoi = {'Xếp hạng': '-', 'Lớp': '⭐ TOÀN KHỐI', 'Sĩ số': len(df_hien_tai), 'Điểm TB': round(tb_khoi, 2), 'Chênh lệch CT': round(tb_khoi - chi_tieu_mon, 2)}
                for col in labels: d_toan_khoi[col] = dong_toan_khoi_pd[col].values[0]
                df_tong_hop.loc[len(df_tong_hop)] = d_toan_khoi

                st.markdown(f"#### 📥 Phân tích Phổ điểm & Xếp hạng (Môn {chon_mon} - {chon_lan})")
                st.dataframe(df_tong_hop, use_container_width=True, hide_index=True)
                
                c_btn1, c_btn2, c_btn3 = st.columns(3)
                with c_btn1:
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        df_tong_hop.to_excel(writer, sheet_name=f'Bao_Cao', index=False)
                    st.download_button("💾 Tải file Excel Báo cáo", data=buffer.getvalue(), file_name=f"Bao_Cao_{chon_mon}_{chon_lan}.xlsx", mime="application/vnd.ms-excel", use_container_width=True)
                with c_btn2:
                    if st.button("🚀 XUẤT LÊN GOOGLE SHEETS", type="primary", use_container_width=True):
                        if is_admin:
                            thanh_cong, msg = ghi_ket_qua_len_sheet(df_tong_hop, gsheet_url, f"Báo Cáo {chon_mon} - {chon_lan}")
                            if thanh_cong: st.success(msg)
                            else: st.error(msg)
                        else: st.warning("🔒 Vui lòng đăng nhập quyền Quản trị!")
                with c_btn3:
                    if st.button("🤖 AI Tham mưu Báo cáo", type="primary", use_container_width=True):
                        if is_admin:
                            with st.spinner("Đang phân tích phổ điểm..."):
                                try:
                                    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
                                    cac_model = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
                                    model = genai.GenerativeModel(next((m for m in cac_model if 'flash' in m), cac_model[0]))
                                    prompt = f"""
                                    Phân tích kết quả môn {chon_mon} đợt {chon_lan}. Bảng điểm chi tiết:
                                    {df_tong_hop.to_string(index=False)}
                                    Viết báo cáo đánh giá sự chênh lệch giữa các lớp, nhận diện lớp yếu kém, điểm liệt và đề xuất giải pháp.
                                    """
                                    st.session_state.ai_ket_qua = model.generate_content(prompt).text
                                except Exception as e: st.error(f"Lỗi AI: {e}")
                        else: st.warning("🔒 Cần quyền Quản trị để dùng AI!")
                        
                if "ai_ket_qua" in st.session_state and st.session_state.ai_ket_qua != "":
                    st.markdown("#### 📝 Văn bản Tham mưu")
                    st.text_area("Khung Soạn thảo:", value=st.session_state.ai_ket_qua, height=300)

        # ---------------------------------------------------------------------
        # TAB 4: BẢNG TỔNG HỢP TOÀN DIỆN (TẤT CẢ CÁC MÔN & SO SÁNH)
        # ---------------------------------------------------------------------
        with tab4:
            st.markdown("#### 🏆 Bảng Xếp Hạng Tổng Hợp Các Môn & Đánh giá Sự Tiến Bộ")
            st.info("💡 Bảng này giúp BGH nhìn toàn cảnh điểm TB của tất cả các môn trên cùng 1 hàng, đồng thời so sánh chênh lệch (+/-) so với lần thi trước.")
            
            c_lan1, c_lan2 = st.columns(2)
            with c_lan1: lan_truoc = st.selectbox("So sánh từ (Mốc thời gian):", ds_lan_thi, index=0, key='lan_truoc_t4')
            with c_lan2: lan_sau = st.selectbox("Đến (Lần hiện tại):", ds_lan_thi, index=len(ds_lan_thi)-1 if len(ds_lan_thi)>1 else 0, key='lan_sau_t4')
            
            df_2lan = df_doc[df_doc['Lan_Thi'].isin([lan_truoc, lan_sau])]
            
            danh_sach_lop = sorted(df_doc['Lop'].unique())
            danh_sach_lop.append("⭐ TOÀN KHỐI")
            
            du_lieu_bang = []
            for lop in danh_sach_lop:
                df_lop = df_2lan if lop == "⭐ TOÀN KHỐI" else df_2lan[df_2lan['Lop'] == lop]
                
                tb_truoc = df_lop[df_lop['Lan_Thi'] == lan_truoc]['Diem_Thi'].mean()
                tb_sau = df_lop[df_lop['Lan_Thi'] == lan_sau]['Diem_Thi'].mean()
                
                row = {
                    'Lớp': lop,
                    'TB Chung': round(tb_sau, 2) if pd.notna(tb_sau) else None,
                    '+/- Chung': round(tb_sau - tb_truoc, 2) if pd.notna(tb_sau) and pd.notna(tb_truoc) else None
                }
                
                for mon in ds_mon:
                    df_mon = df_lop[df_lop['Mon_Hoc'] == mon]
                    m_truoc = df_mon[df_mon['Lan_Thi'] == lan_truoc]['Diem_Thi'].mean()
                    m_sau = df_mon[df_mon['Lan_Thi'] == lan_sau]['Diem_Thi'].mean()
                    
                    row[f'{mon}'] = round(m_sau, 2) if pd.notna(m_sau) else None
                    row[f'+/- {mon}'] = round(m_sau - m_truoc, 2) if pd.notna(m_sau) and pd.notna(m_truoc) else None
                
                du_lieu_bang.append(row)
                
            df_tong_hop_all = pd.DataFrame(du_lieu_bang)
            df_chi_tiet = df_tong_hop_all[df_tong_hop_all['Lớp'] != "⭐ TOÀN KHỐI"].sort_values(by='TB Chung', ascending=False)
            df_toan_khoi = df_tong_hop_all[df_tong_hop_all['Lớp'] == "⭐ TOÀN KHỐI"]
            df_tong_hop_all = pd.concat([df_chi_tiet, df_toan_khoi]).reset_index(drop=True)
            
            st.dataframe(df_tong_hop_all, use_container_width=True, hide_index=True)
            
            c_x1, c_x2 = st.columns(2)
            with c_x1:
                buffer_all = io.BytesIO()
                with pd.ExcelWriter(buffer_all, engine='xlsxwriter') as writer:
                    df_tong_hop_all.to_excel(writer, sheet_name='Tong_Hop_Toan_Dien', index=False)
                st.download_button("💾 Tải Excel Bảng Tổng Hợp", data=buffer_all.getvalue(), file_name=f"Tong_Hop_{lan_truoc}_den_{lan_sau}.xlsx", mime="application/vnd.ms-excel", use_container_width=True)
            with c_x2:
                if st.button("🚀 XUẤT BẢNG TỔNG HỢP LÊN GOOGLE SHEETS", type="primary", use_container_width=True):
                    if is_admin:
                        thanh_cong, msg = ghi_ket_qua_len_sheet(df_tong_hop_all, gsheet_url, f"Tổng hợp {lan_truoc} - {lan_sau}")
                        if thanh_cong: st.success(msg)
                        else: st.error(msg)
                    else: st.warning("🔒 Vui lòng đăng nhập quyền Quản trị!")