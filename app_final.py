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

# Cấu hình tab trình duyệt
st.set_page_config(page_title="Quản trị KHTN 2026 - THPT Yên Minh", page_icon="📝", layout="wide")

# ==========================================
# 0. GIAO DIỆN HEADER
# ==========================================
st.markdown("""
<div style="text-align: center; margin-top: 10px; margin-bottom: 0px;">
    <h1 style="color: #1A365D; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-weight: 900; font-size: 2.6rem; text-transform: uppercase; letter-spacing: 2px;">
        QUẢN TRỊ CHẤT LƯỢNG 2026
    </h1>
</div>
""", unsafe_allow_html=True)

col_trai, col_giua, col_phai = st.columns([1, 2, 1]) 
with col_giua:
    try:
        st.image("logo.png", use_container_width=True)
    except Exception as e:
        st.warning("⚠️ Đang chờ tải ảnh logo.png lên...")
        
st.markdown("<hr style='border: 0; height: 1px; background-image: linear-gradient(to right, rgba(0,0,0,0), rgba(0,0,0,0.1), rgba(0,0,0,0)); margin-bottom: 30px;'>", unsafe_allow_html=True)

# ==========================================
# 1. CÁC HÀM XỬ LÝ LÕI
# ==========================================
@st.cache_data(ttl=10)
def load_and_transform_data(url):
    try:
        # 1. Tải dữ liệu thô từ Google Sheets
        file_id = url.split("/d/")[1].split("/")[0]
        export_url = f"https://docs.google.com/spreadsheets/d/{file_id}/gviz/tq?tqx=out:csv"
        df_raw = pd.read_csv(export_url, header=None)
        
        # 2. XÁC ĐỊNH CHÍNH XÁC 2 DÒNG TIÊU ĐỀ
        header_idx = 0
        for i in range(min(5, len(df_raw))):
            row_str = " ".join([str(x).lower() for x in df_raw.iloc[i].values])
            if "họ và tên" in row_str or "họ_và_tên" in row_str or "lớp" in row_str or "lop" in row_str:
                header_idx = i
                break
                
        row0 = df_raw.iloc[header_idx - 1].copy() if header_idx > 0 else df_raw.iloc[0].copy()
        row1 = df_raw.iloc[header_idx].copy()
        
        # 3. LẤP ĐẦY TRỘN Ô
        for i in range(len(row0)):
            val = str(row0.iloc[i]).strip()
            if val == "" or val.lower() in ["nan", "none", "unnamed"]:
                row0.iloc[i] = None
        row0 = row0.ffill() 
        
        # 4. GỘP TIÊU ĐỀ & ĐÁNH DẤU CỘT RÁC
        new_cols = []
        for c0, c1 in zip(row0, row1):
            c0_str = str(c0).strip() if pd.notna(c0) else ""
            c1_str = str(c1).strip() if pd.notna(c1) else ""
            
            if c0_str.lower() in ["nan", "none"] or "unnamed" in c0_str.lower(): c0_str = ""
            if c1_str.lower() in ["nan", "none"] or "unnamed" in c1_str.lower(): c1_str = ""
            
            if c1_str == "":
                new_cols.append("CỘT_RÁC")
            elif c0_str == "":
                new_cols.append(c1_str)
            else:
                new_cols.append(f"{c1_str}|{c0_str}")
                
        # 5. CẮT BỎ TIÊU ĐỀ CŨ VÀ DỌN SẠCH CỘT RÁC
        df_ngang = df_raw.iloc[header_idx + 1:].reset_index(drop=True)
        df_ngang.columns = new_cols
        df_ngang = df_ngang.loc[:, [c for c in df_ngang.columns if c != "CỘT_RÁC"]]
        
        # 6. NHẬN DIỆN CỘT THÔNG TIN VÀ CỘT ĐIỂM
        cac_cot_co_dinh = [c for c in df_ngang.columns if '|' not in c]
        cac_cot_diem = [c for c in df_ngang.columns if '|' in c]
        
        # 7. ÉP DỌC DỮ LIỆU
        df_doc = pd.melt(df_ngang, id_vars=cac_cot_co_dinh, value_vars=cac_cot_diem, var_name='Mon_Lan', value_name='Diem_Thi')
        
        split_cols = df_doc['Mon_Lan'].str.split('|', n=1, expand=True)
        df_doc['Mon_Hoc'] = split_cols[0]
        df_doc['Lan_Thi'] = split_cols[1]
            
        # 8. ĐỒNG BỘ TÊN CỘT BẰNG TỪ ĐIỂN
        rename_dict = {}
        for col in df_doc.columns:
            cl = str(col).lower().replace("_", " ").strip()
            if 'tên' in cl or 'ten' in cl:
                rename_dict[col] = 'Ten_Hoc_Sinh'
            if 'lớp' in cl or 'lop' in cl:
                rename_dict[col] = 'Lop'
        
        df_doc = df_doc.rename(columns=rename_dict)
        if 'Lop' not in df_doc.columns: df_doc['Lop'] = "Khối Chung"
            
        # 9. DỌN DẸP DỮ LIỆU RÁC
        df_doc = df_doc.dropna(subset=['Diem_Thi', 'Mon_Hoc', 'Lan_Thi'])
        df_doc['Diem_Thi'] = df_doc['Diem_Thi'].astype(str).str.replace(',', '.')
        df_doc['Diem_Thi'] = pd.to_numeric(df_doc['Diem_Thi'], errors='coerce')
        df_doc = df_doc.dropna(subset=['Diem_Thi'])
        
        return df_doc, None
    except Exception as e:
        return None, f"🛑 Lỗi đọc dữ liệu: {e}"

# --- HÀM XUẤT GOOGLE SHEETS BỊ THIẾU ĐÃ ĐƯỢC BỔ SUNG LẠI TẠI ĐÂY ---
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
# 2. GIAO DIỆN PHÂN QUYỀN
# ==========================================
with st.sidebar:
    st.header("⚙️ Quản trị Hệ thống")
    admin_password = st.text_input("🔑 Mật khẩu Quản trị:", type="password")
    
    is_admin = False
    try:
        if admin_password == st.secrets["ADMIN_PASSWORD"]:
            is_admin = True
            st.success("✅ Đã xác thực quyền!")
        elif admin_password != "": st.error("❌ Sai mật khẩu!")
    except: st.error("⚠️ Lỗi cấu hình file bí mật")

    st.divider()
    gsheet_url = st.text_input("🔗 Dán link Google Sheet:")
    
# ==========================================
# 3. LUỒNG XỬ LÝ CHÍNH VÀ BIỂU ĐỒ
# ==========================================
if gsheet_url:
    df_doc, err = load_and_transform_data(gsheet_url)
    if err:
        st.error(err)
    elif df_doc.empty:
        st.warning("Tệp dữ liệu đang trống hoặc không có điểm số hợp lệ.")
    else:
        ds_lan_thi = df_doc['Lan_Thi'].unique()
        ds_mon = df_doc['Mon_Hoc'].unique()
        
        st.markdown("### 🔍 Bộ lọc & Cài đặt")
        col1, col2, col3 = st.columns(3)
        with col1: chon_lan = st.selectbox("Chọn Đợt thi:", sorted(ds_lan_thi))
        with col2: chon_mon = st.selectbox("Chọn Môn phân tích:", sorted(ds_mon))
        with col3: chi_tieu_mon = st.number_input(f"🎯 Chỉ tiêu Điểm TB môn {chon_mon}:", value=6.5, step=0.1)

        df_tat_ca_mon_dot_nay = df_doc[df_doc['Lan_Thi'] == chon_lan]
        df_hien_tai = df_tat_ca_mon_dot_nay[df_tat_ca_mon_dot_nay['Mon_Hoc'] == chon_mon].copy()

        if df_hien_tai.empty:
            st.warning(f"Chưa có dữ liệu cho môn {chon_mon} đợt {chon_lan}.")
        else:
            tb_cac_mon = df_tat_ca_mon_dot_nay.groupby('Mon_Hoc')['Diem_Thi'].mean().sort_values()
            
            if tb_cac_mon.empty:
                st.warning("Không có đủ điểm số hợp lệ để xếp hạng môn học.")
            else:
                mon_yeu_nhat = tb_cac_mon.index[0]
                diem_mon_yeu = tb_cac_mon.iloc[0]
                mon_dan_dau = tb_cac_mon.index[-1]
                diem_mon_dan_dau = tb_cac_mon.iloc[-1]
                
                st.markdown(f"#### 📊 Tổng quan Toàn khối")
                k1, k2, k3, k4 = st.columns(4)
                k1.metric("TB Toàn khối môn chọn", f"{df_hien_tai['Diem_Thi'].mean():.2f}")
                k2.metric("Số HS dự thi", f"{len(df_hien_tai)}")
                k3.metric("Môn yếu nhất hiện tại", f"{mon_yeu_nhat}", f"{diem_mon_yeu:.2f} điểm", delta_color="inverse")
                k4.metric("Môn dẫn đầu hiện tại", f"{mon_dan_dau}", f"{diem_mon_dan_dau:.2f} điểm")

                st.markdown("#### 📈 Biểu đồ trực quan Phổ điểm Toàn khối")
                try:
                    bins = [0, 3.4999, 4.9999, 6.4999, 7.9999, 10]
                    labels = ['<3.5', '3.5-5.0', '5.0-6.5', '6.5-8.0', '8.0-10']
                    
                    df_ve = pd.DataFrame()
                    df_ve['Mức điểm'] = pd.cut(df_hien_tai['Diem_Thi'], bins=bins, labels=labels, include_lowest=True)
                    du_lieu_ve = df_ve['Mức điểm'].value_counts().reindex(labels).reset_index()
                    du_lieu_ve.columns = ['Mức điểm', 'Số lượng HS']
                    
                    fig = px.bar(du_lieu_ve, x='Mức điểm', y='Số lượng HS', text='Số lượng HS', color='Mức điểm', color_discrete_sequence=px.colors.qualitative.Pastel)
                    fig.update_traces(textposition='outside', textfont_size=14)
                    fig.update_layout(showlegend=False, xaxis_title="", yaxis_title="Số học sinh", margin=dict(t=20, b=0, l=0, r=0))
                    
                    st.plotly_chart(fig, use_container_width=True)
                except Exception as e:
                    st.error(f"🛑 Lỗi vẽ biểu đồ: {e}")

                # --- TÍNH TOÁN PHỔ ĐIỂM ---
                bins = [-1, 3.499, 4.999, 6.999, 7.999, 10.1]
                labels = ['< 3.5', '3.5 - < 5.0', '5.0 - < 7.0', '7.0 - < 8.0', '8.0 - 10']
                df_hien_tai['Pho_Diem'] = pd.cut(df_hien_tai['Diem_Thi'], bins=bins, labels=labels, right=False)
                
                bang_pho_diem = pd.crosstab(df_hien_tai['Lop'], df_hien_tai['Pho_Diem']).reindex(columns=labels, fill_value=0)
                dong_toan_khoi_pd = pd.DataFrame(bang_pho_diem.sum()).T
                dong_toan_khoi_pd.index = ['⭐ TOÀN KHỐI']
                bang_pho_diem = pd.concat([bang_pho_diem, dong_toan_khoi_pd])
                
                # --- TÍNH TOÁN BẢNG XẾP HẠNG ---
                bao_cao_list = []
                for lop in sorted(df_hien_tai['Lop'].unique()):
                    lop_data = df_hien_tai[df_hien_tai['Lop'] == lop]
                    bao_cao_list.append({
                        'Lớp': lop, 
                        'Sĩ số': len(lop_data), 
                        'Chỉ tiêu Giao': chi_tieu_mon, 
                        'Điểm TB': round(lop_data['Diem_Thi'].mean(), 2), 
                        'Chênh lệch CT': round(lop_data['Diem_Thi'].mean() - chi_tieu_mon, 2)
                    })
                
                df_bao_cao = pd.DataFrame(bao_cao_list)
                df_bao_cao = df_bao_cao.sort_values(by='Điểm TB', ascending=False).reset_index(drop=True)
                df_bao_cao.insert(0, 'Xếp hạng', range(1, len(df_bao_cao) + 1))
                df_tong_hop = pd.merge(df_bao_cao, bang_pho_diem.reset_index(), left_on='Lớp', right_on='index', how='left').drop(columns=['index'])
                
                tb_khoi = df_hien_tai['Diem_Thi'].mean()
                d_toan_khoi = {
                    'Xếp hạng': '-', 'Lớp': '⭐ TOÀN KHỐI', 'Sĩ số': len(df_hien_tai),
                    'Chỉ tiêu Giao': chi_tieu_mon, 
                    'Điểm TB': round(tb_khoi, 2), 'Chênh lệch CT': round(tb_khoi - chi_tieu_mon, 2)
                }
                for col in labels: d_toan_khoi[col] = dong_toan_khoi_pd[col].values[0]
                df_tong_hop.loc[len(df_tong_hop)] = d_toan_khoi

                # --- GIAO DIỆN BÁO CÁO & AI ---
                st.divider()
                col_bcao, col_ai = st.columns([1.6, 1])
                
                with col_bcao:
                    st.markdown(f"#### 📥 Bảng Xếp hạng & Phổ điểm môn {chon_mon}")
                    st.dataframe(df_tong_hop, use_container_width=True, hide_index=True)
                    
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        df_tong_hop.to_excel(writer, sheet_name=f'Bao_Cao_Pho_Diem', index=False)
                        pd.DataFrame({'Môn Học': tb_cac_mon.index, 'Điểm TB': tb_cac_mon.values.round(2)}).to_excel(writer, sheet_name='TB_Cac_Mon', index=False)
                    st.download_button("💾 Tải file Excel Báo cáo", data=buffer.getvalue(), file_name=f"Bao_Cao_{chon_mon}_{chon_lan}.xlsx", mime="application/vnd.ms-excel", use_container_width=True)
                    
                    if st.button("🚀 XUẤT LÊN GOOGLE SHEETS", type="primary", use_container_width=True) and is_admin:
                        with st.spinner("Đang đồng bộ..."):
                            thanh_cong, msg = ghi_ket_qua_len_sheet(df_tong_hop, gsheet_url, f"Báo Cáo {chon_mon} - {chon_lan}")
                            if thanh_cong: st.success(msg)
                            else: st.error(msg)

                with col_ai:
                    st.markdown(f"#### 🤖 AI Tham mưu Lãnh đạo")
                    if "ai_ket_qua" not in st.session_state: st.session_state.ai_ket_qua = ""

                    if st.button(f"Phân tích Phổ điểm & Đề xuất giải pháp", use_container_width=True):
                        if is_admin:
                            with st.spinner("Đang tổng hợp phổ điểm, định vị điểm liệt và soạn thảo..."):
                                try:
                                    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
                                    cac_model = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
                                    model = genai.GenerativeModel(next((m for m in cac_model if 'flash' in m), cac_model[0]))
                                    
                                    prompt = f"""
                                    Dưới tư cách là Ban Giám Hiệu, hãy phân tích kỳ thi đợt {chon_lan}:
                                    1. Bảng Xếp hạng và Phổ điểm môn {chon_mon}:
                                    {df_tong_hop.to_string(index=False)}
                                    
                                    2. Cảnh báo toàn khối: Môn {mon_yeu_nhat} đang có điểm trung bình thấp nhất