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
        
        # 2. TRUY TÌM CHÍNH XÁC DÒNG "LẦN THI" VÀ DÒNG "MÔN HỌC"
        idx_lan = -1
        idx_mon = -1
        
        for i in range(min(10, len(df_raw))):
            row_str = " ".join([str(x).lower() for x in df_raw.iloc[i].values])
            # Quét tìm dòng chứa đợt thi
            if "lần" in row_str or "đợt" in row_str:
                idx_lan = i
            # Quét tìm dòng chứa tên môn học chuyên môn
            if "toán" in row_str or "ngữ" in row_str or "vật" in row_str or "hóa" in row_str or "sinh" in row_str:
                idx_mon = i
                
        if idx_lan == -1: idx_lan = 0
        if idx_mon == -1: idx_mon = 1
        
        row_lan = df_raw.iloc[idx_lan].copy()
        row_mon = df_raw.iloc[idx_mon].copy()
        
        # 3. LẤP ĐẦY TRỘN Ô CHO DÒNG LẦN THI
        for i in range(len(row_lan)):
            val = str(row_lan.iloc[i]).strip()
            if val == "" or val.lower() in ["nan", "none", "unnamed"]:
                row_lan.iloc[i] = None
        row_lan = row_lan.ffill() 
        
        # 4. GỘP TIÊU ĐỀ THÔNG MINH
        new_cols = []
        for c_lan, c_mon in zip(row_lan, row_mon):
            lan = str(c_lan).strip() if pd.notna(c_lan) else ""
            mon = str(c_mon).strip() if pd.notna(c_mon) else ""
            
            if lan.lower() in ["nan", "none", "unnamed"]: lan = ""
            if mon.lower() in ["nan", "none", "unnamed"]: mon = ""
            
            if mon == "":
                if "lần" in lan.lower() or "đợt" in lan.lower():
                    new_cols.append("CỘT_RÁC")
                else:
                    new_cols.append(lan) # Bắt các cột cố định (TT, Lớp, Họ tên)
            elif lan == "" or ("lần" not in lan.lower() and "đợt" not in lan.lower()):
                new_cols.append(mon)
            else:
                new_cols.append(f"{mon}|{lan}") # Bắt các cột điểm thi (Toán|Lần 1)
                
        # 5. CHỐT DỮ LIỆU & DỌN CỘT RÁC
        start_data_idx = max(idx_lan, idx_mon) + 1
        df_ngang = df_raw.iloc[start_data_idx:].reset_index(drop=True)
        df_ngang.columns = new_cols
        df_ngang = df_ngang.loc[:, [c for c in df_ngang.columns if c != "CỘT_RÁC"]]
        
        # 6. NHẬN DIỆN VÀ ÉP DỌC DỮ LIỆU
        cac_cot_co_dinh = [c for c in df_ngang.columns if '|' not in c]
        cac_cot_diem = [c for c in df_ngang.columns if '|' in c]
        
        df_doc = pd.melt(df_ngang, id_vars=cac_cot_co_dinh, value_vars=cac_cot_diem, var_name='Mon_Lan', value_name='Diem_Thi')
        
        split_cols = df_doc['Mon_Lan'].str.split('|', n=1, expand=True)
        df_doc['Mon_Hoc'] = split_cols[0]
        df_doc['Lan_Thi'] = split_cols[1]
            
        # 7. CHUẨN HÓA BẮT BUỘC TÊN CỘT LỚP (Để hiện chi tiết Bảng xếp hạng)
        rename_dict = {}
        for col in df_doc.columns:
            cl = str(col).lower().replace("_", " ").strip()
            if 'tên' in cl or 'ten' in cl:
                rename_dict[col] = 'Ten_Hoc_Sinh'
            elif 'lớp' in cl or 'lop' in cl:
                rename_dict[col] = 'Lop'
        
        df_doc = df_doc.rename(columns=rename_dict)
        
        # 8. DỌN DẸP DỮ LIỆU RÁC CỦA TỪNG HỌC SINH
        df_doc = df_doc.dropna(subset=['Diem_Thi', 'Mon_Hoc', 'Lan_Thi'])
        df_doc['Diem_Thi'] = df_doc['Diem_Thi'].astype(str).str.replace(',', '.')
        df_doc['Diem_Thi'] = pd.to_numeric(df_doc['Diem_Thi'], errors='coerce')
        df_doc = df_doc.dropna(subset=['Diem_Thi'])
        
        return df_doc, None
    except Exception as e:
        return None, f"🛑 Lỗi đọc dữ liệu: {e}"
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
                                    
                                    # CHÍNH LÀ CHỖ NÀY: Dấu ngoặc kép đã được khóa chặt!
                                    prompt = f"""
                                    Dưới tư cách là Ban Giám Hiệu, hãy phân tích kỳ thi đợt {chon_lan}:
                                    1. Bảng Xếp hạng và Phổ điểm môn {chon_mon}:
                                    {df_tong_hop.to_string(index=False)}
                                    
                                    2. Cảnh báo toàn khối: Môn {mon_yeu_nhat} đang có điểm trung bình thấp nhất ({diem_mon_yeu:.2f} điểm).
                                    
                                    Yêu cầu viết báo cáo:
                                    - Nhận xét Xếp hạng các lớp môn {chon_mon}. Đánh giá chi tiết sự phân bổ phổ điểm (đặc biệt nhấn mạnh thực trạng học sinh nhóm < 3.5 và nhóm 3.5 - <5.0).
                                    - Chỉ ra nguyên nhân có thể khiến môn {mon_yeu_nhat} tụt dốc.
                                    - Đề xuất 3 giải pháp thực chiến, cấp bách để kéo điểm trung bình, xóa mù điểm liệt, chuẩn bị cho kỳ thi tốt nghiệp THPT sắp tới.
                                    """
                                    st.session_state.ai_ket_qua = model.generate_content(prompt).text
                                except Exception as e: st.error(f"Lỗi AI: {e}")
                        else: st.warning("🔒 Vui lòng đăng nhập quyền Quản trị!")

                    if st.session_state.ai_ket_qua != "":
                        van_ban = st.text_area("Khung Soạn thảo Báo cáo:", value=st.session_state.ai_ket_qua, height=400)
                        st.session_state.ai_ket_qua = van_ban
                        
                        def tao_file_word(noi_dung):
                            doc = docx.Document()
                            for section in doc.sections:
                                section.page_width = Cm(21)
                                section.page_height = Cm(29.7)
                                section.left_margin = Cm(3)
                                section.right_margin = Cm(2)
                                section.top_margin = Cm(2)
                                section.bottom_margin = Cm(2)

                            style = doc.styles['Normal']
                            font = style.font
                            font.name = 'Times New Roman'
                            font.size = Pt(14)
                            
                            p_qh = doc.add_paragraph()
                            p_qh.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            run_qh1 = p_qh.add_run("CỘNG HÒA XÃ HỘI CHỦ NGHĨA VIỆT NAM\n")
                            run_qh1.bold = True
                            run_qh2 = p_qh.add_run("Độc lập - Tự do - Hạnh phúc")
                            run_qh2.bold = True
                            run_qh2.underline = True
                            
                            p_title = doc.add_paragraph()
                            p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
                            run_title = p_title.add_run(f"\nBÁO CÁO THAM MƯU CHUYÊN MÔN\nMÔN: {chon_mon.upper()} - ĐỢT: {chon_lan.upper()}")
                            run_title.bold = True
                            
                            cac_dong = noi_dung.split('\n')
                            for dong in cac_dong:
                                if dong.strip() != "": 
                                    p = doc.add_paragraph(dong)
                                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY 
                                    p.paragraph_format.space_after = Pt(6) 
                                    p.paragraph_format.line_spacing = 1.2 
                                    
                            buffer_word = io.BytesIO()
                            doc.save(buffer_word)
                            buffer_word.seek(0)
                            return buffer_word
                        
                        file_word_san_sang = tao_file_word(st.session_state.ai_ket_qua)
                        st.download_button(
                            label="📄 Tải Báo cáo Word (.docx)",
                            data=file_word_san_sang,
                            file_name=f"Bao_Cao_Tham_Muu_AI_{chon_mon}_{chon_lan}.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            use_container_width=True
                        )