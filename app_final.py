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
st.set_page_config(page_title="Quản trị KHTN 2026 - THPT Yên Minh", page_icon="🎓", layout="wide")

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
        
        active_lan = "Lần 1"
        new_cols = []
        
        for i in range(len(row_mon)):
            c_lan = str(row_lan.iloc[i]).strip() if pd.notna(row_lan.iloc[i]) else ""
            c_mon = str(row_mon.iloc[i]).strip() if pd.notna(row_mon.iloc[i]) else ""
            
            match = re.search(r'(lần\s*\d+|đợt\s*\d+)', c_lan, re.IGNORECASE)
            if match: active_lan = match.group(1).title()
                
            if c_mon.lower() in ["nan", "none", "unnamed", ""]: c_mon = ""
            c_mon_lower = c_mon.lower()
            
            is_fixed = False
            if c_mon_lower in ['tt', 'stt', 'sbd', 'họ tên', 'ngày sinh', 'lớp', 'trường', 'ghi chú', 'họ và tên', 'họ_và_tên', 'ngày_tháng_năm_sinh', 'lop', 'ten_hoc_sinh', 'phòng thi', 'phòng', 'mã hs']:
                is_fixed = True
            elif any(k in c_mon_lower for k in ['lớp 10', 'tb 10', 'đtb 10', 'tb10']): is_fixed = True
            elif any(k in c_mon_lower for k in ['lớp 11', 'tb 11', 'đtb 11', 'tb11']): is_fixed = True
            elif any(k in c_mon_lower for k in ['lớp 12', 'tb 12', 'đtb 12', 'tb12']): is_fixed = True
            elif any(k in c_mon_lower for k in ['ưu tiên', 'uu tien', 'điểm ut']): is_fixed = True
            elif c_mon_lower == 'ut': is_fixed = True
            elif any(k in c_mon_lower for k in ['khuyến khích', 'khuyen khich', 'điểm kk']): is_fixed = True
            elif c_mon_lower == 'kk': is_fixed = True
            
            if is_fixed: new_cols.append(c_mon)
            elif c_mon != "": new_cols.append(f"{c_mon}|{active_lan}")
            else: new_cols.append("CỘT_RÁC")
                
        start_data_idx = idx_mon + 1
        df_ngang = df_raw.iloc[start_data_idx:].reset_index(drop=True)
        df_ngang.columns = new_cols
        
        mask_not_rac = [c != "CỘT_RÁC" for c in df_ngang.columns]
        df_ngang = df_ngang.loc[:, mask_not_rac]
        df_ngang = df_ngang.loc[:, ~df_ngang.columns.duplicated()]
        
        cac_cot_thong_tin = [c for c in df_ngang.columns if '|' not in c]
        cac_cot_diem = [c for c in df_ngang.columns if '|' in c]
        
        df_doc = pd.melt(df_ngang, id_vars=cac_cot_thong_tin, value_vars=cac_cot_diem, var_name='_VAR_AI_', value_name='_VAL_AI_')
        
        split_cols = df_doc['_VAR_AI_'].str.split('|', n=1, expand=True)
        df_doc['Mon_Hoc'] = split_cols[0]
        df_doc['Lan_Thi'] = split_cols[1]
        df_doc['Diem_Thi'] = df_doc['_VAL_AI_']
            
        rename_dict = {}
        has_ten = False
        has_lop = False
        for col in df_doc.columns:
            cl = str(col).lower().replace("_", " ").strip()
            if not has_ten and ('tên' in cl or 'ten' in cl) and 'ưu tiên' not in cl:
                rename_dict[col] = 'Ten_Hoc_Sinh'
                has_ten = True
            elif not has_lop and ('lớp' in cl or 'lop' in cl) and not any(x in cl for x in ['10', '11', '12']):
                rename_dict[col] = 'Lop'
                has_lop = True
            elif '10' in cl and any(x in cl for x in ['tb', 'lớp', 'điểm', 'đtb']): rename_dict[col] = 'TB_10'
            elif '11' in cl and any(x in cl for x in ['tb', 'lớp', 'điểm', 'đtb']): rename_dict[col] = 'TB_11'
            elif '12' in cl and any(x in cl for x in ['tb', 'lớp', 'điểm', 'đtb']): rename_dict[col] = 'TB_12'
            elif 'ưu tiên' in cl or 'uu tien' in cl or cl == 'ut' or 'điểm ut' in cl: rename_dict[col] = 'Diem_UT'
            elif 'khuyến khích' in cl or 'khuyen khich' in cl or cl == 'kk' or 'điểm kk' in cl: rename_dict[col] = 'Diem_KK'
        
        df_doc = df_doc.rename(columns=rename_dict)
        if 'Lop' not in df_doc.columns: df_doc['Lop'] = "Khối Chung"
        if 'Ten_Hoc_Sinh' not in df_doc.columns: df_doc['Ten_Hoc_Sinh'] = "Chưa rõ"
            
        cols_to_keep = ['Ten_Hoc_Sinh', 'Lop', 'Mon_Hoc', 'Lan_Thi', 'Diem_Thi']
        for ext in ['TB_10', 'TB_11', 'TB_12', 'Diem_UT', 'Diem_KK']:
            if ext in df_doc.columns: cols_to_keep.append(ext)
            
        df_clean = df_doc[cols_to_keep].copy()
        
        for col in ['TB_10', 'TB_11', 'TB_12', 'Diem_UT', 'Diem_KK']:
            if col in df_clean.columns:
                df_clean[col] = df_clean[col].astype(str).str.replace(',', '.')
                df_clean[col] = pd.to_numeric(df_clean[col], errors='coerce')

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
        df_safe = df_ket_qua.fillna("")
        du_lieu_ghi = [df_safe.columns.values.tolist()] + df_safe.values.tolist()
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
# 3. LUỒNG PHÂN TÍCH CHÍNH & 5 TAB BÁO CÁO
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
        
        tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 1. TỔNG QUAN", "📈 2. TIẾN TRÌNH 1 MÔN", "🔎 3. PHÂN TÍCH LỚP", "🏆 4. BẢNG TỔNG HỢP", "🎓 5. XÉT TN & ĐẠI HỌC"])
        
        with tab1:
            st.markdown("#### 🌟 Biến động Điểm Trung bình Chung qua các Đợt thi")
            tb_khoi_cac_lan = df_doc.groupby('Lan_Thi')['Diem_Thi'].mean().reset_index()
            tb_khoi_cac_lan['Điểm TB Chung'] = tb_khoi_cac_lan['Diem_Thi'].round(2)
            tb_khoi_cac_lan['Chỉ tiêu Giao'] = chi_tieu_chung
            tb_khoi_cac_lan['Chênh lệch'] = (tb_khoi_cac_lan['Điểm TB Chung'] - chi_tieu_chung).round(2)
            
            fig_chung = go.Figure()
            fig_chung.add_trace(go.Scatter(x=tb_khoi_cac_lan['Lan_Thi'], y=tb_khoi_cac_lan['Điểm TB Chung'], mode='lines+markers+text', name='Thực tế', text=tb_khoi_cac_lan['Điểm TB Chung'], textposition='top center', line=dict(color='blue', width=3), marker=dict(size=10)))
            fig_chung.add_trace(go.Scatter(x=tb_khoi_cac_lan['Lan_Thi'], y=tb_khoi_cac_lan['Chỉ tiêu Giao'], mode='lines', name='Chỉ tiêu', line=dict(color='red', width=2, dash='dash')))
            st.plotly_chart(fig_chung, use_container_width=True)

        with tab2:
            chon_mon_tab2 = st.selectbox("🔍 Chọn Môn học để xem tiến trình:", ds_mon, key='mon_tab2')
            df_mon_tien_trinh = df_doc[df_doc['Mon_Hoc'] == chon_mon_tab2]
            tb_mon_cac_lan = df_mon_tien_trinh.groupby('Lan_Thi')['Diem_Thi'].mean().reset_index()
            tb_mon_cac_lan['Điểm TB Môn'] = tb_mon_cac_lan['Diem_Thi'].round(2)
            
            c_bieudo, c_bang = st.columns([2, 1])
            with c_bieudo:
                fig_mon = px.bar(tb_mon_cac_lan, x='Lan_Thi', y='Điểm TB Môn', text='Điểm TB Môn', color='Lan_Thi', color_discrete_sequence=px.colors.qualitative.Set2)
                fig_mon.add_hline(y=chi_tieu_mon, line_dash="dash", line_color="red", annotation_text="Chỉ tiêu môn")
                fig_mon.update_traces(textposition='outside')
                st.plotly_chart(fig_mon, use_container_width=True)

        with tab3:
            cc1, cc2 = st.columns(2)
            with cc1: chon_mon = st.selectbox("Chọn Môn học:", ds_mon, key='mon_tab3')
            with cc2: chon_lan = st.selectbox("Chọn Đợt thi:", ds_lan_thi, key='lan_tab3')
            
            df_hien_tai = df_doc[(df_doc['Mon_Hoc'] == chon_mon) & (df_doc['Lan_Thi'] == chon_lan)].copy()
            if not df_hien_tai.empty:
                bins = [-1, 3.499, 4.999, 6.999, 7.999, 10.1]
                labels = ['< 3.5', '3.5 - < 5.0', '5.0 - < 7.0', '7.0 - < 8.0', '8.0 - 10']
                df_hien_tai['Pho_Diem'] = pd.cut(df_hien_tai['Diem_Thi'], bins=bins, labels=labels, right=False)
                
                # Sửa lỗi: Lấy danh sách Toàn bộ các lớp của trường để 100% không bị mất lớp nào
                danh_sach_lop_all = sorted(df_doc['Lop'].unique())
                bang_pho_diem = pd.crosstab(df_hien_tai['Lop'], df_hien_tai['Pho_Diem']).reindex(index=danh_sach_lop_all, columns=labels, fill_value=0)
                dong_toan_khoi_pd = pd.DataFrame(bang_pho_diem.sum()).T
                dong_toan_khoi_pd.index = ['⭐ TOÀN KHỐI']
                bang_pho_diem = pd.concat([bang_pho_diem, dong_toan_khoi_pd])
                
                bao_cao_list = []
                for lop in danh_sach_lop_all:
                    lop_data = df_hien_tai[df_hien_tai['Lop'] == lop]
                    si_so = len(lop_data)
                    dtb = lop_data['Diem_Thi'].mean()
                    bao_cao_list.append({
                        'Lớp': lop, 'Sĩ số': si_so, 
                        'Điểm TB': round(dtb, 2) if pd.notna(dtb) else None, 
                        'Chênh lệch CT': round(dtb - chi_tieu_mon, 2) if pd.notna(dtb) else None
                    })
                
                df_bao_cao = pd.DataFrame(bao_cao_list).sort_values(by='Điểm TB', ascending=False, na_position='last').reset_index(drop=True)
                df_bao_cao.insert(0, 'Xếp hạng', range(1, len(df_bao_cao) + 1))
                df_tong_hop = pd.merge(df_bao_cao, bang_pho_diem.reset_index(), left_on='Lớp', right_on='index', how='left').drop(columns=['index'])
                
                tb_khoi = df_hien_tai['Diem_Thi'].mean()
                d_toan_khoi = {'Xếp hạng': '-', 'Lớp': '⭐ TOÀN KHỐI', 'Sĩ số': len(df_hien_tai), 'Điểm TB': round(tb_khoi, 2), 'Chênh lệch CT': round(tb_khoi - chi_tieu_mon, 2)}
                for col in labels: d_toan_khoi[col] = dong_toan_khoi_pd[col].values[0]
                df_tong_hop.loc[len(df_tong_hop)] = d_toan_khoi

                st.dataframe(df_tong_hop, use_container_width=True, hide_index=True)

        with tab4:
            c_lan1, c_lan2 = st.columns(2)
            with c_lan1: lan_truoc = st.selectbox("So sánh từ:", ds_lan_thi, index=0, key='lan_truoc_t4')
            with c_lan2: lan_sau = st.selectbox("Đến (Lần hiện tại):", ds_lan_thi, index=len(ds_lan_thi)-1 if len(ds_lan_thi)>1 else 0, key='lan_sau_t4')
            
            df_2lan = df_doc[df_doc['Lan_Thi'].isin([lan_truoc, lan_sau])]
            danh_sach_lop = sorted(df_doc['Lop'].unique())
            danh_sach_lop.append("⭐ TOÀN KHỐI")
            
            du_lieu_bang = []
            for lop in danh_sach_lop:
                df_lop = df_2lan if lop == "⭐ TOÀN KHỐI" else df_2lan[df_2lan['Lop'] == lop]
                tb_truoc = df_lop[df_lop['Lan_Thi'] == lan_truoc]['Diem_Thi'].mean()
                tb_sau = df_lop[df_lop['Lan_Thi'] == lan_sau]['Diem_Thi'].mean()
                row = {'Lớp': lop, 'TB Chung': round(tb_sau, 2) if pd.notna(tb_sau) else None, '+/- Chung': round(tb_sau - tb_truoc, 2) if pd.notna(tb_sau) and pd.notna(tb_truoc) else None}
                
                for mon in ds_mon:
                    df_mon = df_lop[df_lop['Mon_Hoc'] == mon]
                    m_truoc = df_mon[df_mon['Lan_Thi'] == lan_truoc]['Diem_Thi'].mean()
                    m_sau = df_mon[df_mon['Lan_Thi'] == lan_sau]['Diem_Thi'].mean()
                    row[f'{mon}'] = round(m_sau, 2) if pd.notna(m_sau) else None
                    row[f'+/- {mon}'] = round(m_sau - m_truoc, 2) if pd.notna(m_sau) and pd.notna(m_truoc) else None
                du_lieu_bang.append(row)
                
            df_tong_hop_all = pd.DataFrame(du_lieu_bang)
            df_chi_tiet = df_tong_hop_all[df_tong_hop_all['Lớp'] != "⭐ TOÀN KHỐI"].sort_values(by='TB Chung', ascending=False, na_position='last')
            df_toan_khoi = df_tong_hop_all[df_tong_hop_all['Lớp'] == "⭐ TOÀN KHỐI"]
            df_tong_hop_all = pd.concat([df_chi_tiet, df_toan_khoi]).reset_index(drop=True)
            st.dataframe(df_tong_hop_all, use_container_width=True, hide_index=True)

        # ---------------------------------------------------------------------
        # TAB 5: XÉT TỐT NGHIỆP THPT (CÁ NHÂN HÓA 100% THEO FILE) & ĐẠI HỌC
        # ---------------------------------------------------------------------
        with tab5:
            st.markdown("#### 🎓 HỆ THỐNG XÉT TỐT NGHIỆP VÀ ĐẠI HỌC 2026")
            st.info("""
            💡 Nếu file Excel CÓ nhập ĐTB 10, 11, 12, UT, KK: Máy sẽ tính riêng từng học sinh. Nếu bị trống, máy tự bù bằng các số giả lập bên dưới để không một học sinh nào bị loại khỏi bảng.
            """)
            
            c_lan_tab5, c_t10, c_t11, c_t12, c_ut, c_kk = st.columns([1.5, 1, 1, 1, 1, 1])
            with c_lan_tab5: lan_tab5 = st.selectbox("Chọn Đợt thi giả lập:", ds_lan_thi, key='lan_tab5')
            with c_t10: tb_lop10 = st.number_input("🎯 Giả lập Lớp 10:", value=7.0, step=0.1)
            with c_t11: tb_lop11 = st.number_input("🎯 Giả lập Lớp 11:", value=7.0, step=0.1)
            with c_t12: tb_lop12 = st.number_input("🎯 Giả lập Lớp 12:", value=7.0, step=0.1)
            with c_ut: diem_ut = st.number_input("⭐ Giả lập UT:", value=0.0, step=0.25)
            with c_kk: diem_kk = st.number_input("🌟 Giả lập KK:", value=0.0, step=0.5)
            
            df_dot = df_doc[df_doc['Lan_Thi'] == lan_tab5].copy()
            
            # --- CHÌA KHÓA: Bù dữ liệu trước khi Pivot để Pandas KHÔNG THỂ vứt học sinh ---
            if 'TB_10' in df_dot.columns: df_dot['TB_10_Thuc'] = df_dot['TB_10'].fillna(tb_lop10)
            else: df_dot['TB_10_Thuc'] = tb_lop10
            
            if 'TB_11' in df_dot.columns: df_dot['TB_11_Thuc'] = df_dot['TB_11'].fillna(tb_lop11)
            else: df_dot['TB_11_Thuc'] = tb_lop11
            
            if 'TB_12' in df_dot.columns: df_dot['TB_12_Thuc'] = df_dot['TB_12'].fillna(tb_lop12)
            else: df_dot['TB_12_Thuc'] = tb_lop12
            
            if 'Diem_UT' in df_dot.columns: df_dot['UT_Thuc'] = df_dot['Diem_UT'].fillna(diem_ut)
            else: df_dot['UT_Thuc'] = diem_ut
            
            if 'Diem_KK' in df_dot.columns: df_dot['KK_Thuc'] = df_dot['Diem_KK'].fillna(diem_kk)
            else: df_dot['KK_Thuc'] = diem_kk

            index_cols = ['Ten_Hoc_Sinh', 'Lop', 'TB_10_Thuc', 'TB_11_Thuc', 'TB_12_Thuc', 'UT_Thuc', 'KK_Thuc']
            
            df_wide = df_dot.pivot_table(index=index_cols, columns='Mon_Hoc', values='Diem_Thi').reset_index()
            mon_cols = [c for c in df_wide.columns if c not in index_cols]

            # Tính toán chuẩn TT 24/2024/TT-BGDĐT
            dtb_cac_nam = (df_wide['TB_10_Thuc'] * 1 + df_wide['TB_11_Thuc'] * 2 + df_wide['TB_12_Thuc'] * 3) / 6
            df_wide['Tổng 4 Môn'] = df_wide[mon_cols].sum(axis=1)
            df_wide['Điểm Liệt'] = df_wide[mon_cols].min(axis=1)
            
            df_wide['Điểm Xét TN'] = ((((df_wide['Tổng 4 Môn'] + df_wide['KK_Thuc']) / 4) + dtb_cac_nam) / 2 + df_wide['UT_Thuc']).round(2)
            df_wide['Kết quả TN'] = df_wide.apply(lambda row: "ĐỖ ✅" if row['Điểm Xét TN'] >= 5.0 and row['Điểm Liệt'] > 1.0 else "TRƯỢT ❌", axis=1)
            
            # --- TÍNH CÁC KHỐI ĐẠI HỌC ---
            def get_col(danh_sach_cot, keywords):
                for c in danh_sach_cot:
                    for kw in keywords:
                        if kw in str(c).lower(): return c
                return None
            
            mon_map = {
                'Toán': get_col(mon_cols, ['toán', 'toan']),
                'Văn': get_col(mon_cols, ['văn', 'ngữ']),
                'Anh': get_col(mon_cols, ['anh', 'ngoại ngữ']),
                'Lý': get_col(mon_cols, ['lý', 'lí', 'vật']),
                'Hóa': get_col(mon_cols, ['hóa', 'hoa']),
                'Sinh': get_col(mon_cols, ['sinh']),
                'Sử': get_col(mon_cols, ['sử', 'lịch']),
                'Địa': get_col(mon_cols, ['địa', 'dia']),
                'KTPL': get_col(mon_cols, ['ktpl', 'gdcd', 'kinh tế', 'pháp luật', 'gdk'])
            }
            
            ds_to_hop = {
                'A00': ['Toán', 'Lý', 'Hóa'], 'A01': ['Toán', 'Lý', 'Anh'], 'B00': ['Toán', 'Hóa', 'Sinh'], 
                'C00': ['Văn', 'Sử', 'Địa'], 'C14': ['Toán', 'Văn', 'KTPL'], 'C19': ['Văn', 'Sử', 'KTPL'],
                'C20': ['Văn', 'Địa', 'KTPL'], 'D01': ['Toán', 'Văn', 'Anh'], 'D07': ['Toán', 'Hóa', 'Anh']
            }
            
            to_hop_hien_co = []
            for ten_khoi, ds_mon_thanh_phan in ds_to_hop.items():
                cot_thuc_te = [mon_map[m] for m in ds_mon_thanh_phan if mon_map[m] is not None]
                if len(cot_thuc_te) == 3:
                    ten_cot_moi = f"{ten_khoi} ({'-'.join(ds_mon_thanh_phan)})"
                    df_wide[ten_cot_moi] = df_wide[cot_thuc_te].sum(axis=1, skipna=False).round(2)
                    to_hop_hien_co.append(ten_cot_moi)
            
            khoi_truyen_thong = [k for k in to_hop_hien_co if any(x in k for x in ["A00", "A01", "B00", "C00", "D01"])]
            st.markdown("---")
            chon_to_hop = st.multiselect("📌 CHỌN TỔ HỢP ĐẠI HỌC MUỐN XEM:", options=to_hop_hien_co, default=khoi_truyen_thong)
            
            # Đổi tên lại cho đẹp trên bảng hiển thị
            df_wide = df_wide.rename(columns={'TB_10_Thuc': 'ĐTB 10', 'TB_11_Thuc': 'ĐTB 11', 'TB_12_Thuc': 'ĐTB 12', 'UT_Thuc': 'UT', 'KK_Thuc': 'KK'})
            
            cols_to_show = ['Ten_Hoc_Sinh', 'Lop', 'ĐTB 10', 'ĐTB 11', 'ĐTB 12', 'UT', 'KK'] + mon_cols + ['Tổng 4 Môn', 'Điểm Xét TN', 'Kết quả TN'] + chon_to_hop
            df_wide_show = df_wide[cols_to_show]
            
            st.dataframe(df_wide_show, use_container_width=True, hide_index=True)
            
            c_x1, c_x2 = st.columns(2)
            with c_x1:
                buffer_5 = io.BytesIO()
                with pd.ExcelWriter(buffer_5, engine='xlsxwriter') as writer:
                    df_wide_show.to_excel(writer, sheet_name='Xet_Tuyen', index=False)
                st.download_button("💾 Tải Bảng Xét Tốt Nghiệp", data=buffer_5.getvalue(), file_name=f"Xet_TN_{lan_tab5}.xlsx", mime="application/vnd.ms-excel", use_container_width=True)
            with c_x2:
                if st.button("🚀 XUẤT LÊN GOOGLE SHEETS", type="primary", use_container_width=True, key='btn_g5'):
                    if is_admin:
                        thanh_cong, msg = ghi_ket_qua_len_sheet(df_wide_show, gsheet_url, f"Xét TN - {lan_tab5}")
                        if thanh_cong: st.success(msg)
                        else: st.error(msg)
                    else: st.warning("🔒 Vui lòng đăng nhập quyền Quản trị!")