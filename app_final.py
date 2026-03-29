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
import numpy as np

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
# 1. CÁC HÀM XỬ LÝ LÕI ĐỌC DỮ LIỆU & TẠO FILE WORD
# ==========================================
def tao_file_word(noi_dung_ai, tieu_de_bao_cao):
    """Hàm tạo file Word từ văn bản AI sinh ra"""
    doc = docx.Document()
    # Thêm tiêu đề
    h = doc.add_heading(tieu_de_bao_cao, level=1)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Thêm nội dung (Có xử lý in đậm **text** của Markdown cơ bản)
    for line in noi_dung_ai.split('\n'):
        line = line.strip()
        if line:
            p = doc.add_paragraph()
            # Xử lý gạch đầu dòng
            if line.startswith('* ') or line.startswith('- '):
                p.style = 'List Bullet'
                line = line[2:]
            
            # Xử lý in đậm
            parts = line.split('**')
            for idx, part in enumerate(parts):
                run = p.add_run(part)
                if idx % 2 != 0:  # Những phần nằm giữa ** **
                    run.bold = True
                    
    buffer = io.BytesIO()
    doc.save(buffer)
    return buffer.getvalue()

@st.cache_data(ttl=10)
def load_and_transform_data(url):
    try:
        file_id = url.split("/d/")[1].split("/")[0]
        export_url = f"https://docs.google.com/spreadsheets/d/{file_id}/export?format=csv"
        if "gid=" in url:
            gid = url.split("gid=")[1].split("&")[0]
            export_url += f"&gid={gid}"
            
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
        
        chi_tieu_chung_val = 6.0
        dict_chi_tieu_mon = {}
        
        for i in range(min(20, len(df_raw))):
            row_vals = [str(x).strip().lower() for x in df_raw.iloc[i].values]
            row_str = " ".join(row_vals)
            
            match_chung = re.search(r'chỉ tiêu chung\s*[:\-\=]?\s*(\d+[\.,]\d+|\d+)', row_str)
            if match_chung:
                chi_tieu_chung_val = float(match_chung.group(1).replace(',', '.'))
                
            if any('chỉ tiêu' in v or 'chi tieu' in v for v in row_vals):
                for j, val in enumerate(df_raw.iloc[i].values):
                    mon_name = str(row_mon.iloc[j]).strip()
                    if mon_name and mon_name.lower() not in ['tt', 'stt', 'sbd', 'họ tên', 'ngày sinh', 'lớp', 'trường', 'ghi chú', 'họ và tên', 'họ_và_tên', 'ngày_tháng_năm_sinh', 'lop', 'ten_hoc_sinh', 'phòng thi', 'phòng', 'mã hs'] and mon_name.lower() not in ['nan', 'none']:
                        try:
                            v_str = str(val).replace(',', '.').strip()
                            m = re.search(r'[-+]?\d*\.\d+|\d+', v_str)
                            if m:
                                diem_ct = float(m.group())
                                c_mon_sach = re.sub(r'(?i)(lần|đợt)\s*\d+', '', mon_name).strip()
                                c_mon_sach = re.sub(r'[\(\)]', '', c_mon_sach).strip()
                                dict_chi_tieu_mon[c_mon_sach] = diem_ct
                        except: pass
        
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
            elif '10' in c_mon_lower and any(k in c_mon_lower for k in ['lớp', 'tb', 'đtb', 'cn', 'điểm']): is_fixed = True
            elif '11' in c_mon_lower and any(k in c_mon_lower for k in ['lớp', 'tb', 'đtb', 'cn', 'điểm']): is_fixed = True
            elif '12' in c_mon_lower and any(k in c_mon_lower for k in ['lớp', 'tb', 'đtb', 'cn', 'điểm']): is_fixed = True
            elif any(k in c_mon_lower for k in ['ưu tiên', 'uu tien', 'điểm ut']): is_fixed = True
            elif c_mon_lower == 'ut': is_fixed = True
            elif any(k in c_mon_lower for k in ['khuyến khích', 'khuyen khich', 'điểm kk']): is_fixed = True
            elif c_mon_lower == 'kk': is_fixed = True
            
            if is_fixed: 
                new_cols.append(c_mon)
            elif c_mon != "": 
                c_mon_sach = re.sub(r'(?i)(lần|đợt)\s*\d+', '', c_mon).strip()
                c_mon_sach = re.sub(r'[\(\)]', '', c_mon_sach).strip()
                new_cols.append(f"{c_mon_sach}|{active_lan}")
            else: 
                new_cols.append("CỘT_RÁC")
                
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
        mapped_targets = set()
        has_ten = False
        has_lop = False
        for col in df_doc.columns:
            cl = str(col).lower().replace("_", " ").strip()
            target = None
            if not has_ten and ('tên' in cl or 'ten' in cl) and 'ưu tiên' not in cl:
                target = 'Ten_Hoc_Sinh'
                has_ten = True
            elif not has_lop and ('lớp' in cl or 'lop' in cl) and not any(x in cl for x in ['10', '11', '12']):
                target = 'Lop'
                has_lop = True
            elif '10' in cl and any(x in cl for x in ['tb', 'lớp', 'điểm', 'đtb', 'cn']): target = 'TB_10'
            elif '11' in cl and any(x in cl for x in ['tb', 'lớp', 'điểm', 'đtb', 'cn']): target = 'TB_11'
            elif '12' in cl and any(x in cl for x in ['tb', 'lớp', 'điểm', 'đtb', 'cn']): target = 'TB_12'
            elif 'ưu tiên' in cl or 'uu tien' in cl or cl == 'ut' or 'điểm ut' in cl: target = 'Diem_UT'
            elif 'khuyến khích' in cl or 'khuyen khich' in cl or cl == 'kk' or 'điểm kk' in cl: target = 'Diem_KK'
            
            if target and target not in mapped_targets:
                rename_dict[col] = target
                mapped_targets.add(target)
        
        df_doc = df_doc.rename(columns=rename_dict)
        if 'Lop' not in df_doc.columns: df_doc['Lop'] = "Khối Chung"
        if 'Ten_Hoc_Sinh' not in df_doc.columns: df_doc['Ten_Hoc_Sinh'] = "Chưa rõ"
        
        df_doc['Lop'] = df_doc['Lop'].fillna("Chưa rõ Lớp").astype(str).replace('nan', 'Chưa rõ Lớp').replace('', 'Chưa rõ Lớp')
        df_doc['Ten_Hoc_Sinh'] = df_doc['Ten_Hoc_Sinh'].fillna("Chưa rõ").astype(str).replace('nan', 'Chưa rõ').replace('', 'Chưa rõ')
        
        def is_hoc_sinh_that(row):
            ten = str(row['Ten_Hoc_Sinh']).lower().strip()
            lop = str(row['Lop']).lower().strip()
            
            if (ten in ['', 'nan', 'none', 'chưa rõ']) and (lop in ['', 'nan', 'none', 'chưa rõ lớp']):
                return False
                
            tu_khoa_rac = ['chỉ tiêu', 'chi tieu', 'trung bình', 'trung binh', 'tổng cộng', 'tổng điểm', 'điểm tb', 'toàn trường', 'toàn khối', 'tỉ lệ', 'tỷ lệ', 'chênh lệch']
            if any(k in ten for k in tu_khoa_rac) or any(k in lop for k in tu_khoa_rac):
                return False
                
            if ten in ['tb', 'đtb', 'tổng', 'tb chung'] or lop in ['tb', 'đtb', 'tổng', 'tb chung']:
                return False
                
            return True
            
        df_doc = df_doc[df_doc.apply(is_hoc_sinh_that, axis=1)].reset_index(drop=True)
        danh_sach_toan_bo_lop = sorted(list(df_doc['Lop'].unique()))
        
        cols_to_keep = ['Ten_Hoc_Sinh', 'Lop', 'Mon_Hoc', 'Lan_Thi', 'Diem_Thi']
        for ext in ['TB_10', 'TB_11', 'TB_12', 'Diem_UT', 'Diem_KK']:
            if ext in df_doc.columns: cols_to_keep.append(ext)
            
        df_clean = df_doc[cols_to_keep].copy()
        
        def extract_float(val):
            if pd.isna(val): return None
            s = str(val).replace(',', '.').strip()
            if s.lower() in ['', 'nan', 'none', 'null']: return None
            m = re.search(r'[-+]?\d*\.\d+|\d+', s)
            if m: return float(m.group())
            return None

        for col in ['TB_10', 'TB_11', 'TB_12', 'Diem_UT', 'Diem_KK', 'Diem_Thi']:
            if col in df_clean.columns:
                df_clean[col] = df_clean[col].apply(extract_float)

        df_clean = df_clean.dropna(subset=['Diem_Thi', 'Mon_Hoc', 'Lan_Thi'])
        
        return df_clean, danh_sach_toan_bo_lop, chi_tieu_chung_val, dict_chi_tieu_mon, None
    except Exception as e:
        return None, None, 6.0, {}, f"🛑 Lỗi đọc dữ liệu: {e}"

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
    df_doc, list_all_classes, ct_chung_doc, dict_ct_mon_doc, err = load_and_transform_data(gsheet_url)
    if err:
        st.error(err)
    elif df_doc.empty:
        st.warning("Tệp dữ liệu đang trống hoặc cấu trúc chưa chuẩn.")
    else:
        ds_lan_thi = sorted(df_doc['Lan_Thi'].unique())
        ds_mon = sorted(df_doc['Mon_Hoc'].unique())
        
        st.markdown("### 🎯 THIẾT LẬP CHỈ TIÊU KỲ VỌNG")
        st.info("🤖 **Tính năng Nhận diện Tự động:** Hệ thống đã tự động quét dòng 'Chỉ tiêu' từ file Excel. Bạn hoàn toàn có thể tinh chỉnh lại ở bảng dưới đây:")
        
        c1, c2 = st.columns(2)
        with c1: 
            chi_tieu_chung = st.number_input("📈 Chỉ tiêu Điểm TB Chung (Toàn khối):", value=float(ct_chung_doc), step=0.1)
        with c2: 
            chi_tieu_mon_fallback = st.number_input("🎯 Chỉ tiêu Môn mặc định (nếu Excel bị sót môn):", value=6.5, step=0.1)
            
        if dict_ct_mon_doc:
            st.caption(f"📌 *Chỉ tiêu riêng đã đọc từ Excel:* " + " | ".join([f"**{m}:** {d}" for m, d in dict_ct_mon_doc.items()]))
        
        def get_ct_mon(mon_hoc):
            return dict_ct_mon_doc.get(mon_hoc, chi_tieu_mon_fallback)
        
        st.markdown("<br>", unsafe_allow_html=True)
        
        tab1, tab2, tab3, tab4, tab5 = st.tabs(["📊 1. TỔNG QUAN", "📈 2. TIẾN TRÌNH 1 MÔN", "🔎 3. PHÂN TÍCH LỚP", "🏆 4. BẢNG TỔNG HỢP", "🎓 5. XÉT TN & ĐẠI HỌC"])
        
        # --- TAB 1 ---
        with tab1:
            st.markdown("#### 🌟 Biến động Điểm Trung bình Chung qua các Đợt thi")
            df_tb_hs = df_doc.groupby(['Lan_Thi', 'Ten_Hoc_Sinh', 'Lop'])['Diem_Thi'].mean().reset_index()
            tb_khoi_cac_lan = df_tb_hs.groupby('Lan_Thi')['Diem_Thi'].mean().reset_index()
            tb_khoi_cac_lan['Điểm TB Chung'] = tb_khoi_cac_lan['Diem_Thi'].round(2)
            tb_khoi_cac_lan['Chỉ tiêu Giao'] = chi_tieu_chung
            tb_khoi_cac_lan['Chênh lệch'] = (tb_khoi_cac_lan['Điểm TB Chung'] - chi_tieu_chung).round(2)
            
            fig_chung = go.Figure()
            fig_chung.add_trace(go.Scatter(x=tb_khoi_cac_lan['Lan_Thi'], y=tb_khoi_cac_lan['Điểm TB Chung'], mode='lines+markers+text', name='Thực tế', text=tb_khoi_cac_lan['Điểm TB Chung'], textposition='top center', line=dict(color='blue', width=3), marker=dict(size=10)))
            fig_chung.add_trace(go.Scatter(x=tb_khoi_cac_lan['Lan_Thi'], y=tb_khoi_cac_lan['Chỉ tiêu Giao'], mode='lines', name='Chỉ tiêu', line=dict(color='red', width=2, dash='dash')))
            st.plotly_chart(fig_chung, use_container_width=True)
            
            df_t1 = tb_khoi_cac_lan[['Lan_Thi', 'Điểm TB Chung', 'Chỉ tiêu Giao', 'Chênh lệch']]
            st.dataframe(df_t1, use_container_width=True, hide_index=True)
            
            if st.button("🚀 XUẤT LÊN GOOGLE SHEETS", type="primary", key="btn_g1"):
                if is_admin:
                    thanh_cong, msg = ghi_ket_qua_len_sheet(df_t1, gsheet_url, "Tổng quan Chung")
                    if thanh_cong: st.success(msg)
                    else: st.error(msg)
                else: st.warning("🔒 Vui lòng đăng nhập quyền Quản trị!")

        # --- TAB 2 ---
        with tab2:
            chon_mon_tab2 = st.selectbox("🔍 Chọn Môn học để xem tiến trình:", ds_mon, key='mon_tab2')
            
            ct_mon_hien_tai = get_ct_mon(chon_mon_tab2)
            
            df_mon_tien_trinh = df_doc[df_doc['Mon_Hoc'] == chon_mon_tab2]
            tb_mon_cac_lan = df_mon_tien_trinh.groupby('Lan_Thi')['Diem_Thi'].mean().reset_index()
            tb_mon_cac_lan['Điểm TB Môn'] = tb_mon_cac_lan['Diem_Thi'].round(2)
            
            c_bieudo, c_bang = st.columns([2, 1])
            with c_bieudo:
                fig_mon = px.bar(tb_mon_cac_lan, x='Lan_Thi', y='Điểm TB Môn', text='Điểm TB Môn', color='Lan_Thi', color_discrete_sequence=px.colors.qualitative.Set2)
                fig_mon.add_hline(y=ct_mon_hien_tai, line_dash="dash", line_color="red", annotation_text=f"Chỉ tiêu ({ct_mon_hien_tai})")
                fig_mon.update_traces(textposition='outside')
                st.plotly_chart(fig_mon, use_container_width=True)
            with c_bang:
                st.markdown("<br><br>", unsafe_allow_html=True)
                tb_mon_cac_lan['Chỉ tiêu Môn'] = ct_mon_hien_tai
                tb_mon_cac_lan['Chênh lệch'] = (tb_mon_cac_lan['Điểm TB Môn'] - ct_mon_hien_tai).round(2)
                df_t2 = tb_mon_cac_lan[['Lan_Thi', 'Điểm TB Môn', 'Chênh lệch']]
                st.dataframe(df_t2, hide_index=True)
                
                if st.button("🚀 XUẤT LÊN GOOGLE SHEETS", type="primary", use_container_width=True, key="btn_g2"):
                    if is_admin:
                        thanh_cong, msg = ghi_ket_qua_len_sheet(df_t2, gsheet_url, f"Tiến trình {chon_mon_tab2}")
                        if thanh_cong: st.success(msg)
                        else: st.error(msg)
                    else: st.warning("🔒 Vui lòng đăng nhập quyền Quản trị!")

        # --- TAB 3 ---
        with tab3:
            cc1, cc2 = st.columns(2)
            with cc1: chon_mon = st.selectbox("Chọn Môn học:", ds_mon, key='mon_tab3')
            with cc2: chon_lan = st.selectbox("Chọn Đợt thi:", ds_lan_thi, key='lan_tab3')
            
            ct_mon_tab3 = get_ct_mon(chon_mon)
            
            df_hien_tai = df_doc[(df_doc['Mon_Hoc'] == chon_mon) & (df_doc['Lan_Thi'] == chon_lan)].copy()
            if not df_hien_tai.empty:
                bins = [-1, 3.499, 4.999, 6.999, 7.999, 10.1]
                labels = ['< 3.5', '3.5 - < 5.0', '5.0 - < 7.0', '7.0 - < 8.0', '8.0 - 10']
                df_hien_tai['Pho_Diem'] = pd.cut(df_hien_tai['Diem_Thi'], bins=bins, labels=labels, right=False)
                
                bang_pho_diem = pd.crosstab(df_hien_tai['Lop'], df_hien_tai['Pho_Diem']).reindex(index=list_all_classes, columns=labels, fill_value=0)
                dong_toan_khoi_pd = pd.DataFrame(bang_pho_diem.sum()).T
                dong_toan_khoi_pd.index = ['⭐ TOÀN KHỐI']
                bang_pho_diem = pd.concat([bang_pho_diem, dong_toan_khoi_pd])
                
                bao_cao_list = []
                for lop in list_all_classes:
                    lop_data = df_hien_tai[df_hien_tai['Lop'] == lop]
                    si_so = len(lop_data)
                    dtb = lop_data['Diem_Thi'].mean()
                    bao_cao_list.append({
                        'Lớp': lop, 'Sĩ số': si_so, 
                        'Điểm TB': round(dtb, 2) if pd.notna(dtb) else None, 
                        'Chênh lệch CT': round(dtb - ct_mon_tab3, 2) if pd.notna(dtb) else None
                    })
                
                df_bao_cao = pd.DataFrame(bao_cao_list).sort_values(by='Điểm TB', ascending=False, na_position='last').reset_index(drop=True)
                df_bao_cao.insert(0, 'Xếp hạng', range(1, len(df_bao_cao) + 1))
                df_tong_hop = pd.merge(df_bao_cao, bang_pho_diem.reset_index(), left_on='Lớp', right_on='index', how='left').drop(columns=['index'])
                
                def clean_zero(val):
                    if pd.isna(val) or val == 0 or val == 0.0: return ""
                    return val
                
                for c in ['Điểm TB', 'Chênh lệch CT'] + labels:
                    df_tong_hop[c] = df_tong_hop[c].apply(clean_zero)
                
                tb_khoi = df_hien_tai['Diem_Thi'].mean()
                d_toan_khoi = {'Xếp hạng': '-', 'Lớp': '⭐ TOÀN KHỐI', 'Sĩ số': len(df_hien_tai), 'Điểm TB': round(tb_khoi, 2), 'Chênh lệch CT': round(tb_khoi - ct_mon_tab3, 2)}
                for col in labels: d_toan_khoi[col] = dong_toan_khoi_pd[col].values[0]
                df_tong_hop.loc[len(df_tong_hop)] = d_toan_khoi 

                st.dataframe(df_tong_hop, use_container_width=True, hide_index=True)

                c_btn1, c_btn2, c_btn3 = st.columns(3)
                with c_btn1:
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        df_tong_hop.to_excel(writer, sheet_name=f'Bao_Cao', index=False)
                    st.download_button("💾 Tải file Excel Báo cáo", data=buffer.getvalue(), file_name=f"Bao_Cao_{chon_mon}_{chon_lan}.xlsx", mime="application/vnd.ms-excel", use_container_width=True, key="dl_t3")
                with c_btn2:
                    if st.button("🚀 XUẤT LÊN GOOGLE SHEETS", type="primary", use_container_width=True, key="btn_g3"):
                        if is_admin:
                            thanh_cong, msg = ghi_ket_qua_len_sheet(df_tong_hop, gsheet_url, f"Báo Cáo {chon_mon} - {chon_lan}")
                            if thanh_cong: st.success(msg)
                            else: st.error(msg)
                        else: st.warning("🔒 Vui lòng đăng nhập quyền Quản trị!")
                with c_btn3:
                    if st.button("🤖 AI TƯ VẤN ĐỀ XUẤT", type="primary", use_container_width=True, key="btn_ai_t3"):
                        if is_admin:
                            with st.spinner("Đang phân tích và lập kế hoạch chiến lược..."):
                                try:
                                    genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
                                    cac_model = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
                                    model = genai.GenerativeModel(next((m for m in cac_model if 'flash' in m), cac_model[0]))
                                    prompt = f"""
                                    Đóng vai trò là Chuyên gia Cố vấn Giáo dục. Hãy phân tích kết quả thi môn {chon_mon} đợt {chon_lan} dựa trên bảng số liệu sau:
                                    {df_tong_hop.to_string(index=False)}
                                    Chỉ tiêu kỳ vọng của môn này là {ct_mon_tab3}.
                                    Yêu cầu phân tích:
                                    1. Nhận diện các lớp đang dẫn đầu và các lớp đang ở nhóm rủi ro cao.
                                    2. Đánh giá mức độ đạt chỉ tiêu của toàn khối.
                                    3. ĐỀ XUẤT GIẢI PHÁP CHIẾN LƯỢC: Phương pháp bồi dưỡng, phụ đạo để nâng cao chất lượng.
                                    """
                                    st.session_state.ai_ket_qua_t3 = model.generate_content(prompt).text
                                except Exception as e: st.error(f"Lỗi AI: {e}")
                        else: st.warning("🔒 Cần quyền Quản trị để dùng AI!")
                        
                if "ai_ket_qua_t3" in st.session_state and st.session_state.ai_ket_qua_t3 != "":
                    st.markdown("#### 💡 Giải pháp Nâng cao Chất lượng (AI Đề xuất)")
                    st.info(st.session_state.ai_ket_qua_t3)
                    # NÚT XUẤT WORD
                    word_data_t3 = tao_file_word(st.session_state.ai_ket_qua_t3, f"BÁO CÁO PHÂN TÍCH CHUYÊN MÔN - {chon_mon.upper()} ({chon_lan.upper()})")
                    st.download_button(
                        label="📄 Tải Báo cáo AI (Định dạng Word)",
                        data=word_data_t3,
                        file_name=f"Bao_cao_AI_{chon_mon}_{chon_lan}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key="btn_word_t3"
                    )

        # --- TAB 4 ---
        with tab4:
            st.markdown("#### 🏆 Bảng Xếp Hạng Tổng Hợp & Đánh giá Sự Tiến Bộ (Chi tiết)")
            
            c_lan1, c_lan2 = st.columns(2)
            with c_lan1: lan_truoc = st.selectbox("So sánh từ:", ds_lan_thi, index=0, key='lan_truoc_t4')
            with c_lan2: lan_sau = st.selectbox("Đến (Lần hiện tại):", ds_lan_thi, index=len(ds_lan_thi)-1 if len(ds_lan_thi)>1 else 0, key='lan_sau_t4')
            
            df_2lan = df_doc[df_doc['Lan_Thi'].isin([lan_truoc, lan_sau])]
            
            # TÍNH CHUẨN XÁC: TB Học sinh trước, rồi mới lấy TB Lớp
            df_hs_tab4 = df_2lan.groupby(['Lan_Thi', 'Ten_Hoc_Sinh', 'Lop'])['Diem_Thi'].mean().reset_index()
            
            danh_sach_lop = list(list_all_classes)
            danh_sach_lop.append("⭐ TOÀN KHỐI")
            
            du_lieu_bang = []
            for lop in danh_sach_lop:
                if lop == "⭐ TOÀN KHỐI":
                    tb_truoc = df_hs_tab4[df_hs_tab4['Lan_Thi'] == lan_truoc]['Diem_Thi'].mean()
                    tb_sau = df_hs_tab4[df_hs_tab4['Lan_Thi'] == lan_sau]['Diem_Thi'].mean()
                    df_lop = df_2lan
                else:
                    tb_truoc = df_hs_tab4[(df_hs_tab4['Lan_Thi'] == lan_truoc) & (df_hs_tab4['Lop'] == lop)]['Diem_Thi'].mean()
                    tb_sau = df_hs_tab4[(df_hs_tab4['Lan_Thi'] == lan_sau) & (df_hs_tab4['Lop'] == lop)]['Diem_Thi'].mean()
                    df_lop = df_2lan[df_2lan['Lop'] == lop]
                
                row = {
                    'Lớp': lop, 
                    f'TB Chung ({lan_truoc})': round(tb_truoc, 2) if pd.notna(tb_truoc) else None,
                    f'TB Chung ({lan_sau})': round(tb_sau, 2) if pd.notna(tb_sau) else None,
                    '+/- 2 Lần': round(tb_sau - tb_truoc, 2) if pd.notna(tb_sau) and pd.notna(tb_truoc) else None,
                    '+/- Chỉ tiêu': round(tb_sau - chi_tieu_chung, 2) if pd.notna(tb_sau) else None
                }
                
                for mon in ds_mon:
                    df_mon = df_lop[df_lop['Mon_Hoc'] == mon]
                    m_truoc = df_mon[df_mon['Lan_Thi'] == lan_truoc]['Diem_Thi'].mean()
                    m_sau = df_mon[df_mon['Lan_Thi'] == lan_sau]['Diem_Thi'].mean()
                    
                    row[f'{mon} ({lan_truoc})'] = round(m_truoc, 2) if pd.notna(m_truoc) else None
                    row[f'{mon} ({lan_sau})'] = round(m_sau, 2) if pd.notna(m_sau) else None
                    row[f'+/- {mon}'] = round(m_sau - m_truoc, 2) if pd.notna(m_sau) and pd.notna(m_truoc) else None
                du_lieu_bang.append(row)
                
            df_tong_hop_all = pd.DataFrame(du_lieu_bang)
            
            col_sort = f'TB Chung ({lan_sau})'
            df_chi_tiet = df_tong_hop_all[df_tong_hop_all['Lớp'] != "⭐ TOÀN KHỐI"].sort_values(by=col_sort, ascending=False, na_position='last')
            df_toan_khoi = df_tong_hop_all[df_tong_hop_all['Lớp'] == "⭐ TOÀN KHỐI"]
            df_tong_hop_all = pd.concat([df_chi_tiet, df_toan_khoi]).reset_index(drop=True)
            
            def hide_zero(val):
                if pd.isna(val) or val == 0 or val == 0.0: return ""
                return val
            
            for col in df_tong_hop_all.columns:
                if col != 'Lớp':
                    df_tong_hop_all[col] = df_tong_hop_all[col].apply(hide_zero)

            def color_chenh_lech(val):
                try:
                    if val == "": return ''
                    v = float(val)
                    if v > 0: return 'color: #155724; background-color: #d4edda; font-weight: bold;'
                    elif v < 0: return 'color: #721c24; background-color: #f8d7da; font-weight: bold;'
                except: pass
                return ''

            cot_can_to_mau = [c for c in df_tong_hop_all.columns if '+/-' in c]
            styled_df = df_tong_hop_all.style.map(color_chenh_lech, subset=cot_can_to_mau)
            
            st.dataframe(styled_df, use_container_width=True, hide_index=True)

            c_x1, c_x2, c_x3 = st.columns(3)
            with c_x1:
                buffer_all = io.BytesIO()
                with pd.ExcelWriter(buffer_all, engine='xlsxwriter') as writer:
                    df_tong_hop_all.to_excel(writer, sheet_name='Tong_Hop', index=False)
                st.download_button("💾 Tải Excel Bảng Tổng Hợp", data=buffer_all.getvalue(), file_name=f"Tong_Hop_{lan_truoc}_{lan_sau}.xlsx", mime="application/vnd.ms-excel", use_container_width=True, key="dl_t4")
            with c_x2:
                if st.button("🚀 XUẤT LÊN GOOGLE SHEETS", type="primary", use_container_width=True, key="btn_g4"):
                    if is_admin:
                        thanh_cong, msg = ghi_ket_qua_len_sheet(df_tong_hop_all, gsheet_url, f"Tổng hợp {lan_truoc} - {lan_sau}")
                        if thanh_cong: st.success(msg)
                        else: st.error(msg)
                    else: st.warning("🔒 Vui lòng đăng nhập quyền Quản trị!")
            with c_x3:
                if st.button("🤖 AI PHÂN TÍCH TOÀN TRƯỜNG", type="primary", use_container_width=True, key="btn_ai_t4"):
                    if is_admin:
                        with st.spinner("Đang đánh giá sự tiến bộ của toàn trường..."):
                            try:
                                genai.configure(api_key=st.secrets["GEMINI_API_KEY"])
                                cac_model = [m.name for m in genai.list_models() if 'generateContent' in m.supported_generation_methods]
                                model = genai.GenerativeModel(next((m for m in cac_model if 'flash' in m), cac_model[0]))
                                prompt = f"""
                                Đóng vai trò là Hiệu trưởng phụ trách chuyên môn. Hãy phân tích bảng dữ liệu tổng hợp sự tiến bộ của toàn trường:
                                {df_tong_hop_all.to_string(index=False)}
                                1. Nhận xét tổng quan.
                                2. Vinh danh lớp tiến bộ (+/- tăng cao).
                                3. Cảnh báo lớp/môn học tụt dốc.
                                4. ĐỀ XUẤT GIẢI PHÁP QUẢN TRỊ.
                                """
                                st.session_state.ai_ket_qua_t4 = model.generate_content(prompt).text
                            except Exception as e: st.error(f"Lỗi AI: {e}")
                    else: st.warning("🔒 Cần quyền Quản trị để dùng AI!")
                        
            if "ai_ket_qua_t4" in st.session_state and st.session_state.ai_ket_qua_t4 != "":
                st.markdown("#### 💡 Cố vấn Quản trị Chất lượng (AI Đề xuất)")
                st.info(st.session_state.ai_ket_qua_t4)
                # NÚT XUẤT WORD
                word_data_t4 = tao_file_word(st.session_state.ai_ket_qua_t4, "BÁO CÁO PHÂN TÍCH CHẤT LƯỢNG TOÀN TRƯỜNG")
                st.download_button(
                    label="📄 Tải Báo cáo AI (Định dạng Word)",
                    data=word_data_t4,
                    file_name="Bao_cao_AI_Toan_Truong.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="btn_word_t4"
                )

        # ---------------------------------------------------------------------
        # TAB 5: XÉT TỐT NGHIỆP THPT & ĐẠI HỌC (ĐÃ BỔ SUNG ĐẦY ĐỦ TỔ HỢP)
        # ---------------------------------------------------------------------
        with tab5:
            st.markdown("#### 🎓 HỆ THỐNG XÉT TỐT NGHIỆP VÀ ĐẠI HỌC 2026")
            
            c_lan_tab5, c_t10, c_t11, c_t12, c_ut, c_kk = st.columns([1.5, 1, 1, 1, 1, 1])
            with c_lan_tab5: lan_tab5 = st.selectbox("Chọn Đợt thi giả lập:", ds_lan_thi, key='lan_tab5')
            with c_t10: tb_lop10 = st.number_input("🎯 Giả lập Lớp 10:", value=7.0, step=0.1)
            with c_t11: tb_lop11 = st.number_input("🎯 Giả lập Lớp 11:", value=7.0, step=0.1)
            with c_t12: tb_lop12 = st.number_input("🎯 Giả lập Lớp 12:", value=7.0, step=0.1)
            with c_ut: diem_ut = st.number_input("⭐ Giả lập UT:", value=0.0, step=0.25)
            with c_kk: diem_kk = st.number_input("🌟 Giả lập KK:", value=0.0, step=0.5)
            
            df_dot = df_doc[df_doc['Lan_Thi'] == lan_tab5].copy()
            
            for col_goc, mock_val, col_thuc in [
                ('TB_10', tb_lop10, 'TB_10_Thuc'),
                ('TB_11', tb_lop11, 'TB_11_Thuc'),
                ('TB_12', tb_lop12, 'TB_12_Thuc'),
                ('Diem_UT', diem_ut, 'UT_Thuc'),
                ('Diem_KK', diem_kk, 'KK_Thuc')
            ]:
                if col_goc in df_dot.columns:
                    df_dot[col_thuc] = df_dot[col_goc].fillna(0.0)
                else:
                    df_dot[col_thuc] = mock_val

            index_cols = ['Ten_Hoc_Sinh', 'Lop', 'TB_10_Thuc', 'TB_11_Thuc', 'TB_12_Thuc', 'UT_Thuc', 'KK_Thuc']
            
            df_wide = df_dot.pivot_table(index=index_cols, columns='Mon_Hoc', values='Diem_Thi').reset_index()
            mon_cols = [c for c in df_wide.columns if c not in index_cols]

            dtb_cac_nam = (df_wide['TB_10_Thuc'] * 1 + df_wide['TB_11_Thuc'] * 2 + df_wide['TB_12_Thuc'] * 3) / 6
            df_wide['Tổng 4 Môn'] = df_wide[mon_cols].sum(axis=1)
            df_wide['Điểm Liệt'] = df_wide[mon_cols].min(axis=1)
            
            df_wide['Điểm Xét TN'] = ((((df_wide['Tổng 4 Môn'] + df_wide['KK_Thuc']) / 4) + dtb_cac_nam) / 2 + df_wide['UT_Thuc']).round(2)
            df_wide['Kết quả TN'] = df_wide.apply(lambda row: "ĐỖ ✅" if row['Điểm Xét TN'] >= 5.0 and row['Điểm Liệt'] > 1.0 else "TRƯỢT ❌", axis=1)
            
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
            
            # --- BỔ SUNG ĐẦY ĐỦ CÁC TỔ HỢP HIỆN HÀNH THEO GDPT 2018 ---
            ds_to_hop = {
                'A00': ['Toán', 'Lý', 'Hóa'], 'A01': ['Toán', 'Lý', 'Anh'], 'A02': ['Toán', 'Lý', 'Sinh'],
                'A03': ['Toán', 'Lý', 'Sử'], 'A04': ['Toán', 'Lý', 'Địa'], 'A05': ['Toán', 'Hóa', 'Sử'],
                'A06': ['Toán', 'Hóa', 'Địa'], 'A07': ['Toán', 'Sử', 'Địa'], 'A08': ['Toán', 'Sử', 'KTPL'],
                'A09': ['Toán', 'Địa', 'KTPL'], 'A10': ['Toán', 'Lý', 'KTPL'], 'A11': ['Toán', 'Hóa', 'KTPL'],
                'B00': ['Toán', 'Hóa', 'Sinh'], 'B02': ['Toán', 'Sinh', 'Địa'], 'B03': ['Toán', 'Sinh', 'Sử'], 'B08': ['Toán', 'Sinh', 'Anh'],
                'C00': ['Văn', 'Sử', 'Địa'], 'C01': ['Toán', 'Văn', 'Lý'], 'C02': ['Toán', 'Văn', 'Hóa'],
                'C03': ['Toán', 'Văn', 'Sử'], 'C04': ['Toán', 'Văn', 'Địa'], 'C05': ['Văn', 'Lý', 'Hóa'],
                'C06': ['Văn', 'Lý', 'Sinh'], 'C07': ['Văn', 'Lý', 'Sử'], 'C08': ['Văn', 'Hóa', 'Sinh'],
                'C09': ['Văn', 'Lý', 'Địa'], 'C10': ['Văn', 'Hóa', 'Sử'], 'C11': ['Văn', 'Hóa', 'Địa'],
                'C12': ['Văn', 'Sinh', 'Sử'], 'C13': ['Văn', 'Sinh', 'Địa'], 'C14': ['Toán', 'Văn', 'KTPL'],
                'C16': ['Văn', 'Lý', 'KTPL'], 'C17': ['Văn', 'Hóa', 'KTPL'], 'C18': ['Văn', 'Sinh', 'KTPL'],
                'C19': ['Văn', 'Sử', 'KTPL'], 'C20': ['Văn', 'Địa', 'KTPL'],
                'D01': ['Toán', 'Văn', 'Anh'], 'D07': ['Toán', 'Hóa', 'Anh'], 'D08': ['Toán', 'Sinh', 'Anh'],
                'D09': ['Toán', 'Sử', 'Anh'], 'D10': ['Toán', 'Địa', 'Anh'], 'D11': ['Văn', 'Lý', 'Anh'],
                'D12': ['Văn', 'Hóa', 'Anh'], 'D13': ['Văn', 'Sinh', 'Anh'], 'D14': ['Văn', 'Sử', 'Anh'],
                'D15': ['Văn', 'Địa', 'Anh']
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
            
            df_wide = df_wide.rename(columns={'TB_10_Thuc': 'ĐTB 10', 'TB_11_Thuc': 'ĐTB 11', 'TB_12_Thuc': 'ĐTB 12', 'UT_Thuc': 'UT', 'KK_Thuc': 'KK'})
            
            cac_cot_thong_tin_co = [c for c in ['TB_10', 'TB_11', 'TB_12', 'Diem_UT', 'Diem_KK'] if c in df_dot.columns]
            
            hien_thi_cols = []
            if 'TB_10' in cac_cot_thong_tin_co: hien_thi_cols.append('ĐTB 10')
            if 'TB_11' in cac_cot_thong_tin_co: hien_thi_cols.append('ĐTB 11')
            if 'TB_12' in cac_cot_thong_tin_co: hien_thi_cols.append('ĐTB 12')
            if 'Diem_UT' in cac_cot_thong_tin_co: hien_thi_cols.append('UT')
            if 'Diem_KK' in cac_cot_thong_tin_co: hien_thi_cols.append('KK')

            cols_to_show = ['Ten_Hoc_Sinh', 'Lop'] + hien_thi_cols + mon_cols + ['Tổng 4 Môn', 'Điểm Xét TN', 'Kết quả TN'] + chon_to_hop
            df_wide_show = df_wide[cols_to_show]
            
            def hide_zero_t5(val):
                if pd.isna(val) or val == 0 or val == 0.0: return ""
                return val
            for col in df_wide_show.columns:
                if col not in ['Ten_Hoc_Sinh', 'Lop', 'Kết quả TN']:
                    df_wide_show[col] = df_wide_show[col].apply(hide_zero_t5)
            
            st.dataframe(df_wide_show, use_container_width=True, hide_index=True)
            
            c_x1, c_x2 = st.columns(2)
            with c_x1:
                buffer_5 = io.BytesIO()
                with pd.ExcelWriter(buffer_5, engine='xlsxwriter') as writer:
                    df_wide_show.to_excel(writer, sheet_name='Xet_Tuyen', index=False)
                st.download_button("💾 Tải Bảng Xét Tốt Nghiệp", data=buffer_5.getvalue(), file_name=f"Xet_TN_{lan_tab5}.xlsx", mime="application/vnd.ms-excel", use_container_width=True, key="dl_t5")
            with c_x2:
                if st.button("🚀 XUẤT LÊN GOOGLE SHEETS", type="primary", use_container_width=True, key='btn_g5'):
                    if is_admin:
                        thanh_cong, msg = ghi_ket_qua_len_sheet(df_wide_show, gsheet_url, f"Xét TN - {lan_tab5}")
                        if thanh_cong: st.success(msg)
                        else: st.error(msg)
                    else: st.warning("🔒 Vui lòng đăng nhập quyền Quản trị!")