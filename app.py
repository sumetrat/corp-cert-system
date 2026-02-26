import streamlit as st
import pandas as pd
import datetime
import io
import zipfile
import os
import time
import json
from PIL import Image, ImageDraw, ImageFont
import gspread
from google.oauth2.service_account import Credentials

# ==========================================
# 1. ระบบเชื่อมต่อ Google Sheets API
# ==========================================
SHEET_NAME = "Corp_Cert_DB"

@st.cache_resource
def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    try:
        if os.path.exists("service_account.json"):
            creds = Credentials.from_service_account_file("service_account.json", scopes=scopes)
        else:
            try:
                if "GCP_CREDENTIALS" in st.secrets:
                    creds_dict = json.loads(st.secrets["GCP_CREDENTIALS"])
                    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
                else:
                    st.error("❌ ไม่พบทั้งไฟล์ service_account.json และ Secrets บน Cloud")
                    st.stop()
            except FileNotFoundError:
                st.error("❌ ไม่พบไฟล์กุญแจ service_account.json")
                st.stop()
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"❌ Connection Error: {e}")
        st.stop()

def get_records_sheet():
    gc = get_gspread_client()
    return gc.open(SHEET_NAME).worksheet("Records")

def get_users_sheet():
    gc = get_gspread_client()
    return gc.open(SHEET_NAME).worksheet("Users")

def generate_serial():
    sheet = get_records_sheet()
    records = sheet.get_all_records()
    prefix = f"CERT-{datetime.datetime.now().strftime('%Y%m')}"
    count = sum(1 for r in records if str(r.get('serial_number', '')).startswith(prefix))
    return f"{prefix}-{count + 1:04d}"

def save_to_db(serial, name, course, date):
    sheet = get_records_sheet()
    timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    sheet.append_row([serial, name, course, date, timestamp])

# ==========================================
# 2. ฟังก์ชันวาดภาพ (ปรับปรุงใหม่: แยกอิสระ)
# ==========================================
def create_certificate_image(template_source, name, course_name, date_str, serial,
                             name_size, course_size, name_y_adjust, course_y_adjust, is_name_bold):
    try:
        img = Image.open(template_source)
    except:
        if os.path.exists("template.png"):
            img = Image.open("template.png")
        else:
            return None
        
    draw = ImageDraw.Draw(img)
    W, H = img.size
    
    # --- ตั้งค่าฟอนต์ ---
    font_regular_path = "THSarabunNew.ttf"
    font_bold_path = "THSarabunNew Bold.ttf" 
    
    # 1. ฟอนต์ชื่อ (Name)
    try:
        if is_name_bold and os.path.exists(font_bold_path):
            font_name = ImageFont.truetype(font_bold_path, name_size)
        else:
            font_name = ImageFont.truetype(font_regular_path, name_size)
    except:
        font_name = ImageFont.load_default()

    # 2. ฟอนต์รายละเอียด (Course)
    try:
        font_detail = ImageFont.truetype(font_regular_path, course_size)
        font_serial = ImageFont.truetype(font_regular_path, 36)
    except:
        font_detail = ImageFont.load_default()
        font_serial = ImageFont.load_default()
    
    # --- วาดชื่อคน (Name) ---
    # อ้างอิงจากกึ่งกลางภาพ (H/2) ลบด้วย 80 แล้วบวกค่าปรับจูน
    left, top, right, bottom = draw.textbbox((0, 0), name, font=font_name)
    w = right - left
    h = bottom - top
    name_y = ((H - h) / 2) - 80 + name_y_adjust 
    draw.text(((W - w) / 2, name_y), name, font=font_name, fill=(0, 0, 0))
    
    # --- วาดรายละเอียด (Course) ---
    # 🟢 แก้ไขใหม่: ไม่อ้างอิง name_y แล้ว!
    # อ้างอิงจากกึ่งกลางภาพ (H/2) บวกด้วย 50 (ให้เริ่มต่ำกว่ากลางภาพนิดหน่อย) แล้วบวกค่าปรับจูน
    base_course_y = (H / 2) + 50 
    current_y = base_course_y + course_y_adjust
    
    lines = course_name.split('\n')
    for line in lines:
        left, top, right, bottom = draw.textbbox((0, 0), line.strip(), font=font_detail)
        w_line = right - left
        draw.text(((W - w_line) / 2, current_y), line.strip(), font=font_detail, fill=(50, 50, 50))
        current_y += (course_size + 15) 
        
    # --- วาดวันที่ & Serial ---
    date_text = f"ให้ไว้ ณ วันที่ {date_str}"
    draw.text((350, H - 200), date_text, font=font_serial, fill=(30, 30, 30))
    
    serial_text = f"Ref: {serial}"
    left, top, right, bottom = draw.textbbox((0, 0), serial_text, font=font_serial)
    w_serial = right - left
    draw.text((W - w_serial - 350, H - 200), serial_text, font=font_serial, fill=(30, 30, 30))
    
    return img

# ==========================================
# 3. ตั้งค่าหน้าเพจ & CSS
# ==========================================
st.set_page_config(page_title="ระบบออกเกียรติบัตรองค์กร", page_icon="🎓", layout="wide", initial_sidebar_state="expanded")

custom_css = """
<style>
    #MainMenu {visibility: hidden;}
    [data-testid="stDeployButton"] {display:none;}
    footer {visibility: hidden;}
    
    .stTextInput > div > div > input, 
    .stDateInput > div > div > input, 
    .stTextArea > div > div > textarea,
    .stSelectbox > div > div > div {
        background-color: #f8f9fc !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 10px !important;
        padding: 10px 15px !important;
    }
    .stTextInput > div > div > input:focus, 
    .stTextArea > div > div > textarea:focus {
        border-color: #00796B !important;
        box-shadow: 0 0 0 3px rgba(0, 121, 107, 0.15) !important;
    }
</style>
"""
st.markdown(custom_css, unsafe_allow_html=True)

# ==========================================
# 🌟 ระบบ Login
# ==========================================
if "logged_in" not in st.session_state:
    st.session_state["logged_in"] = False
if "username" not in st.session_state:
    st.session_state["username"] = ""

if not st.session_state["logged_in"]:
    st.write("<br><br>", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<h2 style='text-align: center;'>🔐 ระบบจัดการเกียรติบัตร</h2>", unsafe_allow_html=True)
        with st.container(border=True):
            user_input = st.text_input("Username:")
            pass_input = st.text_input("Password:", type="password")
            if st.button("Login", type="primary", use_container_width=True):
                try:
                    users_sheet = get_users_sheet()
                    users_data = users_sheet.get_all_records()
                    found = False
                    for u in users_data:
                        if str(u.get('username')) == user_input and str(u.get('password')) == pass_input:
                            st.session_state["logged_in"] = True
                            st.session_state["username"] = user_input
                            found = True
                            st.rerun()
                    if not found:
                        st.error("❌ ข้อมูลไม่ถูกต้อง")
                except Exception as e:
                    st.error(f"Login Error: {e}")
    st.stop()

# ==========================================
# 🌟 เมนู Sidebar
# ==========================================
with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3135/3135715.png", width=80) 
    st.success(f"👤 User: **{st.session_state['username']}**")
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state["logged_in"] = False
        st.rerun()
    st.divider()
    
    menu = st.radio("เลือกเมนู:", ["🎓 ออกเกียรติบัตร", "🗄️ ฐานข้อมูล", "🔑 เปลี่ยนรหัสผ่าน"])
    
    if menu == "🎓 ออกเกียรติบัตร":
        st.divider()
        st.markdown("### 🎨 1. ตั้งค่า Template")
        uploaded_template = st.file_uploader("📂 อัปโหลดรูปพื้นหลัง:", type=['png', 'jpg'])
        current_template = uploaded_template if uploaded_template else "template.png"
        
        st.divider()
        st.markdown("### 🔠 2. ตั้งค่าตัวอักษร")
        
        st.caption("👤 **ส่วนชื่อ-นามสกุล**")
        name_font_size = st.slider("ขนาดชื่อ:", 50, 300, 100)
        name_is_bold = st.checkbox("ตัวหนา (Bold)", value=True)
        name_y_adjust = st.slider("ตำแหน่งชื่อ (จากกึ่งกลาง):", -1000, 1000, 0, help="ลบ=เลื่อนขึ้น, บวก=เลื่อนลง")
        
        st.caption("📝 **ส่วนรายละเอียดหลักสูตร**")
        course_font_size = st.slider("ขนาดรายละเอียด:", 30, 150, 50)
        # 🟢 เปลี่ยนคำอธิบายใหม่ เพื่อให้เข้าใจว่า "อิสระ" แล้ว
        course_y_adjust = st.slider("ตำแหน่งรายละเอียด (จากกึ่งกลาง):", -1000, 1000, 0, help="ปรับตำแหน่งอิสระ ไม่เกี่ยวกับชื่อแล้ว")

# ==========================================
# 5. หน้าจอออกเกียรติบัตร
# ==========================================
if menu == "🎓 ออกเกียรติบัตร":
    st.title("🎓 Certificate Automation System")
    
    with st.container(border=True):
        col1, col2 = st.columns([2, 1])
        course_name = col1.text_area("📌 รายละเอียดหลักสูตร:", height=115)
        issue_date = col2.date_input("🗓️ วันที่:")
        date_str_formatted = issue_date.strftime('%d/%m/%Y')
    
    tab1, tab2 = st.tabs(["👤 รายบุคคล (Preview & Save)", "📁 แบบกลุ่ม (Excel Batch)"])
    
    # --- Tab 1: Single ---
    with tab1:
        single_name = st.text_input("ชื่อ-นามสกุล ผู้รับ:")
        
        st.write("เลือกประเภทไฟล์ที่ต้องการดาวน์โหลด:")
        col_c1, col_c2 = st.columns(2)
        with col_c1: need_png = st.checkbox("ไฟล์รูปภาพ (PNG)", value=True, key="s_png")
        with col_c2: need_pdf = st.checkbox("ไฟล์เอกสาร (PDF)", value=False, key="s_pdf")

        col_preview, col_save = st.columns(2)
        
        # 1. Preview
        with col_preview:
            if st.button("✨ ดูตัวอย่าง (Preview)", use_container_width=True):
                if single_name and course_name:
                    img = create_certificate_image(
                        current_template, single_name, course_name, date_str_formatted, "PREVIEW-001",
                        name_font_size, course_font_size, name_y_adjust, course_y_adjust, name_is_bold
                    )
                    st.image(img, caption="ภาพตัวอย่าง (ตำแหน่งอิสระ)", use_container_width=True)
                else:
                    st.warning("กรุณากรอกข้อมูลให้ครบ")

        # 2. Save
        with col_save:
            if st.button("💾 บันทึกและสร้างไฟล์จริง", type="primary", use_container_width=True):
                if not need_png and not need_pdf:
                    st.warning("⚠️ กรุณาเลือกประเภทไฟล์อย่างน้อย 1 อย่างครับ")
                elif single_name and course_name:
                    try:
                        serial = generate_serial()
                        save_to_db(serial, single_name, course_name, str(issue_date))
                        
                        img_final = create_certificate_image(
                            current_template, single_name, course_name, date_str_formatted, serial,
                            name_font_size, course_font_size, name_y_adjust, course_y_adjust, name_is_bold
                        )
                        
                        st.success(f"✅ บันทึกสำเร็จ! Ref: {serial}")
                        
                        if need_png:
                            buf_png = io.BytesIO()
                            img_final.save(buf_png, format="PNG")
                            st.download_button("📩 ดาวน์โหลด (PNG)", data=buf_png.getvalue(), file_name=f"{serial}.png", mime="image/png", use_container_width=True)
                        
                        if need_pdf:
                            buf_pdf = io.BytesIO()
                            img_final.convert('RGB').save(buf_pdf, format="PDF")
                            st.download_button("📄 ดาวน์โหลด (PDF)", data=buf_pdf.getvalue(), file_name=f"{serial}.pdf", mime="application/pdf", use_container_width=True)
                        
                    except Exception as e:
                        st.error(f"Error: {e}")
                else:
                    st.warning("กรุณากรอกข้อมูลให้ครบ")

    # --- Tab 2: Batch ---
    with tab2:
        st.info("💡 ระบบจะใช้การตั้งค่าฟอนต์และตำแหน่งจากเมนูซ้ายมือ")
        uploaded_excel = st.file_uploader("Upload Excel (ต้องมีช่อง Name)", type=["xlsx"])
        
        st.write("เลือกประเภทไฟล์ที่ต้องการใน ZIP (เลือกได้ทั้งคู่):")
        col_b1, col_b2 = st.columns(2)
        with col_b1: batch_png = st.checkbox("ไฟล์รูปภาพ (PNG)", value=True, key="b_png")
        with col_b2: batch_pdf = st.checkbox("ไฟล์เอกสาร (PDF)", value=False, key="b_pdf")
        
        if uploaded_excel and st.button("🚀 รันระบบ Batch (สร้างไฟล์ ZIP)"):
            if not batch_png and not batch_pdf:
                 st.warning("⚠️ กรุณาเลือกประเภทไฟล์อย่างน้อย 1 อย่างครับ")
            else:
                df = pd.read_excel(uploaded_excel)
                if 'Name' not in df.columns:
                    st.error("❌ ไม่พบคอลัมน์ 'Name'")
                else:
                    progress_bar = st.progress(0)
                    zip_buffer = io.BytesIO()
                    new_db_rows = []
                    
                    sheet = get_records_sheet()
                    records = sheet.get_all_records()
                    prefix = f"CERT-{datetime.datetime.now().strftime('%Y%m')}"
                    current_count = sum(1 for r in records if str(r.get('serial_number', '')).startswith(prefix))
                    
                    with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                        total = len(df)
                        for i, row in df.iterrows():
                            name = str(row['Name']).strip()
                            if not name or name == "nan": continue
                            
                            current_count += 1
                            serial = f"{prefix}-{current_count:04d}"
                            timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                            new_db_rows.append([serial, name, course_name, str(issue_date), timestamp])
                            
                            if uploaded_template:
                                uploaded_template.seek(0)
                                tmpl = uploaded_template
                            else:
                                tmpl = current_template
                                
                            img = create_certificate_image(
                                tmpl, name, course_name, date_str_formatted, serial,
                                name_font_size, course_font_size, name_y_adjust, course_y_adjust, name_is_bold
                            )
                            
                            if batch_png:
                                img_buf = io.BytesIO()
                                img.save(img_buf, format="PNG")
                                zip_file.writestr(f"{serial}_{name}.png", img_buf.getvalue())

                            if batch_pdf:
                                pdf_buf = io.BytesIO()
                                img.convert('RGB').save(pdf_buf, format="PDF")
                                zip_file.writestr(f"{serial}_{name}.pdf", pdf_buf.getvalue())
                            
                            progress_bar.progress((i + 1) / total, text=f"Creating: {name}")
                    
                    if new_db_rows:
                        sheet.append_rows(new_db_rows)
                    
                    st.success("✅ เสร็จสิ้น!")
                    st.download_button("⬇️ ดาวน์โหลด ZIP", data=zip_buffer.getvalue(), file_name="Certificates.zip", mime="application/zip")

# ==========================================
# 6. Admin & Other Menus
# ==========================================
elif menu == "🗄️ ฐานข้อมูล":
    st.title("🗄️ Database Management")
    with st.container(border=True):
        sheet = get_records_sheet()
        records = sheet.get_all_records()
        df = pd.DataFrame(records) if records else pd.DataFrame()
        edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True, height=500)
        
        if st.button("💾 บันทึกการแก้ไขลง Google Sheets", type="primary"):
            sheet.clear()
            sheet.append_row(list(edited_df.columns))
            if not edited_df.empty:
                sheet.append_rows(edited_df.values.tolist())
            st.success("✅ บันทึกเรียบร้อย!")

elif menu == "🔑 เปลี่ยนรหัสผ่าน":
    st.title("🔑 Change Password")
    with st.container(border=True):
        old_p = st.text_input("รหัสเดิม:", type="password")
        new_p = st.text_input("รหัสใหม่:", type="password")
        conf_p = st.text_input("ยืนยัน:", type="password")
        if st.button("บันทึก"):
            if new_p != conf_p:
                st.error("รหัสผ่านไม่ตรงกัน")
            else:
                users_sheet = get_users_sheet()
                users = users_sheet.get_all_records()
                found = False
                for i, u in enumerate(users):
                    if str(u.get('username')) == st.session_state["username"] and str(u.get('password')) == old_p:
                        users_sheet.update_cell(i + 2, 2, new_p)
                        st.success("เปลี่ยนรหัสผ่านสำเร็จ!")
                        found = True
                        break
                if not found: st.error("รหัสเดิมผิด")