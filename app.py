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
# 1. ระบบเชื่อมต่อ Google Sheets API (ฉลาดเลือก)
# ==========================================
SHEET_NAME = "Corp_Cert_DB"

@st.cache_resource
def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    try:
        # 🟢 1. เช็คก่อนว่ามีไฟล์กุญแจในเครื่องไหม (สำหรับตอนรันในคอมตัวเอง Local)
        if os.path.exists("service_account.json"):
            creds = Credentials.from_service_account_file("service_account.json", scopes=scopes)
            
        # ☁️ 2. ถ้าไม่มีในเครื่อง ให้ไปดูในตู้เซฟ Secrets (สำหรับตอนรันบนเว็บ Streamlit Cloud)
        else:
            try:
                if "GCP_CREDENTIALS" in st.secrets:
                    creds_dict = json.loads(st.secrets["GCP_CREDENTIALS"])
                    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
                else:
                    st.error("❌ ไม่พบทั้งไฟล์ service_account.json และ Secrets บน Cloud")
                    st.stop()
            except FileNotFoundError:
                st.error("❌ ไม่พบไฟล์กุญแจ service_account.json ในโฟลเดอร์โปรเจกต์")
                st.stop()
            
        return gspread.authorize(creds)
        
    except Exception as e:
        st.error(f"❌ เกิดข้อผิดพลาดในการเชื่อมต่อ Google API: {e}")
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
# 2. ฟังก์ชันหลักสำหรับวาดภาพเกียรติบัตร 
# ==========================================
def create_certificate_image(template_source, font_path, name, course_name, date_str, serial):
    # รองรับทั้ง "ชื่อไฟล์" (String) และ "ไฟล์ที่อัปโหลด" (UploadedFile)
    try:
        img = Image.open(template_source)
    except:
        # ถ้าเปิดไม่ได้ ให้ลองใช้ไฟล์ Default
        if os.path.exists("template.png"):
            img = Image.open("template.png")
        else:
            st.error("❌ หาไฟล์ template.png ไม่เจอ กรุณาเอาไฟล์วางในโฟลเดอร์ด้วยครับ")
            st.stop()
        
    draw = ImageDraw.Draw(img)
    W, H = img.size
    
    # กำหนดฟอนต์ (ต้องมีไฟล์ .ttf ในโฟลเดอร์)
    try:
        font_name = ImageFont.truetype(font_path, 80)    
        font_detail = ImageFont.truetype(font_path, 40)  
        font_serial = ImageFont.truetype(font_path, 36)
    except:
        st.error(f"❌ ไม่พบฟอนต์ {font_path} ในโฟลเดอร์")
        st.stop()
    
    # 1. ชื่อคน (จัดกึ่งกลาง)
    left, top, right, bottom = draw.textbbox((0, 0), name, font=font_name)
    w = right - left
    h = bottom - top
    name_y = ((H - h) / 2) - 80 
    draw.text(((W - w) / 2, name_y), name, font=font_name, fill=(0, 0, 0))
    
    # 2. ชื่อหลักสูตร (ใต้ชื่อคน)
    lines = course_name.split('\n')
    current_y = name_y + 120 
    for line in lines:
        left, top, right, bottom = draw.textbbox((0, 0), line.strip(), font=font_detail)
        w_line = right - left
        draw.text(((W - w_line) / 2, current_y), line.strip(), font=font_detail, fill=(50, 50, 50))
        current_y += 60 
        
    # 3. วันที่ (มุมซ้ายล่าง)
    date_text = f"ให้ไว้ ณ วันที่ {date_str}"
    draw.text((350, H - 200), date_text, font=font_serial, fill=(30, 30, 30))
    
    # 4. รหัส Serial (มุมขวาล่าง)
    serial_text = f"Ref: {serial}"
    left, top, right, bottom = draw.textbbox((0, 0), serial_text, font=font_serial)
    w_serial = right - left
    draw.text((W - w_serial - 350, H - 200), serial_text, font=font_serial, fill=(30, 30, 30))
    
    return img

# ==========================================
# 3. ตั้งค่าหน้าเพจ & ตกแต่ง CSS
# ==========================================
st.set_page_config(page_title="ระบบออกเกียรติบัตรองค์กร", page_icon="🎓", layout="wide", initial_sidebar_state="expanded")

custom_css = """
<style>
    #MainMenu {visibility: hidden;}
    [data-testid="stDeployButton"] {display:none;}
    
    /* ลบบรรทัดที่สั่งซ่อน Toolbar ออกแล้ว เพื่อให้ปุ่มลูกศรกลับมาครับ */
    
    footer {visibility: hidden;}

    .custom-banner {
        background: linear-gradient(135deg, #00796B, #26A69A);
        padding: 30px;
        border-radius: 15px;
        color: white;
        text-align: center;
        box-shadow: 0 10px 20px rgba(0, 0, 0, 0.1);
        margin-bottom: 25px;
    }
    .custom-banner h1 {
        color: white;
        margin: 0;
        font-size: 2.8rem;
        font-weight: 700;
        text-shadow: 2px 2px 4px rgba(0,0,0,0.2);
    }
    .custom-banner p {
        margin: 10px 0 0 0;
        font-size: 1.2rem;
        opacity: 0.9;
    }

    [data-testid="baseButton-primary"] {
        background-color: #00796B !important;
        border: none !important;
        border-radius: 8px !important;
        transition: all 0.3s ease !important;
        font-weight: bold !important;
        padding: 0.5rem 1rem !important;
    }
    [data-testid="baseButton-primary"]:hover {
        background-color: #005A4F !important;
        transform: translateY(-3px) !important;
        box-shadow: 0 6px 12px rgba(0, 121, 107, 0.3) !important;
    }

    /* Modern Inputs */
    .stTextInput > div > div > input, 
    .stDateInput > div > div > input, 
    .stTextArea > div > div > textarea,
    .stSelectbox > div > div > div {
        background-color: #f8f9fc !important;
        border: 1px solid #e2e8f0 !important;
        border-radius: 10px !important;
        padding: 10px 15px !important;
        font-size: 16px !important;
        transition: all 0.3s ease !important;
    }
    
    .stTextInput > div > div > input:focus, 
    .stDateInput > div > div > input:focus, 
    .stTextArea > div > div > textarea:focus,
    .stSelectbox > div > div > div:focus {
        border-color: #00796B !important;
        box-shadow: 0 0 0 3px rgba(0, 121, 107, 0.15) !important;
        background-color: #ffffff !important;
    }
    
    .stTextArea label, .stDateInput label, .stSelectbox label {
        font-weight: 600 !important;
        color: #2c3e50 !important;
        margin-bottom: 5px !important;
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
        st.markdown("""
        <div class="custom-banner" style="margin-top: 0;">
            <h1 style="font-size: 2rem;">🔐 ระบบจัดการเกียรติบัตร</h1>
            <p style="font-size: 1rem;">เข้าสู่ระบบ (เชื่อมต่อ Google Sheets)</p>
        </div>
        """, unsafe_allow_html=True)
        
        with st.container(border=True):
            user_input = st.text_input("👤 ชื่อผู้ใช้งาน (Username):")
            pass_input = st.text_input("🔑 รหัสผ่าน (Password):", type="password")
            
            if st.button("เข้าสู่ระบบ (Login)", type="primary", use_container_width=True):
                try:
                    users_sheet = get_users_sheet()
                    users_data = users_sheet.get_all_records()
                    login_success = False
                    for u in users_data:
                        if str(u.get('username')) == user_input and str(u.get('password')) == pass_input:
                            login_success = True
                            break   
                    if login_success:
                        st.session_state["logged_in"] = True
                        st.session_state["username"] = user_input
                        st.rerun()
                    else:
                        st.error("❌ ชื่อผู้ใช้งาน หรือ รหัสผ่านไม่ถูกต้อง!")
                except Exception as e:
                    st.error(f"เกิดข้อผิดพลาดในการดึงข้อมูล Login: {e}")
                    
    st.stop()

# ==========================================
# 🌟 ส่วนหัวและเมนู (Sidebar)
# ==========================================
st.write("<br>", unsafe_allow_html=True) 
top_col1, top_col2, top_col3 = st.columns([7, 1, 2])
with top_col3:
    if st.button("🔴 ปิดหน้าต่าง (Close)", key="shutdown_top", use_container_width=True):
        st.success("✅ กากบาท (X) ปิดหน้าต่างเบราว์เซอร์นี้ได้เลยครับ")
        time.sleep(3)

with st.sidebar:
    st.image("https://cdn-icons-png.flaticon.com/512/3135/3135715.png", width=80) 
    st.success(f"👤 ผู้ใช้งาน: **{st.session_state['username']}**")
    
    if st.button("🚪 ออกจากระบบ (Logout)", use_container_width=True):
        st.session_state["logged_in"] = False
        st.session_state["username"] = ""
        st.rerun()

    st.divider()
    st.markdown("### ⚙️ เมนูจัดการระบบ")
    menu = st.radio("เลือกการทำงาน:", ["🎓 ออกเกียรติบัตร", "🗄️ ฐานข้อมูล (Google Sheets)", "🔑 เปลี่ยนรหัสผ่าน"], label_visibility="collapsed")
    
    # --- 🌟 ส่วนอัปโหลด Template (แสดงเฉพาะหน้าออกเกียรติบัตร) ---
    if menu == "🎓 ออกเกียรติบัตร":
        st.divider()
        st.markdown("### 🎨 ตั้งค่าใบเกียรติบัตร")
        st.info("คุณสามารถอัปโหลดไฟล์ภาพพื้นหลัง (Template) ของคุณเองได้ที่นี่")
        
        # 1. ปุ่มโหลด Mockup
        if os.path.exists("template.png"):
            with open("template.png", "rb") as file:
                st.download_button(
                    label="⬇️ ดาวน์โหลดภาพตัวอย่าง (Mockup)",
                    data=file,
                    file_name="template_mockup.png",
                    mime="image/png",
                    help="ดาวน์โหลดไฟล์นี้ไปเป็นแบบ เพื่อดูตำแหน่งการวางข้อความ",
                    use_container_width=True
                )
            
        # 2. ปุ่มอัปโหลดไฟล์เอง
        uploaded_template = st.file_uploader("📂 อัปโหลด Template ใหม่ (png/jpg):", type=['png', 'jpg', 'jpeg'])
        
        if uploaded_template is not None:
            # ถ้ามีการอัปโหลด ให้ใช้ไฟล์นั้น
            current_template = uploaded_template
            st.success("✅ ใช้ Template ที่อัปโหลด")
        else:
            # ถ้าไม่มี ให้ใช้ไฟล์ Default (String path)
            current_template = "template.png"

# ==========================================
# 5. หน้าจอออกเกียรติบัตร (User Mode)
# ==========================================
if menu == "🎓 ออกเกียรติบัตร":
    st.markdown("""
        <div class="custom-banner">
            <h1>🎓 Certificate Automation System</h1>
            <p>ระบบสร้างและแพ็กไฟล์เกียรติบัตรอัตโนมัติ (เชื่อมต่อ Cloud)</p>
        </div>
    """, unsafe_allow_html=True)
    
    with st.container(border=True):
        col1, col2 = st.columns([2, 1])
        course_name = col1.text_area("📌 หัวข้อการอบรม (กด Enter พิมพ์หลายบรรทัดได้):", height=115, help="พิมพ์ชื่อหลักสูตร หรือรายละเอียดที่ต้องการให้ปรากฏกลางเกียรติบัตร")
        issue_date = col2.date_input("🗓️ กำหนดวันที่ให้เกียรติบัตร:", help="คลิกที่ไอคอนปฏิทินมุมขวาของกล่อง เพื่อเลือกวันที่อย่างรวดเร็ว")
        date_str_formatted = issue_date.strftime('%d/%m/%Y')
    
    st.write("") 
    tab1, tab2 = st.tabs(["👤 พิมพ์รายบุคคล (Single)", "📁 โหลดไฟล์ Excel (Bulk Generate)"])
    
    # --- Tab 1: พิมพ์รายบุคคล ---
    with tab1:
        with st.container(border=True):
            st.markdown("#### 🖨️ ออกเอกสารเฉพาะบุคคล")
            single_name = st.text_input("👤 ชื่อ-นามสกุล ผู้เข้ารับการอบรม:")
            if st.button("✨ สร้างเกียรติบัตร", type="primary"): 
                if single_name and course_name:
                    try:
                        serial = generate_serial()
                        save_to_db(serial, single_name, course_name, str(issue_date))
                        
                        # สร้างภาพ (ส่ง current_template ที่เลือกไว้ใน Sidebar)
                        img = create_certificate_image(current_template, "THSarabunNew.ttf", single_name, course_name, date_str_formatted, serial)
                        
                        st.success(f"✅ บันทึกข้อมูลสำเร็จ! รหัสอ้างอิง: {serial}")
                        st.image(img, caption=f"ตัวอย่างผลลัพธ์", use_container_width=True)
                        
                        buf_png = io.BytesIO()
                        img.save(buf_png, format="PNG")
                        buf_pdf = io.BytesIO()
                        img_pdf = img.convert('RGB')
                        img_pdf.save(buf_pdf, format="PDF", resolution=100.0)
                        
                        dl_col1, dl_col2 = st.columns(2)
                        with dl_col1:
                            st.download_button("📩 ดาวน์โหลด (PNG)", data=buf_png.getvalue(), file_name=f"Certificate_{serial}_{single_name}.png", mime="image/png", use_container_width=True)
                        with dl_col2:
                            st.download_button("📄 ดาวน์โหลด (PDF)", data=buf_pdf.getvalue(), file_name=f"Certificate_{serial}_{single_name}.pdf", mime="application/pdf", use_container_width=True)

                    except Exception as e:
                        st.error(f"❌ เกิดข้อผิดพลาด: {e}")
                else:
                    st.warning("⚠️ กรุณากำหนดข้อมูลให้ครบถ้วนก่อนครับ")
                
    # --- Tab 2: พิมพ์แบบกลุ่ม ---
    with tab2:
        with st.container(border=True):
            st.markdown("#### 📦 ระบบสร้างเกียรติบัตรแบบกลุ่ม (Batch Processing)")
            uploaded_file = st.file_uploader("ลากไฟล์ Excel มาวางตรงนี้ (ต้องมีคอลัมน์ Name)", type=["xlsx"])
            
            if uploaded_file and course_name:
                df = pd.read_excel(uploaded_file)
                if 'Name' not in df.columns:
                    st.error("❌ ระบบหาคอลัมน์คำว่า 'Name' ไม่เจอครับ")
                else:
                    st.success(f"✅ โหลดไฟล์สำเร็จ! พบผู้เข้าอบรมทั้งหมด: **{len(df)}** ท่าน")
                    st.markdown("**📋 ตารางตัวอย่างรายชื่อจากไฟล์ Excel:**")
                    st.dataframe(df.head(5), use_container_width=True)
                    
                    st.divider()
                    st.markdown("##### เลือกรูปแบบไฟล์เกียรติบัตรที่จะอยู่ในไฟล์ ZIP:")
                    export_format = st.radio("รูปแบบไฟล์:", ["📄 ไฟล์เอกสาร PDF (แนะนำสำหรับการสั่งพิมพ์)", "🖼️ ไฟล์รูปภาพ PNG (แนะนำสำหรับส่งต่อผ่าน Social/LINE)"], horizontal=True, label_visibility="collapsed")
                    
                    st.divider()
                    col_p1, col_p2 = st.columns([1, 2])
                    with col_p1:
                        if st.button("🔎 ดูภาพตัวอย่าง (รายชื่อที่ 1)"):
                            try:
                                preview_name = str(df.iloc[0]['Name']).strip()
                                # พรีวิวโดยใช้ Template ที่เลือก
                                img_preview = create_certificate_image(current_template, "THSarabunNew.ttf", preview_name, course_name, date_str_formatted, "CERT-PREVIEW-001")
                                st.image(img_preview, caption=f"ตัวอย่างการจัดวาง", use_container_width=True)
                            except Exception as e:
                                st.error(f"เกิดข้อผิดพลาด: {e}")
                    
                    st.divider()
                    if st.button("🚀 แพ็กรายชื่อทั้งหมดเป็นไฟล์ ZIP", type="primary", use_container_width=True):
                        progress_text = "กำลังสร้างและแพ็กไฟล์... (กรุณารอสักครู่)"
                        my_bar = st.progress(0, text=progress_text)
                        zip_buffer = io.BytesIO()
                        
                        try:
                            # เตรียมข้อมูลรันเลข Batch
                            sheet = get_records_sheet()
                            records = sheet.get_all_records()
                            prefix = f"CERT-{datetime.datetime.now().strftime('%Y%m')}"
                            current_count = sum(1 for r in records if str(r.get('serial_number', '')).startswith(prefix))
                            new_db_rows = [] 
                            
                            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                                total_rows = len(df)
                                for index, row in df.iterrows():
                                    name = str(row['Name']).strip()
                                    if not name or name == "nan": continue 
                                    
                                    current_count += 1
                                    serial = f"{prefix}-{current_count:04d}"
                                    timestamp = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                                    new_db_rows.append([serial, name, course_name, str(issue_date), timestamp])
                                    
                                    # จัดการ Pointer ของไฟล์ Template (สำคัญมากกรณีไฟล์อัปโหลด)
                                    if uploaded_template:
                                        uploaded_template.seek(0)
                                        template_source_for_func = uploaded_template
                                    else:
                                        template_source_for_func = current_template

                                    img = create_certificate_image(template_source_for_func, "THSarabunNew.ttf", name, course_name, date_str_formatted, serial)
                                    
                                    img_buffer = io.BytesIO()
                                    if "PDF" in export_format:
                                        img_pdf = img.convert('RGB')
                                        img_pdf.save(img_buffer, format="PDF", resolution=100.0)
                                        zip_file.writestr(f"Certificate_{serial}_{name}.pdf", img_buffer.getvalue())
                                    else:
                                        img.save(img_buffer, format="PNG")
                                        zip_file.writestr(f"Certificate_{serial}_{name}.png", img_buffer.getvalue())
                                    
                                    percent_complete = int(((index + 1) / total_rows) * 100)
                                    my_bar.progress(percent_complete, text=f"กำลังสร้าง: {name} ({index+1}/{total_rows})")
                            
                            # บันทึกลง Google Sheets ทีเดียว
                            if new_db_rows:
                                my_bar.progress(99, text="กำลังอัปเดตฐานข้อมูล Google Sheets...")
                                sheet.append_rows(new_db_rows)
                                
                            my_bar.empty()
                            st.success("🎉 บันทึกข้อมูลและแพ็กไฟล์เสร็จสมบูรณ์!")
                            st.download_button("⬇️ คลิกที่นี่เพื่อดาวน์โหลดไฟล์ ZIP", data=zip_buffer.getvalue(), file_name=f"Certificates_{issue_date}.zip", mime="application/zip", use_container_width=True, type="primary")
                        except Exception as e:
                            st.error(f"❌ เกิดข้อผิดพลาด: {e}")

# ==========================================
# 6. หน้าจอ Admin CRUD
# ==========================================
elif menu == "🗄️ ฐานข้อมูล (Google Sheets)":
    st.markdown("""
        <div class="custom-banner" style="background: linear-gradient(135deg, #424242, #212121);">
            <h1>🗄️ Database Management</h1>
            <p>แก้ไขข้อมูลบน Google Sheets โดยตรงผ่านหน้านี้</p>
        </div>
    """, unsafe_allow_html=True)
    
    with st.container(border=True):
        st.info("💡 เมื่อกดบันทึก ข้อมูลใน Google Sheets จะถูกอัปเดตใหม่ทั้งหมดตามตารางนี้")
        sheet = get_records_sheet()
        records = sheet.get_all_records()
        if records:
            df_db = pd.DataFrame(records)
        else:
            df_db = pd.DataFrame(columns=["serial_number", "full_name", "course_name", "issue_date", "created_at"])
        edited_df = st.data_editor(df_db, num_rows="dynamic", use_container_width=True, key="db_editor", height=500)
        if st.button("💾 บันทึกทับข้อมูลลง Google Sheets", type="primary"):
            sheet.clear()
            sheet.append_row(list(edited_df.columns))
            if not edited_df.empty:
                sheet.append_rows(edited_df.values.tolist())
            st.success("✅ อัปเดตข้อมูลขึ้น Google Sheets สำเร็จ!")

# ==========================================
# 7. หน้าจอเปลี่ยนรหัสผ่าน
# ==========================================
elif menu == "🔑 เปลี่ยนรหัสผ่าน":
    st.markdown("""
        <div class="custom-banner" style="background: linear-gradient(135deg, #E65100, #F57C00);">
            <h1>🔑 Change Password</h1>
            <p>เปลี่ยนรหัสผ่านสำหรับการเข้าสู่ระบบ</p>
        </div>
    """, unsafe_allow_html=True)

    with st.container(border=True):
        old_pass = st.text_input("รหัสผ่านเดิม:", type="password")
        new_pass = st.text_input("รหัสผ่านใหม่:", type="password")
        confirm_pass = st.text_input("ยืนยันรหัสผ่านใหม่:", type="password")
        
        if st.button("💾 บันทึกรหัสผ่านใหม่", type="primary"):
            if new_pass == "" or confirm_pass == "":
                st.warning("⚠️ กรุณากรอกรหัสผ่านใหม่ให้ครบถ้วน")
            elif new_pass != confirm_pass:
                st.error("❌ รหัสผ่านใหม่และการยืนยันรหัสผ่านไม่ตรงกัน!")
            else:
                users_sheet = get_users_sheet()
                users_data = users_sheet.get_all_records()
                found = False
                for i, u in enumerate(users_data):
                    if str(u.get('username')) == st.session_state["username"] and str(u.get('password')) == old_pass:
                        users_sheet.update_cell(i + 2, 2, new_pass)
                        st.success("✅ เปลี่ยนรหัสผ่านและบันทึกลง Google Sheets เรียบร้อยแล้ว!")
                        found = True
                        break
                if not found:
                    st.error("❌ รหัสผ่านเดิมไม่ถูกต้อง!")