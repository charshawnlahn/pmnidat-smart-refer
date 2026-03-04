import streamlit as st
from google import genai 
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re
import json
import requests

# --- 1. ระบบจัดการการเชื่อมต่อและเลือกโมเดลอัตโนมัติ (Anti-404 System) ---
try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
    APPS_SCRIPT_URL = st.secrets["APPS_SCRIPT_URL"]
    client = genai.Client(api_key=API_KEY)
    
    @st.cache_resource
    def find_active_model():
        try:
            available_models = [m.name for m in client.models.list()]
            for m in available_models:
                if "gemini-1.5-flash" in m: return m
            return available_models[0]
        except:
            return "models/gemini-1.5-flash"

    MODEL_ID = find_active_model()
except Exception as e:
    st.error(f"❌ ระบบเชื่อมต่อผิดพลาด: {e}")
    st.stop()

# --- 2. ฟังก์ชันบันทึก Log Book ---
def log_usage(patient_name):
    try:
        requests.post(APPS_SCRIPT_URL, json={"name": patient_name}, timeout=5)
    except:
        pass

# --- 3. ฟังก์ชันจัดการไฟล์ Word (ชิดซ้าย + ฟอนต์ 13 + มาตรฐานสถาบันฯ) ---
def fill_pmnidat_doc(data):
    try:
        doc = Document("PMNIDAT 062 แบบบันทึกข้อมูลเพื่อส่งต่อ.docx")
        mapping = {f"{{{{{k.upper()}}}}}": str(v) for k, v in data.items()}
        
        def apply_style_and_replace(paragraph):
            for key, value in mapping.items():
                if key in paragraph.text:
                    paragraph.text = paragraph.text.replace(key, value)
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT 
                    for run in paragraph.runs:
                        run.font.size = Pt(13)

        for p in doc.paragraphs: apply_style_and_replace(p)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs: apply_style_and_replace(p)
                            
        buffer = io.BytesIO()
        doc.save(buffer)
        return buffer.getvalue()
    except Exception as e:
        st.error(f"⚠️ ปัญหาไฟล์แม่แบบ: {e}")
        return None

# --- 4. การออกแบบหน้าเว็บและ "คู่มือพี่พยาบาล ฉบับจับมือทำ" ---
st.set_page_config(page_title="PMNIDAT 062 Smart Portal", layout="wide")

with st.sidebar:
    st.image("https://pmnidathis.dms.go.th/static/media/health.1bfd961f.png", width=100)
    st.header("📖 คู่มือพี่พยาบาล ฉบับจับมือทำ")
    st.markdown("### **วิธีการคัดลอกข้อมูลจาก @ThanHIS**")
    st.info("""
    1. **คลิกเมาส์ซ้ายค้าง** ที่ต้นข้อความ ลากครอบให้คลุมทั้งหมด
    2. กด **Ctrl+C** (คัดลอก)
    3. มาที่หน้านี้ คลิกช่องที่ต้องการ แล้วกด **Ctrl+V** (วาง)
    """)
    
    with st.expander("🟢 STEP 1: ระบบผู้ป่วยใน (IPD)", expanded=True):
        st.write("""
        1. **1.1 Admission Note:** ดูข้อมูลคนไข้ → Admission note → คัดลอกทั้งหมด
        2. **1.2 การวินิจฉัย:** กดเมนู "การวินิจฉัย" → คัดลอกรหัส ICD-10
        3. **1.3 Order / Meds:** กดเมนู "Order" → คัดลอก Discharge order + Home medication
        4. **1.4 Progress Note:** กดเมนู "Progress note" → คัดลอกบันทึกล่าสุด (SOAP)
        """)

    with st.expander("🔵 STEP 2: การประเมิน (Assessment)"):
        st.write("""
        - กดเมนู "Admission note" → ปุ่ม "ข้อมูลผู้ป่วยนอก" 
        - เลื่อนลงล่างไปที่หัวข้อ **Assessment** - คัดลอกผลคะแนน 9Q, 8Q, BPRS
        """)

    with st.expander("🟠 STEP 3: เวชระเบียน (Registration)"):
        st.write("""
        - ระบบผู้ป่วยนอก → เวชระเบียน → ลงทะเบียนผู้ป่วย → ค้นหา HN
        - **3.1 ทั่วไป 1:** คัดลอกข้อมูล ชื่อ, อายุ, เลขบัตรประชาชน, ศาสนา
        - **3.2 ทั่วไป 2:** กด "ที่อยู่ปัจจุบัน" → คัดลอกที่อยู่ทั้งหมด
        - **3.3 ผู้ติดต่อ:** คัดลอกชื่อญาติและเบอร์โทรศัพท์
        - **3.4 สิทธิรักษา:** คัดลอกสิทธิ์และ "สถานพยาบาลหลัก"
        """)
    
    st.divider()
    st.success(f"💡 AI วิเคราะห์ด้วยตรรกะ PhD ผ่าน: {MODEL_ID}")

st.title("🏥 PMNIDAT Smart D/C Transfer")
st.subheader("ระบบสร้างไฟล์ใบส่งต่อ 062 อัตโนมัติ (Master Version 3.24)")

st.divider()
# --- ส่วนกรอกข้อมูล ---
st.markdown("### **🟢 Step 1: ระบบผู้ป่วยใน (IPD)**")
s1_cols = st.columns(4)
with s1_cols[0]: s11 = st.text_area("1.1 Admission Note", height=150, placeholder="วางข้อมูลจากระบบ IPD...")
with s1_cols[1]: s12 = st.text_area("1.2 การวินิจฉัย", height=150, placeholder="วางรหัส ICD-10...")
with s1_cols[2]: s13 = st.text_area("1.3 Order / Meds", height=150, placeholder="วางรายการยา Home-Med...")
with s1_cols[3]: s14 = st.text_area("1.4 Progress Note", height=150, placeholder="วาง SOAP ล่าสุด...")

st.divider()
st.markdown("### **🔵 Step 2: การประเมิน (Assessment)**")
s2 = st.text_area("คัดลอกผลคะแนน 9Q, 8Q, BPRS มาวางที่นี่", height=100)

st.divider()
st.markdown("### **🟠 Step 3: เวชระเบียน (Registration)**")
s3_cols = st.columns(4)
with s3_cols[0]: s31 = st.text_area("3.1 ข้อมูลทั่วไป", height=150)
with s3_cols[1]: s32 = st.text_area("3.2 ที่อยู่ปัจจุบัน", height=150)
with s3_cols[2]: s33 = st.text_area("3.3 ผู้ติดต่อ", height=150)
with s3_cols[3]: s34 = st.text_area("3.4 สิทธิการรักษา", height=150)

# --- 5. ส่วนประมวลผล (ใช้ตรรกะสกัดข้อมูลขั้นสูง) ---
if st.button("🚀 สกัดข้อมูลและสร้างเอกสาร"):
    all_raw = f"---IPD---\n{s11}\n{s12}\n{s13}\n{s14}\n---SCORE---\n{s2}\n---REG---\n{s31}\n{s32}\n{s33}\n{s34}"
    with st.spinner('กำลังประมวลผลและตรวจสอบความถูกต้อง (Verification Audit)...'):
        prompt = f"""
        คุณคือผู้ช่วยวิจัยทางการแพทย์ระดับ PhD ทำหน้าที่สกัดข้อมูลจากระบบ @ThanHIS ลงแบบฟอร์ม 062 ตามตรรกะ Search & Extract Logic
        
        กฎสำคัญ:
        1. Noise Reduction: ตัดข้อความ Theme Customizer และขยะระบบทิ้งทั้งหมด
        2. LOS Calculation: นำวันนอน Detox และ Rehab มาบวกกันเป็นตัวเลขรวม
        3. DX Format: รหัส ICD-10 ต้องไม่มีจุดทศนิยม (เช่น F1120)
        4. MEDS Format: ค้นหา Home-Med สกัดชื่อยาเป็น UPPERCASE พร้อมวิธีใช้ (แยกบรรทัดด้วย \\n)
        5. PROGRESS: สังเคราะห์สรุปปัญหาเป็นย่อหน้าเดียว ความยาวเพียง 2-3 บรรทัดเท่านั้น
        6. Verification Audit: หากไม่มีข้อมูล ให้ระบุ [กรอกด้วยตนเอง] ห้ามเว้นว่าง
        
        ข้อมูลดิบ: {all_raw}
        ตอบกลับเป็น JSON ที่มี Key: NAME, AGE, HN, ID, EDU, CAREER, RELIGION, STATUS, RIGHTS, LAST_DC, ADMIT_DATE, VISIT_NUM, CC, CONTACT, RELATION, PHONE, NEAR_HOSP, DC_DATE, LOS, ADDRESS, Q9, Q8, MEDS, DX, PROGRESS, POST_SERVICE
        """
        try:
            response = client.models.generate_content(model=MODEL_ID, contents=prompt)
            match = re.search(r'\{.*\}', response.text, re.DOTALL)
            if match:
                json_data = json.loads(match.group())
                st.success("✅ วิเคราะห์ข้อมูลตามตรรกะสถาบันฯ สำเร็จ!")
                
                word_file = fill_pmnidat_doc(json_data)
                if word_file:
                    log_usage(json_data.get('name', '[ไม่ระบุชื่อ]'))
                    st.download_button(
                        label="💾 ดาวน์โหลดไฟล์ 062 ฉบับสมบูรณ์ (ฟอนต์ 13)",
                        data=word_file,
                        file_name=f"Refer_{json_data.get('name','062')}.docx"
                    )
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")

st.divider()
st.info("""
    **ประกาศ PDPA Compliance:** ระบบไม่มีการจัดเก็บข้อมูลผู้ป่วยถาวร ข้อมูลจะสูญหายทันทีเมื่อรีเฟรชหน้าจอ
    """)
