import streamlit as st
from google import genai 
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re
import json
import requests

# --- 1. ระบบจัดการการเชื่อมต่อ (Stable API Connection) ---
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
                    # แก้ไขจุดที่ 1: เปลี่ยนจาก RIGHT เป็น LEFT ตามคำแนะนำของคุณหมอ
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT 
                    # บังคับฟอนต์ขนาด 13 ตามมาตรฐาน PMNIDAT
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

# --- 4. การออกแบบหน้าเว็บและคู่มือการคัดลอกฉบับเต็ม ---
st.set_page_config(page_title="PMNIDAT 062 Smart Portal", layout="wide")

with st.sidebar:
    st.image("https://pmnidathis.dms.go.th/static/media/health.1bfd961f.png", width=100)
    st.header("📖 คู่มือการคัดลอกข้อมูล")
    st.info("""
    **ขั้นตอนง่ายๆ สำหรับพี่พยาบาล:** วิธีคัดลอกข้อมูลจากระบบ @ThanHIS มาวางตามช่องย่อย
    1. ให้คลิ๊กเมาส์ซ้ายค้างไว้ที่ส่วนต้นของข้อความ แล้วลากเมาส์ลงมาให้ครอบคลุมข้อความทั้งหมด
    2. จากนั้นคลิ๊กขวาเพื่อเลือก copy หรือกด **Ctrl+C** ก็ได้ครับ
    3. มาที่จอหน้านี้ กดที่ช่องย่อยที่จะวาง จากนั้นคลิ๊กขวาเพื่อเลือก paste หรือกด **Ctrl+V** ก็ได้ครับ
    """)
    
    st.markdown("""
    **🟢 STEP 1: ระบบผู้ป่วยใน (IPD)**
    1. **Admission Note:** ข้อมูลแรกรับ
    2. **การวินิจฉัย:** รหัส ICD-10
    3. **Order/Meds:** รายการยา
    4. **Progress Note:** บันทึกความก้าวหน้า
    
    **🔵 STEP 2: การประเมิน**
    * คัดลอกผลคะแนน 9Q, 8Q, BPRS จากส่วน Assessment
    
    **🟠 STEP 3: เวชระเบียน (Registration)**
    * ข้อมูลทั่วไป, ที่อยู่ปัจจุบัน, ผู้ติดต่อ และสิทธิการรักษา
    """)
    st.divider()
    st.success(f"💡 AI เชื่อมต่อสำเร็จผ่านโมเดล: {MODEL_ID}")

st.title("🏥 PMNIDAT Smart D/C Transfer")
st.subheader("ระบบสร้างไฟล์ใบส่งต่อ 062 อัตโนมัติ (Master Version 3.21)")

st.divider()
st.markdown("### **🟢 Step 1: ข้อมูลระบบผู้ป่วยใน (IPD)**")
s1_cols = st.columns(4)
with s1_cols[0]: s11 = st.text_area("1.1 Admission Note", height=150)
with s1_cols[1]: s12 = st.text_area("1.2 การวินิจฉัย", height=150)
with s1_cols[2]: s13 = st.text_area("1.3 Order / Meds", height=150)
with s1_cols[3]: s14 = st.text_area("1.4 Progress Note", height=150)

st.divider()
st.markdown("### **🔵 Step 2: การประเมิน (Assessment)**")
s2 = st.text_area("คัดลอกผลคะแนน 9Q, 8Q, BPRS มาวางที่นี่", height=120)

st.divider()
st.markdown("### **🟠 Step 3: ข้อมูลเวชระเบียน (Registration)**")
s3_cols = st.columns(4)
with s3_cols[0]: s31 = st.text_area("3.1 ทั่วไป 1", height=150)
with s3_cols[1]: s32 = st.text_area("3.2 ที่อยู่ปัจจุบัน", height=150)
with s3_cols[2]: s33 = st.text_area("3.3 ผู้ติดต่อ", height=150)
with s3_cols[3]: s34 = st.text_area("3.4 สิทธิการรักษา", height=150)

# --- 5. ส่วนประมวลผล (ปรับ Prompt เพื่อสรุปสั้น 2-3 บรรทัด) ---
if st.button("🚀 สกัดข้อมูลและสร้างเอกสาร"):
    all_raw = f"{s11} {s12} {s13} {s14} {s2} {s31} {s32} {s33} {s34}"
    with st.spinner('Gemini กำลังสรุปข้อมูลอย่างกระชับ...'):
        prompt = f"""
        จงสกัดข้อมูลเวชระเบียนลงแบบฟอร์ม 062 ตามกฎเหล็ก:
        1. ยา (MEDS): ต้องมีเลขลำดับ และเคาะบรรทัดแยกรายการ (\\n) ชื่อยา UPPERCASE
        2. วินิจฉัย (DX): รหัส ICD-10 ติดกัน(ไม่มีจุด)
        3. PROGRESS (สรุปปัญหา): แก้ไขจุดที่ 2: ให้สังเคราะห์เป็นข้อความย่อหน้าเดียว ความยาวเพียง 2-3 บรรทัด เท่านั้น เน้นประเด็นสำคัญที่สุด
        4. ข้อมูลขาดหาย: ให้ระบุ [พิมพ์ด้วยตนเอง] ห้ามเว้นว่าง
        
        ข้อมูลดิบ: {all_raw}
        ตอบกลับในรูปแบบ JSON เท่านั้น
        """
        try:
            response = client.models.generate_content(model=MODEL_ID, contents=prompt)
            match = re.search(r'\{.*\}', response.text, re.DOTALL)
            if match:
                json_data = json.loads(match.group())
                st.success("✅ สรุปข้อมูลสำเร็จ!")
                
                word_file = fill_pmnidat_doc(json_data)
                if word_file:
                    log_usage(json_data.get('name', '[ไม่ระบุชื่อ]'))
                    st.download_button(
                        label="💾 ดาวน์โหลดไฟล์ 062 (จัดรูปแบบใหม่)",
                        data=word_file,
                        file_name=f"Refer_{json_data.get('name','062')}.docx"
                    )
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")

# --- 6. มาตรการ PDPA ---
st.divider()
st.info("""
    **ประกาศ: มาตรการรักษาความปลอดภัยข้อมูลทางการแพทย์ (PDPA Compliance)**
    * ระบบไม่เก็บข้อมูลผู้ป่วยไว้ในเซิร์ฟเวอร์ ข้อมูลจะสูญหายทันทีหากรีเฟรชหน้าจอ โปรดบันทึกไฟล์ให้เรียบร้อยก่อนออกจากระบบ
    """)
