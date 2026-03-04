import streamlit as st
from google import genai  # เปลี่ยนการ Import เป็น Library ใหม่ปี 2026
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re
import json
import requests

# --- 1. การดึงคีย์และการตั้งค่า Client (2026 Updated Syntax) ---
try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
    APPS_SCRIPT_URL = st.secrets["APPS_SCRIPT_URL"]
    
    # สร้าง Client ด้วย google-genai SDK
    client = genai.Client(api_key=API_KEY)
    
    # ระบุรุ่นโมเดล Gemini 3 Flash ตามที่ปรากฏใน Paid Tier 1 Dashboard
    MODEL_ID = "gemini-3-flash" 
except Exception as e:
    st.error("❌ ระบบตรวจไม่พบรหัสความปลอดภัยใน Secrets หรือการเชื่อมต่อล้มเหลว")
    st.stop()

# --- 2. ฟังก์ชันบันทึก Log Book ---
def log_usage(patient_name):
    try:
        requests.post(APPS_SCRIPT_URL, json={"name": patient_name}, timeout=5)
    except:
        pass

# --- 3. ฟังก์ชันจัดการไฟล์ Word ---
def fill_pmnidat_doc(data):
    try:
        doc = Document("PMNIDAT 062 แบบบันทึกข้อมูลเพื่อส่งต่อ.docx")
        mapping = {f"{{{{{k.upper()}}}}}": str(v) for k, v in data.items()}
        
        def apply_style_and_replace(paragraph):
            for key, value in mapping.items():
                if key in paragraph.text:
                    paragraph.text = paragraph.text.replace(key, value)
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT 
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

# --- 4. การออกแบบหน้าเว็บ ---
st.set_page_config(page_title="PMNIDAT Smart Portal", layout="wide")

with st.sidebar:
    st.image("https://pmnidathis.dms.go.th/static/media/health.1bfd961f.png", width=100)
    st.header("📖 คู่มือการคัดลอกข้อมูล")
    st.info("""
    **ขั้นตอนง่ายๆ สำหรับพี่พยาบาล:** วิธีคัดลอกข้อมูลจากระบบ @ThanHIS มาวางตามช่องย่อย
    1. ลากเมาส์ครอบข้อความทั้งหมด
    2. กด **Ctrl+C** เพื่อคัดลอก
    3. มาที่ช่องย่อยในหน้านี้ กด **Ctrl+V** เพื่อวาง
    """)
    st.divider()
    st.success("💡 AI จะจัดรูปแบบข้อมูลให้สวยงามอัตโนมัติครับ")

st.title("🏥 PMNIDAT Smart D/C Transfer")
st.subheader("ระบบสร้างไฟล์ใบส่งต่อ 062 อัตโนมัติ (Master Version 3.13 - 2026 Engine)")

st.divider()
st.markdown("### **🟢 Step 1: ข้อมูลระบบผู้ป่วยใน (IPD)**")
s1_cols = st.columns(4)
with s1_cols[0]: s11 = st.text_area("1.1 Admission Note", height=150, placeholder="วางข้อมูลแรกรับ...")
with s1_cols[1]: s12 = st.text_area("1.2 การวินิจฉัย", height=150, placeholder="วางรหัส ICD-10...")
with s1_cols[2]: s13 = st.text_area("1.3 Order / Meds", height=150, placeholder="วางรายการยา...")
with s1_cols[3]: s14 = st.text_area("1.4 Progress Note", height=150, placeholder="วาง SOAP...")

st.divider()
st.markdown("### **🔵 Step 2: การประเมิน (Assessment)**")
s2 = st.text_area("คัดลอกผลคะแนน 9Q, 8Q, BPRS มาวางที่นี่", height=120)

st.divider()
st.markdown("### **🟠 Step 3: ข้อมูลเวชระเบียน (Registration)**")
s3_cols = st.columns(4)
with s3_cols[0]: s31 = st.text_area("3.1 ทั่วไป 1", height=150, placeholder="ชื่อ, อายุ, การศึกษา...")
with s3_cols[1]: s32 = st.text_area("3.2 ที่อยู่ปัจจุบัน", height=150, placeholder="ที่อยู่ติดต่อได้จริง...")
with s3_cols[2]: s33 = st.text_area("3.3 ผู้ติดต่อ", height=150, placeholder="ชื่อญาติ และเบอร์โทร...")
with s3_cols[3]: s34 = st.text_area("3.4 สิทธิการรักษา", height=150, placeholder="สิทธิ์รักษา และ รพ.หลัก...")

# --- 5. ส่วนประมวลผล (2026 API Call) ---
if st.button("🚀 สกัดข้อมูลและสร้างเอกสาร"):
    all_raw = f"{s11} {s12} {s13} {s14} {s2} {s31} {s32} {s33} {s34}"
    with st.spinner('Gemini 3 Flash กำลังสังเคราะห์เนื้อหา...'):
        prompt = f"""
        จงสกัดข้อมูลเวชระเบียนลงแบบฟอร์ม 062 ตามกฎเหล็ก:
        1. ยา (MEDS): ต้องมีเลขลำดับ และเคาะบรรทัดแยกรายการ (\\n) ชื่อยาต้อง UPPERCASE พร้อมวิธีใช้ครบถ้วน
        2. วินิจฉัย (DX): เคาะบรรทัดแยกแต่ละโรค รหัส ICD-10 ติดกัน(ไม่มีจุด) + ชื่อโรคภาษาอังกฤษฉบับเต็ม
        3. สรุปปัญหา (PROGRESS): สังเคราะห์เป็น 3 ย่อหน้า (แรกรับ, พัฒนาการดีขึ้น, สถานะปัจจุบันและข้อควรระวัง)
        4. ข้อมูลขาดหาย: ให้ระบุ [พิมพ์ด้วยตนเอง] ห้ามเว้นว่าง
        
        ข้อมูล: {all_raw}
        ตอบเป็น JSON เท่านั้น
        """
        try:
            # ใช้การเรียก Generate Content แบบใหม่ปี 2026
            response = client.models.generate_content(
                model=MODEL_ID,
                contents=prompt
            )
            
            # สกัด JSON จากข้อความตอบกลับ
            match = re.search(r'\{.*\}', response.text, re.DOTALL)
            if match:
                json_data = json.loads(match.group())
                st.success("✅ วิเคราะห์ข้อมูลสำเร็จ!")
                
                word_file = fill_pmnidat_doc(json_data)
                if word_file:
                    log_usage(json_data.get('name', '[ไม่ระบุชื่อ]'))
                    st.download_button(
                        label="💾 ดาวน์โหลดไฟล์ 062 ฉบับสมบูรณ์ (ฟอนต์ 13)",
                        data=word_file,
                        file_name=f"Refer_{json_data.get('name','062')}.docx"
                    )
            else:
                st.error("AI ตอบกลับมาในรูปแบบที่ไม่ถูกต้อง กรุณาลองใหม่อีกครั้ง")
                
        except Exception as e:
            st.error(f"ระบบขัดข้อง: {e}")

st.divider()
st.info("""
    **ประกาศ: มาตรการรักษาความปลอดภัยข้อมูลทางการแพทย์ (PDPA Compliance)**
    * ระบบไม่จัดเก็บข้อมูลผู้ป่วยไว้ในเซิร์ฟเวอร์ ข้อมูลจะสูญหายทันทีหากรีเฟรชหน้าจอ โปรดบันทึกไฟล์ให้เรียบร้อยก่อนออกจากระบบ
    """)
