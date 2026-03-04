import streamlit as st
from google import genai  # ใช้ Library มาตรฐานใหม่ปี 2026
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re
import json
import requests

# --- 1. การดึงคีย์และการตั้งค่า Client (2026 Stable Syntax) ---
try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
    APPS_SCRIPT_URL = st.secrets["APPS_SCRIPT_URL"]
    
    # สร้าง Client ด้วย google-genai SDK
    client = genai.Client(api_key=API_KEY)
    
    # ใช้รุ่น gemini-1.5-flash เพื่อความเสถียรสูงสุดและป้องกัน Error 404
    MODEL_ID = "gemini-1.5-flash" 
except Exception as e:
    st.error("❌ ระบบตรวจไม่พบรหัสความปลอดภัยใน Secrets กรุณาตรวจสอบการตั้งค่า")
    st.stop()

# --- 2. ฟังก์ชันบันทึก Log Book ---
def log_usage(patient_name):
    try:
        # ส่งข้อมูลไปยัง Google Sheets ผ่าน Apps Script URL
        requests.post(APPS_SCRIPT_URL, json={"name": patient_name}, timeout=5)
    except:
        pass # ป้องกันแอปหยุดทำงานหากบันทึก Log ไม่สำเร็จ

# --- 3. ฟังก์ชันจัดการไฟล์ Word (ชิดขวา + ฟอนต์ 13 + No Stretching) ---
def fill_pmnidat_doc(data):
    try:
        doc = Document("PMNIDAT 062 แบบบันทึกข้อมูลเพื่อส่งต่อ.docx")
        mapping = {f"{{{{{k.upper()}}}}}": str(v) for k, v in data.items()}
        
        def apply_style_and_replace(paragraph):
            for key, value in mapping.items():
                if key in paragraph.text:
                    # แทนที่ข้อความ (รองรับ \n สำหรับการเคาะบรรทัดรายการยาและโรค)
                    paragraph.text = paragraph.text.replace(key, value)
                    
                    # บังคับชิดขวา เพื่อยกเลิกการยืดข้อความ (Thai Distributed) ให้ตัวอักษรเรียงตัวสวยงาม
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT 
                    
                    # บังคับฟอนต์ขนาด 13 ตามมาตรฐานสถาบันฯ
                    for run in paragraph.runs:
                        run.font.size = Pt(13)

        # ตรวจสอบทั้งเนื้อหาปกติและในตาราง
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

# --- 4. การออกแบบหน้าเว็บ (คืนค่าคู่มือแบบละเอียดสำหรับพี่พยาบาล) ---
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
    1. **Admission Note:** กดดูข้อมูลคนไข้ → Admission note → คัดลอกทั้งหมด
    2. **การวินิจฉัย:** กด "การวินิจฉัย" → คัดลอกทั้งหมด
    3. **Order/Meds:** หน้า "Order" → คัดลอกทั้งหมด
    4. **Progress Note:** หน้า "Progress note" → คัดลอกทั้งหมด
    
    **🔵 STEP 2: การประเมิน**
    * กด "Admission note" → กดปุ่ม ข้อมูลผู้ป่วยนอก → เลื่อนลงล่างไปตรง **Assessment** → คัดลอก 9Q, 8Q, BPRS
    
    **🟠 STEP 3: เวชระเบียน (Registration)**
    * ระบบผู้ป่วยนอก → เวชระเบียน → ลงทะเบียนผู้ป่วย → กดปุ่ม ค้นหา → กรอก HN → กดดูข้อมูลคนไข้
    * กดหน้า **ทั่วไป 1** →  คัดลอกทั้งหมด
    * กดหน้า **ทั่วไป 2** →  กดแสดง "ทั่วไป 2.2 (ที่อยู่ปัจจุบัน)" → คัดลอกทั้งหมด
    * กดหน้า **ผู้ติดต่อ** → คัดลอกทั้งหมด
    * กดหน้า **สิทธิการรักษา** → คัดลอกทั้งหมด
    """)
    st.divider()
    st.success("💡 AI จะจัดรูปแบบข้อมูลให้สวยงามแบบแบบฟอร์มอัตโนมัติครับ สงสัยหรือพบปัญหาติดต่อ พ.ชาฌานได้เลยนะครับ")

st.title("🏥 PMNIDAT Smart D/C Transfer")
st.subheader("ระบบสร้างไฟล์ใบส่งต่อ 062 อัตโนมัติ (Master Version 3.13)")

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

# --- 5. ส่วนประมวลผล (2026 Updated API Call) ---
if st.button("🚀 สกัดข้อมูลและสร้างเอกสาร"):
    all_raw = f"{s11} {s12} {s13} {s14} {s2} {s31} {s32} {s33} {s34}"
    with st.spinner('AI กำลังสังเคราะห์เนื้อหาและจัดรูปแบบตามมาตรฐานสถาบัน...'):
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
            # เรียกใช้โมเดลผ่าน Syntax ใหม่
            response = client.models.generate_content(
                model=MODEL_ID,
                contents=prompt
            )
            
            # สกัดเฉพาะส่วนที่เป็น JSON
            match = re.search(r'\{.*\}', response.text, re.DOTALL)
            if match:
                json_data = json.loads(match.group())
                st.success("✅ วิเคราะห์ข้อมูลสำเร็จ!")
                
                word_file = fill_pmnidat_doc(json_data)
                if word_file:
                    # บันทึกสถิติลง Google Sheet (เฉพาะชื่อ)
                    log_usage(json_data.get('name', '[ไม่ระบุชื่อ]'))
                    
                    st.download_button(
                        label="💾 ดาวน์โหลดไฟล์ 062 ฉบับสมบูรณ์ (ฟอนต์ 13)",
                        data=word_file,
                        file_name=f"Refer_{json_data.get('name','062')}.docx"
                    )
            else:
                st.error("AI ไม่สามารถจัดรูปแบบข้อมูลได้ กรุณาลองใหม่อีกครั้ง")
        except Exception as e:
            st.error(f"ระบบขัดข้อง: {e}")

# --- 6. ประกาศมาตรการรักษาความปลอดภัย (PDPA) ---
st.divider()
st.info("""
    **ประกาศ: มาตรการรักษาความปลอดภัยข้อมูลทางการแพทย์ (PDPA Compliance)**
    * ระบบ PMNIDAT Smart D/C Transfer ไม่มีการจัดเก็บหรือสำรองข้อมูลผู้ป่วยไว้ในเซิร์ฟเวอร์หรือฐานข้อมูลใดๆ เพื่อป้องกันการรั่วไหลของข้อมูลส่วนบุคคล 
    * ข้อมูลจะปรากฏเฉพาะในระหว่างการใช้งานของท่านเท่านั้น หากมีการรีเฟรชหน้าจอ ข้อมูลจะสูญหายทันที โปรดบันทึกไฟล์เอกสาร (.docx) ให้เรียบร้อยก่อนออกจากระบบทุกครั้ง
    """)
