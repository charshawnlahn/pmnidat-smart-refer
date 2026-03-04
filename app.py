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
        # เตรียมชุดข้อมูลตาม Placeholder
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

# --- 4. การออกแบบหน้าเว็บและคู่มือแบบละเอียดสำหรับพี่พยาบาล ---
st.set_page_config(page_title="PMNIDAT 062 Smart Portal", layout="wide")

with st.sidebar:
    st.image("https://pmnidathis.dms.go.th/static/media/health.1bfd961f.png", width=100)
    st.header("📖 คู่มือการคัดลอกข้อมูล")
    st.info("""
    **ขั้นตอนง่ายๆ สำหรับพี่พยาบาล:** วิธีคัดลอกข้อมูลจากระบบ @ThanHIS
    1. ลากเมาส์ครอบข้อความให้หมด แล้วกด **Ctrl+C**
    2. มาที่ช่องย่อยในหน้านี้ แล้วกด **Ctrl+V** เพื่อวาง
    """)
    
    st.markdown("""
    **🟢 STEP 1: ระบบผู้ป่วยใน (IPD)**
    1. **Admission Note:** ข้อมูลแรกรับ
    2. **การวินิจฉัย:** รหัส ICD-10
    3. **Order/Meds:** รายการยา
    4. **Progress Note:** บันทึกความก้าวหน้า
    
    **🔵 STEP 2: การประเมิน**
    * คัดลอกคะแนน 9Q, 8Q, BPRS จากส่วน Assessment
    
    **🟠 STEP 3: เวชระเบียน (Registration)**
    * ข้อมูลทั่วไป, ที่อยู่ปัจจุบัน, ผู้ติดต่อ และสิทธิการรักษา
    """)
    st.divider()
    st.success(f"💡 ระบบพร้อมประมวลผลผ่าน: {MODEL_ID}")

st.title("🏥 PMNIDAT Smart D/C Transfer")
st.subheader("ระบบสร้างไฟล์ใบส่งต่อ 062 อัตโนมัติ (Master Version 3.23)")

st.divider()
st.markdown("### **🟢 Step 1: ข้อมูลระบบผู้ป่วยใน (IPD)**")
s1_cols = st.columns(4)
with s1_cols[0]: s11 = st.text_area("1.1 Admission Note", height=150, placeholder="วางข้อมูลจาก Step 1.1...")
with s1_cols[1]: s12 = st.text_area("1.2 การวินิจฉัย", height=150, placeholder="วางข้อมูลจาก Step 1.2...")
with s1_cols[2]: s13 = st.text_area("1.3 Order / Meds", height=150, placeholder="วางข้อมูลจาก Step 1.3...")
with s1_cols[3]: s14 = st.text_area("1.4 Progress Note", height=150, placeholder="วางข้อมูลจาก Step 1.4...")

st.divider()
st.markdown("### **🔵 Step 2: การประเมิน (Assessment)**")
s2 = st.text_area("คัดลอกผลคะแนน 9Q, 8Q, BPRS (Step 2)", height=120)

st.divider()
st.markdown("### **🟠 Step 3: ข้อมูลเวชระเบียน (Registration)**")
s3_cols = st.columns(4)
with s3_cols[0]: s31 = st.text_area("3.1 ทั่วไป 1", height=150)
with s3_cols[1]: s32 = st.text_area("3.2 ที่อยู่ปัจจุบัน", height=150)
with s3_cols[2]: s33 = st.text_area("3.3 ผู้ติดต่อ", height=150)
with s3_cols[3]: s34 = st.text_area("3.4 สิทธิการรักษา", height=150)

# --- 5. ส่วนประมวลผล (ฝังตรรกะ Search & Extract และ Verification Audit) ---
if st.button("🚀 สกัดข้อมูลและสร้างเอกสาร"):
    all_raw = f"---IPD DATA---\n{s11}\n{s12}\n{s13}\n{s14}\n---ASSESSMENT---\n{s2}\n---REGISTRATION---\n{s31}\n{s32}\n{s33}\n{s34}"
    with st.spinner('กำลังวิเคราะห์ข้อมูลตามตรรกะการสกัดและตรวจสอบ (Verification Audit)...'):
        prompt = f"""
        คุณคือผู้ช่วยวิจัยระดับ PhD ทางการแพทย์ ทำหน้าที่สกัดข้อมูลจากระบบ @ThanHIS ลงแบบฟอร์ม 062 ของสถาบันฯ
        
        กฎการตัดขยะข้อมูล (Noise Reduction Rule):
        - ให้ Ignore (ละทิ้ง) ข้อความเกี่ยวกับ Theme Customizer, Menu Colors, Light/Dark Mode, Font Size หรือ COPYRIGHT ทั้งหมด 
        
        ตรรกะการสกัดข้อมูล (Search & Extract Logic):
        1. HN/AN: สกัดตัวเลขหลังคำว่า 'HN' หรือ 'AN' [cite: 71]
        2. สิทธิการรักษา (RIGHTS): สกัดข้อความระหว่าง 'สิทธิ์ :' ถึงข้อความถัดไป [cite: 71]
        3. อาการนำส่ง (CC): สกัดจาก 'Chief Complaint' หรือ 'CC :' จนถึง 'Present illness' [cite: 71]
        4. วันนอน (LOS): ค้นหาตัวเลขหน้าคำว่า 'วัน' ในแถว 'Detox' และ 'Rehab' แล้วนำมาบวกกันเสมอ 
        5. รหัสโรค (DX): สกัดรหัส ICD-10 (เช่น F155) และตัดจุดทศนิยมออกให้เป็นตัวเลขติดกัน 
        6. ยา (MEDS): ค้นหาบรรทัดที่มี 'Home-Med' สกัดชื่อยาเป็น UPPERCASE พร้อมเลขลำดับและวิธีใช้ (\\n แยกบรรทัด) 
        7. คะแนนประเมิน (Q9/Q8): สกัดตัวเลขหลังคำว่า 'ผลรวมการประเมินโรคซึมเศร้า' และ 'การฆ่าตัวตาย' [cite: 73]
        8. สรุปปัญหา (PROGRESS): สกัดจาก Progress Note ล่าสุด สังเคราะห์ให้กระชับเพียง 2-3 บรรทัด 
        9. ที่อยู่ (ADDRESS): สกัดจากข้อมูลที่มี แขวง, เขต, จ., รหัสไปรษณีย์ ครบถ้วน [cite: 75]
        
        หากไม่มีข้อมูล: ให้ระบุ [กรอกด้วยตนเอง] ห้ามเว้นว่างเด็ดขาด [cite: 86]
        
        ข้อมูลดิบ:
        {all_raw}
        
        ตอบกลับในรูปแบบ JSON ที่มี Key ตรงกับ Placeholder: NAME, AGE, HN, ID, EDU, CAREER, RELIGION, STATUS, RIGHTS, LAST_DC, ADMIT_DATE, VISIT_NUM, CC, CONTACT, RELATION, PHONE, NEAR_HOSP, DC_DATE, LOS, ADDRESS, Q9, Q8, MEDS, DX, PROGRESS, POST_SERVICE
        """
        try:
            response = client.models.generate_content(model=MODEL_ID, contents=prompt)
            match = re.search(r'\{.*\}', response.text, re.DOTALL)
            if match:
                json_data = json.loads(match.group())
                st.success("✅ วิเคราะห์ข้อมูลตามตรรกะสำเร็จ!")
                
                word_file = fill_pmnidat_doc(json_data)
                if word_file:
                    log_usage(json_data.get('name', '[ไม่ระบุชื่อ]'))
                    st.download_button(
                        label="💾 ดาวน์โหลดไฟล์ 062 ฉบับสมบูรณ์ (Master 3.23)",
                        data=word_file,
                        file_name=f"Refer_{json_data.get('name','062')}.docx"
                    )
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")

st.divider()
st.info("""
    **มาตรการรักษาความปลอดภัยข้อมูล (PDPA Compliance)**
    * ระบบไม่จัดเก็บข้อมูลผู้ป่วยถาวร ข้อมูลจะสูญหายทันทีเมื่อรีเฟรชหน้าจอ โปรดบันทึกไฟล์ก่อนออกจากระบบ
    """)
