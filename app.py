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
            # ค้นหาโมเดล Flash ที่รองรับสถานะ Paid Tier 1
            for m in available_models:
                if "gemini-1.5-flash" in m: return m
            return available_models[0]
        except:
            return "models/gemini-1.5-flash"

    MODEL_ID = find_active_model()
except Exception as e:
    st.error(f"❌ ระบบเชื่อมต่อผิดพลาด: {e}")
    st.stop()

# --- 2. ระบบ Session State Memory (จดจำค่าที่กรอกและปุ่มทดสอบระบบ) ---
field_keys = ['s11', 's12', 's13', 's14', 's2', 's31', 's32', 's33', 's34']
for key in field_keys:
    if key not in st.session_state:
        st.session_state[key] = ""

def load_test_data():
    """โหลดข้อมูลเคสตัวอย่างสำหรับทดสอบระบบ"""
    st.session_state.s11 = "น.ส. ณัฐณิชา พ่วงประจง อายุ 29 ปี 8 เดือน สิทธิ์ : 70 Admit Date 16/12/2568 จำนวนวัน Detox [29] วัน Rehab [37] วัน CC : เคตามีน เสพซ้ำ"
    st.session_state.s12 = "1 F162 Mental and behavioural disorders due to use of hallucinogens 2 N302 Other chronic cystitis"
    st.session_state.s13 = "Med 1610061 : Quetiapine 200 mg (QuaPine)(C) 1.5เม็ด*HS จำนวน : 45 Home-Med \nOther Details : rehab and recov สบยช + meth"
    st.session_state.s14 = "S: สบายดี กินข้าวได้ นอนหลับได้ O: V/S stable Euthymic mood A: Stable, No craving P: Discharge"
    st.session_state.s2 = "แบบประเมิน 2Q 9Q 8Q ผลรวมการประเมินโรคซึมเศร้า : 20 คะแนน ผลรวมการประเมินการฆ่าตัวตาย : 15 คะแนน"
    st.session_state.s31 = "ชื่อ น.ส. ณัฐณิชา พ่วงประจง อายุ 29 ปี ศาสนา พุทธ เลขที่บัตรประชาชน 1-1014-00221-30-4"
    st.session_state.s32 = "บ้านเลขที่ 52/603 หมู่ 7 ตรอก/ซอย พหลโยธิน87 ต.หลักหก อ.เมืองปทุมธานี จ.ปทุมธานี 12000"
    st.session_state.s33 = "นางดาราณี เทียนประทุม (มารดา) เบอร์โทรศัพท์ 083-024-XXXX"
    st.session_state.s34 = "สิทธิหลัก [70] หลักประกันสุขภาพแห่งชาติ สถานพยาบาลหลัก [13691] ศูนย์บริการสาธารณสุข - 48"

def clear_all_data():
    for key in field_keys:
        st.session_state[key] = ""

# --- 3. ฟังก์ชันจัดการไฟล์ Word (ชิดซ้าย + ฟอนต์ 13 + สรุปกระชับ) ---
def fill_pmnidat_doc(data):
    try:
        doc = Document("PMNIDAT 062 แบบบันทึกข้อมูลเพื่อส่งต่อ.docx")
        mapping = {f"{{{{{k.upper()}}}}}": str(v) for k, v in data.items()}
        def apply_style_and_replace(paragraph):
            for key, value in mapping.items():
                if key in paragraph.text:
                    paragraph.text = paragraph.text.replace(key, value)
                    # จัดวางแบบชิดซ้ายตามคำแนะนำของคุณหมออาร์ม
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

# --- 4. การออกแบบหน้าเว็บและ "คู่มือพี่พยาบาล ฉบับจับมือทำ" (UI/UX) ---
st.set_page_config(page_title="PMNIDAT 062 Smart Portal", layout="wide")

with st.sidebar:
    st.image("https://pmnidathis.dms.go.th/static/media/health.1bfd961f.png", width=100)
    st.header("📖 คู่มือพี่พยาบาล ฉบับจับมือทำ")
    
    st.subheader("🧪 เครื่องมือทดสอบระบบ")
    c1, c2 = st.columns(2)
    with c1: st.button("🧬 ลงข้อมูลตัวอย่าง", on_click=load_test_data, use_container_width=True)
    with c2: st.button("🧹 ล้างข้อมูล", on_click=clear_all_data, use_container_width=True)
    
    st.divider()
    
    # --- ส่วนคู่มือแบบละเอียดที่คุณหมอต้องการ ---
    st.markdown("### **วิธีการคัดลอกข้อมูลจาก @ThanHIS**")
    st.markdown("""
    1. **คลิกเมาส์ซ้ายค้าง** ที่ต้นข้อความ ลากครอบให้คลุมทั้งหมด
    2. กด **Ctrl+C** (คัดลอก)
    3. มาที่หน้านี้ คลิกช่องที่ต้องการ แล้วกด **Ctrl+V** (วาง)
    """)
    
    with st.expander("🟢 STEP 1: ระบบผู้ป่วยใน (IPD)", expanded=True):
        st.markdown("""
        1.1 **Admission Note:** ดูข้อมูลคนไข้ → Admission note → คัดลอกทั้งหมด
        1.2 **การวินิจฉัย:** กดเมนู "การวินิจฉัย" → คัดลอกรหัส ICD-10
        1.3 **Order / Meds:** กดเมนู "Order" → คัดลอก Discharge order + Home medication
        1.4 **Progress Note:** กดเมนู "Progress note" → คัดลอกบันทึกล่าสุด (SOAP)
        """)

    with st.expander("🔵 STEP 2: การประเมิน (Assessment)"):
        st.markdown("""
        - กดเมนู "Admission note" → ปุ่ม "ข้อมูลผู้ป่วยนอก"
        - เลื่อนลงล่างไปที่หัวข้อ **Assessment** - คัดลอกผลคะแนน 9Q, 8Q, BPRS
        """)

    with st.expander("🟠 STEP 3: เวชระเบียน (Registration)"):
        st.markdown("""
        - ระบบผู้ป่วยนอก → เวชระเบียน → ลงทะเบียนผู้ป่วย → ค้นหา HN
        - 3.1 **ทั่วไป 1:** คัดลอกข้อมูล ชื่อ, อายุ, เลขบัตรประชาชน, ศาสนา
        - 3.2 **ทั่วไป 2:** กด "ที่อยู่ปัจจุบัน" → คัดลอกที่อยู่ทั้งหมด
        - 3.3 **ผู้ติดต่อ:** คัดลอกชื่อญาติและเบอร์โทรศัพท์
        - 3.4 **สิทธิรักษา:** คัดลอกสิทธิ์และ "สถานพยาบาลหลัก"
        """)
    
    st.divider()
    st.success(f"💡 AI วิเคราะห์ด้วยตรรกะ PhD: {MODEL_ID}")

st.title("🏥 PMNIDAT Smart D/C Transfer")
st.subheader("ระบบสร้างไฟล์ใบส่งต่อ 062 อัตโนมัติ (Master Version 3.28)")

st.divider()
# --- 5. ส่วนกรอกข้อมูล 9 ช่อง (เชื่อมโยง Session State เพื่อความสวยงามและความสะดวก) ---
st.markdown("### **🟢 Step 1: ระบบผู้ป่วยใน (IPD)**")
s1_cols = st.columns(4)
s11 = s1_cols[0].text_area("1.1 Admission Note", value=st.session_state.s11, height=150, key="input_s11")
s12 = s1_cols[1].text_area("1.2 การวินิจฉัย", value=st.session_state.s12, height=150, key="input_s12")
s13 = s1_cols[2].text_area("1.3 Order / Meds", value=st.session_state.s13, height=150, key="input_s13")
s14 = s1_cols[3].text_area("1.4 Progress Note", value=st.session_state.s14, height=150, key="input_s14")
# อัปเดต Session State
st.session_state.s11, st.session_state.s12, st.session_state.s13, st.session_state.s14 = s11, s12, s13, s14

st.divider()
s2 = st.text_area("🔵 Step 2: คะแนนการประเมิน (9Q, 8Q, BPRS)", value=st.session_state.s2, height=100, key="input_s2")
st.session_state.s2 = s2

st.divider()
st.markdown("### **🟠 Step 3: เวชระเบียน (Registration)**")
s3_cols = st.columns(4)
s31 = s3_cols[0].text_area("3.1 ข้อมูลทั่วไป", value=st.session_state.s31, height=150, key="input_s31")
s32 = s3_cols[1].text_area("3.2 ที่อยู่ปัจจุบัน", value=st.session_state.s32, height=150, key="input_s32")
s33 = s3_cols[2].text_area("3.3 ผู้ติดต่อ", value=st.session_state.s33, height=150, key="input_s33")
s34 = s3_cols[3].text_area("3.4 สิทธิการรักษา", value=st.session_state.s34, height=150, key="input_s34")
st.session_state.s31, st.session_state.s32, st.session_state.s33, st.session_state.s34 = s31, s32, s33, s34

# --- 6. ส่วนประมวลผล (Advanced Synthesis & Extraction Logic) ---
if st.button("🚀 สกัดข้อมูลและสร้างเอกสาร"):
    all_raw = f"{s11}\n{s12}\n{s13}\n{s14}\n{s2}\n{s31}\n{s32}\n{s33}\n{s34}"
    with st.spinner('กำลังวิเคราะห์ข้อมูลและคัดกรองขยะ (Noise Reduction)...'):
        prompt = f"""
        จงสกัดข้อมูลเวชระเบียนลงแบบฟอร์ม 062 ตามตรรกะ Search & Extract Logic:
        1. Noise Reduction: ตัดข้อความ Theme Customizer และขยะระบบทิ้งทั้งหมด
        2. LOS Calc: นำวันนอน Detox และ Rehab มาบวกกันเป็นตัวเลขเดียว
        3. DX Format: รหัส ICD-10 เขียนติดกันโดยไม่มีจุดทศนิยม
        4. MEDS: คัดเฉพาะ Home-Med ชื่อยา UPPERCASE พร้อมวิธีใช้ (\\n แยกบรรทัด)
        5. PROGRESS: สังเคราะห์สรุปปัญหาเป็นย่อหน้าเดียว ความยาวเพียง 2-3 บรรทัดเท่านั้น
        6. Verification: หากไม่มีข้อมูล ให้ระบุ [กรอกด้วยตนเอง] ห้ามเว้นว่าง
        
        ข้อมูลดิบ: {all_raw}
        ตอบกลับในรูปแบบ JSON ที่มี Key ตรงกับ Placeholder ในไฟล์ Word เท่านั้น
        """
        try:
            response = client.models.generate_content(model=MODEL_ID, contents=prompt)
            match = re.search(r'\{.*\}', response.text, re.DOTALL)
            if match:
                json_data = json.loads(match.group())
                st.success("✅ วิเคราะห์ข้อมูลสำเร็จ!")
                
                word_file = fill_pmnidat_doc(json_data)
                if word_file:
                    st.download_button(
                        label="💾 ดาวน์โหลดไฟล์ 062 ฉบับสมบูรณ์ (ฟอนต์ 13)",
                        data=word_file,
                        file_name=f"Refer_{json_data.get('name','062')}.docx"
                    )
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาด: {e}")

st.divider()
st.info("**มาตรการ PDPA:** ระบบไม่มีการจัดเก็บข้อมูลผู้ป่วยถาวร ข้อมูลจะหายไปเมื่อรีเฟรชหน้าจอ")
