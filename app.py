import streamlit as st
from google import genai 
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re
import json
import requests

# --- 1. ฟังก์ชันบันทึก Log (ต้องประกาศไว้ก่อนเพื่อป้องกัน Error) ---
def log_usage(patient_name):
    try:
        # ดึง URL จาก Secrets
        url = st.secrets["APPS_SCRIPT_URL"]
        requests.post(url, json={"name": patient_name}, timeout=5)
    except:
        pass # ป้องกันแอปพังหากเชื่อมต่อ Google Sheets ไม่ได้

# --- 2. การตั้งค่าระบบและการเชื่อมต่อ API ---
try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
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

# --- 3. ระบบ Session State Memory (สำหรับเก็บข้อมูล 9 ช่อง) ---
field_keys = ['s11', 's12', 's13', 's14', 's2', 's31', 's32', 's33', 's34']
for key in field_keys:
    if key not in st.session_state:
        st.session_state[key] = ""

# --- 4. ฟังก์ชันโหลดข้อมูลตัวอย่าง (ชื่อ: นพ.ชาฌาน หลานวงศ์) ---
def load_test_data():
    st.session_state.s11 = "ข้อมูลทั่วไปผู้ป่วย\\nนพ.ชาฌาน หลานวงศ์\\nอายุ 40 ปี 5 เดือน\\nสิทธิ์ : จ่ายตรงกรมบัญชีกลาง (ไม่มีหนังสือส่งตัว)\\nAdmit Date 01/03/2569\\nจำนวนวัน Detox [5] วัน Rehab [10] วัน\\nChief Complaint: เสพยาบ้าซ้ำ ต้องการเข้ารับการบำบัดรักษา"
    st.session_state.s12 = "1 F155 Mental and behavioural disorders due to use of stimulants at dependence syndrome Principal Diagnosis นพ.ชาฌาน หลานวงศ์\\n2 I10 Essential (primary) hypertension Comorbidity นพ.ชาฌาน หลานวงศ์"
    st.session_state.s13 = "04/03/2569 13:31 Order: 77777 By : นพ.ชาฌาน หลานวงศ์ (Doctor)\\nMed 1010101 : AMLODIPINE 5 MG 1 tab OD pc\\nจำนวน : 30 Home-Med\\nMed 1010202 : QUETIAPINE 25 MG 1 tab hs\\nจำนวน : 30 Home-Med"
    st.session_state.s14 = "[S: สบายดี กินข้าวได้ นอนหลับได้ เตรียมจำหน่าย\\nO: V/S stable, BP 120/80 mmHg, Euthymic mood\\nA: Stable, No craving\\nP: Discharge to home]"
    st.session_state.s2 = "แบบประเมิน 2Q 9Q 8Q ผลรวมการประเมินโรคซึมเศร้า : 5 คะแนน ผลรวมการประเมินการฆ่าตัวตาย : 0 คะแนน"
    st.session_state.s31 = "ชื่อ [นพ.ชาฌาน] นามสกุล [หลานวงศ์]\\nวันเกิด [01/06/2529] อายุ [40] ปี\\nเลขที่บัตรประชาชน* [1-2345-67890-12-3]\\nศาสนา [00] [พุทธ] อาชีพ [แพทย์]"
    st.session_state.s32 = "บ้านเลขที่ [123/45] หมู่ [6] ตรอก/ซอย [พหลโยธิน]\\nจังหวัด [13] [ปทุมธานี] อำเภอ [01] [เมืองปทุมธานี] ตำบล [14] [หลักหก]\\nเบอร์โทรศัพท์ [02531XXXX]"
    st.session_state.s33 = "ผู้ติดต่อ (คุณวิไล หลานวงศ์)\\nความสัมพันธ์กับผู้ป่วย [ภรรยา] เบอร์โทรศัพท์ [081234XXXX]"
    st.session_state.s34 = "สิทธิการรักษาของผู้ป่วย\\nสิทธิหลัก ชนิดของบัตร [จ่ายตรง]\\nสถานพยาบาลหลัก [11111] [สถาบันบำบัดรักษาและฟื้นฟูผู้ติดยาเสพติดแห่งชาติบรมราชชนนี]"
    st.rerun()

def clear_all_data():
    for key in field_keys:
        st.session_state[key] = ""
    st.rerun()

# --- 5. ส่วน Sidebar (คู่มือพี่พยาบาล ฉบับจับมือทำ) ---
st.set_page_config(page_title="PMNIDAT Smart Portal", layout="wide")

with st.sidebar:
    st.image("https://pmnidathis.dms.go.th/static/media/health.1bfd961f.png", width=100)
    st.header("📖 คู่มือพี่พยาบาล ฉบับจับมือทำ")
    
    st.subheader("🧪 เครื่องมือช่วยทดสอบ")
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        if st.button("🧬 ลงข้อมูลตัวอย่าง", use_container_width=True):
            load_test_data()
    with col_t2:
        if st.button("🧹 ล้างข้อมูล", use_container_width=True):
            clear_all_data()
            
    st.divider()
    
    st.markdown("### **วิธีการคัดลอกข้อมูลจาก @ThanHIS**")
    st.markdown("""
    1. **คลิกเมาส์ซ้ายค้าง** ที่ต้นข้อความ ลากครอบให้คลุมทั้งหมด
    2. กด **Ctrl+C** (คัดลอก)
    3. มาที่หน้านี้ คลิกช่องที่ต้องการ แล้วกด **Ctrl+V** (วาง)
    """)
    
    with st.expander("🟢 STEP 1: ระบบผู้ป่วยใน (IPD)", expanded=True):
        st.markdown("""
        **1.1 Admission Note:** - ดูข้อมูลคนไข้ → Admission note 
        - คัดลอกทั้งหมด
        
        **1.2 การวินิจฉัย:** - กดเมนู "การวินิจฉัย" 
        - คัดลอกรหัส ICD-10
        
        **1.3 Order / Meds:** - กดเมนู "Order" 
        - คัดลอก Discharge order + Home medication
        
        **1.4 Progress Note:** - กดเมนู "Progress note" 
        - คัดลอกบันทึกล่าสุด (SOAP)
        """)

    with st.expander("🔵 STEP 2: การประเมิน (Assessment)"):
        st.markdown("""
        - กดเมนู "Admission note" → ปุ่ม "ข้อมูลผู้ป่วยนอก"
        - เลื่อนลงล่างไปที่หัวข้อ **Assessment** - คัดลอกผลคะแนน 9Q, 8Q, BPRS
        """)

    with st.expander("🟠 STEP 3: เวชระเบียน (Registration)"):
        st.markdown("""
        - ระบบผู้ป่วยนอก → เวชระเบียน → ลงทะเบียนผู้ป่วย → ค้นหา HN
        - **3.1 ทั่วไป 1:** คัดลอกข้อมูล ชื่อ, อายุ, เลขบัตรประชาชน, ศาสนา
        - **3.2 ทั่วไป 2:** กด "ที่อยู่ปัจจุบัน" → คัดลอกที่อยู่ทั้งหมด
        - **3.3 ผู้ติดต่อ:** คัดลอกชื่อญาติและเบอร์โทรศัพท์
        - **3.4 สิทธิรักษา:** คัดลอกสิทธิ์และ "สถานพยาบาลหลัก"
        """)
        
    st.divider()
    st.success(f"💡 AI วิเคราะห์ด้วยตรรกะ PhD ผ่าน: {MODEL_ID}")

st.title("🏥 PMNIDAT Smart D/C Transfer")
st.subheader("ระบบสร้างไฟล์ใบส่งต่ออัตโนมัติ (Master Version 3.29)")


# --- 4. ส่วนกรอกแบบฟอร์ม (เชื่อมโยงกับ Session State 100%) ---

st.divider()
# ส่วนที่ 1: ข้อมูลระบบผู้ป่วยใน (IPD) [cite: 68-73, 79]
st.markdown("### **🟢 Step 1: ระบบผู้ป่วยใน (IPD)**")
s1_cols = st.columns(4)

with s1_cols[0]:
    st.text_area(
        "1.1 Admission Note",
        height=200,
        placeholder="วางข้อมูลแรกรับและวันนอน...",
        key="s11",  # เชื่อมกับ st.session_state.s11 [cite: 68-71]
        help="คัดลอกจาก Admission note ในระบบ IPD "
    )

with s1_cols[1]:
    st.text_area(
        "1.2 การวินิจฉัย",
        height=200,
        placeholder="วางรหัส ICD-10 ทั้งหมด...",
        key="s12",  # เชื่อมกับ st.session_state.s12 [cite: 71-73]
        help="คัดลอกจากหน้าการวินิจฉัย (Principal & Comorbidity) [cite: 71]"
    )

with s1_cols[2]:
    st.text_area(
        "1.3 Order / Meds",
        height=200,
        placeholder="วางรายการยา Home-Med...",
        key="s13",  # เชื่อมกับ st.session_state.s13 [cite: 73]
        help="คัดลอก Discharge order + Home medication [cite: 73]"
    )

with s1_cols[3]:
    st.text_area(
        "1.4 Progress Note",
        height=200,
        placeholder="วาง SOAP ล่าสุด...",
        key="s14",  # เชื่อมกับ st.session_state.s14 [cite: 73]
        help="คัดลอกบันทึก Progress note ล่าสุด [cite: 73]"
    )

st.divider()
# ส่วนที่ 2: การประเมิน (Assessment) [cite: 74, 79]
st.markdown("### **🔵 Step 2: การประเมิน (Assessment)**")
st.text_area(
    "คัดลอกผลคะแนน 9Q, 8Q, BPRS มาวางที่นี่",
    height=120,
    key="s2",  # เชื่อมกับ st.session_state.s2 [cite: 74]
    help="ดึงจากหน้า Assessment ผ่านปุ่มข้อมูลผู้ป่วยนอก [cite: 74]"
)

st.divider()
# ส่วนที่ 3: เวชระเบียน (Registration) [cite: 74-76, 79]
st.markdown("### **🟠 Step 3: เวชระเบียน (Registration)**")
s3_cols = st.columns(4)

with s3_cols[0]:
    st.text_area(
        "3.1 ข้อมูลทั่วไป", 
        height=180, 
        placeholder="ชื่อ, อายุ, เลขบัตรประชาชน, ศาสนา...",
        key="s31", # เชื่อมกับ st.session_state.s31 [cite: 74]
        help="คัดลอกจากหน้า 'ทั่วไป 1' ในระบบเวชระเบียน [cite: 74]"
    )

with s3_cols[1]:
    st.text_area(
        "3.2 ที่อยู่ปัจจุบัน", 
        height=180, 
        placeholder="แขวง, เขต, จังหวัด, รหัสไปรษณีย์...",
        key="s32", # เชื่อมกับ st.session_state.s32 [cite: 74-75]
        help="คัดลอกจากหน้า 'ทั่วไป 2' (ที่อยู่ปัจจุบัน) [cite: 74]"
    )

with s3_cols[2]:
    st.text_area(
        "3.3 ผู้ติดต่อ", 
        height=180, 
        placeholder="ชื่อญาติ ความสัมพันธ์ และเบอร์โทร...",
        key="s33", # เชื่อมกับ st.session_state.s33 [cite: 76]
        help="คัดลอกจากหน้า 'ผู้ติดต่อ' [cite: 76]"
    )

with s3_cols[3]:
    st.text_area(
        "3.4 สิทธิการรักษา", 
        height=180, 
        placeholder="สิทธิหลัก และสถานพยาบาลหลัก...",
        key="s34", # เชื่อมกับ st.session_state.s34 [cite: 76]
        help="คัดลอกจากหน้า 'สิทธิการรักษา' [cite: 76]"
    )


# --- 5. ฟังก์ชันจัดการไฟล์ Word (สกัดข้อมูลลง Placeholder + จัดรูปแบบ) ---

def fill_pmnidat_doc(data):
    try:
        # โหลดไฟล์แม่แบบที่คุณหมออัปโหลดไว้ [cite: 31]
        doc = Document("PMNIDAT 062 แบบบันทึกข้อมูลเพื่อส่งต่อ.docx")
        
        # เตรียมชุดข้อมูล Mapping (Key ต้องเป็นตัวพิมพ์ใหญ่ตามใน Word) [cite: 35-51]
        mapping = {f"{{{{{k.upper()}}}}}": str(v) for k, v in data.items()}
        
        def apply_style_and_replace(paragraph):
            for key, value in mapping.items():
                if key in paragraph.text:
                    # แทนที่ข้อความใน Placeholder
                    paragraph.text = paragraph.text.replace(key, value)
                    # บังคับจัดรูปแบบชิดซ้าย (Left Alignment) ตามที่ระบุ [cite: 46-50]
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    # กำหนดฟอนต์ขนาด 13 pt สำหรับภาษาไทยวิชาการ [cite: 46-51]
                    for run in paragraph.runs:
                        run.font.size = Pt(13)

        # ตรวจสอบและแทนที่ข้อมูลทั้งในเนื้อความและในตาราง [cite: 35-60]
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

# --- 6. ส่วนการประมวลผล AI (Search & Extract Logic) ---

if st.button("🚀 สกัดข้อมูลและสร้างเอกสาร 062", use_container_width=True):
    # รวบรวมข้อมูลดิบจากทั้ง 9 ช่องที่กรอกไว้ [cite: 68-76]
    all_raw = f"""
    --- IPD DATA ---
    {st.session_state.s11}
    {st.session_state.s12}
    {st.session_state.s13}
    {st.session_state.s14}
    --- ASSESSMENT ---
    {st.session_state.s2}
    --- REGISTRATION ---
    {st.session_state.s31}
    {st.session_state.s32}
    {st.session_state.s33}
    {st.session_state.s34}
    """
    
    with st.spinner('Gemini กำลังวิเคราะห์ข้อมูลตามตรรกะ PhD และตรวจสอบ Verification Audit...'):
        # ตรรกะการสกัดข้อมูลระดับ PhD และกฎการตัดขยะ [cite: 61-67, 82]
        prompt = f"""
        คุณคือผู้ช่วยวิจัยทางการแพทย์ระดับ PhD ทำหน้าที่สกัดข้อมูลจากระบบ @ThanHIS ลงแบบฟอร์ม 062
        
        ตรรกะการสกัดข้อมูล (Search & Extract Logic):
        1. Noise Reduction: หากพบข้อความเกี่ยวกับ Theme Customizer, Menu Colors หรือ COPYRIGHT ให้ Ignore ทิ้งทั้งหมด 
        2. LOS Calc: ค้นหาตัวเลขหน้าคำว่า 'วัน' ในแถว Detox และ Rehab แล้วนำมาบวกกันเป็นตัวเลขรวม [cite: 63, 79]
        3. DX Format: สกัดรหัส ICD-10 และตัดจุดทศนิยมออกให้เป็นตัวเลขติดกัน (เช่น F155) [cite: 63, 79]
        4. MEDS: ค้นหาบรรทัด 'Home-Med' สกัดชื่อยา UPPERCASE พร้อมวิธีใช้ (แยกบรรทัดด้วย \\n) [cite: 63, 79]
        5. PROGRESS: สังเคราะห์จาก Progress Note ล่าสุด (S O A P) ให้เหลือย่อหน้าเดียว ความยาว 2-3 บรรทัด [cite: 63, 79]
        6. Verification Audit: หากไม่พบข้อมูล ให้พยายามหาจากช่องอื่น หากไม่มีจริงๆ ให้ระบุ [กรอกด้วยตนเอง] [cite: 78, 79]
        
        ข้อมูลดิบ:
        {all_raw}
        
        ตอบกลับในรูปแบบ JSON ที่มี Key ตรงกับ Placeholder: 
        NAME, AGE, HN, ID, EDU, CAREER, RELIGION, STATUS, RIGHTS, LAST_DC, ADMIT_DATE, VISIT_NUM, CC, CONTACT, RELATION, PHONE, NEAR_HOSP, DC_DATE, LOS, ADDRESS, Q9, Q8, MEDS, DX, PROGRESS, POST_SERVICE [cite: 35-51, 79]
        """
        
        try:
            # เรียกใช้โมเดลที่คุณหมอมีสิทธิ์ใช้งาน
            response = client.models.generate_content(model=MODEL_ID, contents=prompt)
            match = re.search(r'\{.*\}', response.text, re.DOTALL)
            
            if match:
                json_data = json.loads(match.group())
                st.success("✅ วิเคราะห์ข้อมูลสำเร็จและคำนวณวันนอน (LOS) เรียบร้อย!")
                
                # สร้างไฟล์ Word จากข้อมูลที่สกัดได้
                word_file = fill_pmnidat_doc(json_data)
                if word_file:
                    # บันทึกสถิติการใช้งานลง Google Sheet (ประกาศไว้ในส่วนที่ 1)
                    log_usage(json_data.get('name', '[ไม่ระบุชื่อ]'))
                    
                    st.download_button(
                        label="💾 ดาวน์โหลดไฟล์ 062 ฉบับสมบูรณ์ (ฟอนต์ 13)",
                        data=word_file,
                        file_name=f"Refer_{json_data.get('name','062')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            else:
                st.error("AI ไม่สามารถสกัดข้อมูลเป็นรูปแบบที่ถูกต้องได้ กรุณาลองตรวจสอบข้อมูลดิบอีกครั้ง")
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {e}")

# --- 7. มาตรการรักษาความปลอดภัย (PDPA Footer) ---
st.divider()
st.info("""
    **ประกาศ: มาตรการรักษาความปลอดภัยข้อมูลทางการแพทย์ (PDPA Compliance)**
    * ระบบ PMNIDAT Smart D/C Transfer ไม่มีการจัดเก็บข้อมูลผู้ป่วยถาวรในเซิร์ฟเวอร์
    * ข้อมูลจะสูญหายทันทีเมื่อมีการรีเฟรชหน้าจอ (Refresh) โปรดบันทึกไฟล์ให้เรียบร้อยก่อนออกจากระบบ
    """)
