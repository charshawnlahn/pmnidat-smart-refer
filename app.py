import streamlit as st
from google import genai 
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re
import json
import requests

# --- 1. การประกาศฟังก์ชันเสริม (Helper Functions) ---
def log_usage(patient_name):
    """บันทึกสถิติการใช้งานลง Google Sheets"""
    try:
        url = st.secrets["APPS_SCRIPT_URL"]
        requests.post(url, json={"name": patient_name}, timeout=5)
    except:
        pass

# --- 2. การตั้งค่าการเชื่อมต่อ API (รองรับ Paid Tier 1) ---
try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
    client = genai.Client(api_key=API_KEY)
    
    @st.cache_resource
    def find_active_model():
        try:
            # ตรวจสอบโมเดลที่ใช้งานได้จริงเพื่อป้องกัน 404
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

# --- 3. ระบบ Session State Memory (จดจำข้อมูล 9 ช่อง) ---
field_keys = ['s11', 's12', 's13', 's14', 's2', 's31', 's32', 's33', 's34']
for key in field_keys:
    if key not in st.session_state:
        st.session_state[key] = ""

def load_test_data():
    """ลงข้อมูลตัวอย่าง: นพ.ชาฌาน หลานวงศ์ (จัดรูปแบบโรคไทยและยารายบรรทัด)"""
    st.session_state.s11 = "นพ.ชาฌาน หลานวงศ์ อายุ 40 ปี 5 เดือน\nสิทธิ์ : จ่ายตรงกรมบัญชีกลาง\nAdmit Date 01/03/2569\nจำนวนวัน Detox [5] วัน Rehab [10] วัน\nCC : ต้องการทดสอบระบบสกัดข้อมูลระดับ PhD"
    
    # รายการโรคพร้อมชื่อภาษาไทย (เรียงบรรทัด) [cite: 41-43]
    st.session_state.s12 = "1. F15.5 - โรคจิตจากสารกระตุ้น (Amphetamine Psychosis)\n2. I10 - โรคความดันโลหิตสูง (Essential Hypertension)"
    
    # รายการยา Home-Med (เรียงแถวละ 1 รายการ) [cite: 43]
    st.session_state.s13 = "1. AMLODIPINE 5 MG 1x1 pc (เช้า)\n2. QUETIAPINE 25 MG 1 tab hs (ก่อนนอน)\n(Home-Med ทั้งหมด)"
    
    st.session_state.s14 = "S: สบายดี กินข้าวได้\nO: V/S stable, BP 120/80 mmHg\nA: อาการคงที่ เตรียมจำหน่าย\nP: Discharge to home"
    st.session_state.s2 = "9Q : 5 คะแนน\n8Q : 0 คะแนน\nBPRS : 15 คะแนน"
    st.session_state.s31 = "ชื่อ [นพ.ชาฌาน] นามสกุล [หลานวงศ์]\nเลขที่บัตรประชาชน* [1-2345-67890-12-3]\nศาสนา [00] [พุทธ] อาชีพ [แพทย์]"
    st.session_state.s32 = "ที่อยู่ปัจจุบัน : สถาบันบำบัดรักษาและฟื้นฟูผู้ติดยาเสพติดแห่งชาติบรมราชชนนี (สบยช.) ต.หลักหก อ.เมืองปทุมธานี จ.ปทุมธานี"
    st.session_state.s33 = "คุณวิไล หลานวงศ์ (ภรรยา) เบอร์โทรศัพท์ 081-234-XXXX"
    st.session_state.s34 = "สิทธิหลัก [จ่ายตรง] สถานพยาบาลหลัก [สถาบันบำบัดรักษาและฟื้นฟูผู้ติดยาเสพติดแห่งชาติบรมราชชนนี]"
    st.rerun()

def clear_all_data():
    for key in field_keys:
        st.session_state[key] = ""
    st.rerun()

st.title("🏥 PMNIDAT Smart D/C Transfer")
st.subheader("ระบบสร้างไฟล์ใบส่งต่ออัตโนมัติ (Master Version 3.31)")


# --- 4. การออกแบบแถบเมนูด้านข้าง (Sidebar Manual & Controls) ---

with st.sidebar:
    # แสดงโลโก้สถาบันฯ และหัวข้อคู่มือ
    st.image("https://pmnidathis.dms.go.th/static/media/health.1bfd961f.png", width=100)
    st.header("📖 คู่มือพี่พยาบาล ฉบับจับมือทำ")
    
    # ปุ่มควบคุมสำหรับการทดสอบระบบ (Test Drive & Clear)
    st.subheader("🛠️ เครื่องมือช่วยงาน")
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        if st.button("🧬 ลงข้อมูลตัวอย่าง", use_container_width=True, help="คลิกเพื่อลงข้อมูล นพ.ชาฌาน สำหรับทดสอบระบบ"):
            load_test_data()
    with col_t2:
        if st.button("🧹 ล้างข้อมูล", use_container_width=True, help="ล้างข้อมูลทั้งหมดในช่องกรอก"):
            clear_all_data()
            
    st.divider()
    
    # คำแนะนำพื้นฐานการคัดลอกข้อมูล
    st.markdown("### **วิธีการคัดลอกข้อมูลจาก @ThanHIS**")
    st.info("""
    1. **คลิกเมาส์ซ้ายค้าง** ที่ต้นข้อความ ลากครอบให้คลุมทั้งหมด
    2. กด **Ctrl+C** (คัดลอก)
    3. มาที่หน้านี้ คลิกช่องที่ต้องการ แล้วกด **Ctrl+V** (วาง)
    """)
    
    # รายละเอียดขั้นตอนการคัดลอกแบบแยกบรรทัดชัดเจน
    with st.expander("🟢 STEP 1: ระบบผู้ป่วยใน (IPD)", expanded=True):
        st.markdown("""
        **1.1 Admission Note:**
        - ดูข้อมูลคนไข้ → Admission note
        - [cite_start]คัดลอกข้อมูลทั้งหมด [cite: 38]
        
        **1.2 การวินิจฉัย:**
        - กดเมนู "การวินิจฉัย"
        - [cite_start]คัดลอกรหัส ICD-10 [cite: 41]
        
        **1.3 Order / Meds:**
        - กดเมนู "Order" ด้านซ้าย
        - [cite_start]คัดลอก Discharge order + Home medication [cite: 43]
        
        **1.4 Progress Note:**
        - กดเมนู "Progress note" ด้านซ้าย
        - [cite_start]คัดลอกบันทึกล่าสุด (SOAP) [cite: 43]
        """)

    with st.expander("🔵 STEP 2: การประเมิน (Assessment)"):
        st.markdown("""
        - [cite_start]กดเมนู "Admission note" → ปุ่ม **ข้อมูลผู้ป่วยนอก** [cite: 44]
        - เลื่อนลงล่างไปที่หัวข้อ **Assessment**
        - [cite_start]คัดลอกผลคะแนน 9Q, 8Q, BPRS [cite: 44]
        """)

    with st.expander("🟠 STEP 3: เวชระเบียน (Registration)"):
        st.markdown("""
        - ระบบผู้ป่วยนอก → เวชระเบียน → ลงทะเบียนผู้ป่วย
        - [cite_start]ค้นหา HN เพื่อดูข้อมูลคนไข้ [cite: 44]
        
        [cite_start]**3.1 ทั่วไป 1:** คัดลอกชื่อ, อายุ, เลขบัตรประชาชน, ศาสนา [cite: 44]
        [cite_start]**3.2 ทั่วไป 2:** กดแสดง **ที่อยู่ปัจจุบัน** แล้วคัดลอก [cite: 45]
        [cite_start]**3.3 ผู้ติดต่อ:** คัดลอกชื่อญาติและเบอร์โทรศัพท์ [cite: 45]
        [cite_start]**3.4 สิทธิรักษา:** คัดลอกสิทธิ์และ **สถานพยาบาลหลัก** [cite: 46]
        """)
        
    st.divider()
    # แสดงรุ่นโมเดลที่กำลังทำงาน
    st.success(f"💡 AI วิเคราะห์ด้วยตรรกะ PhD ผ่าน: {MODEL_ID}")

# --- สิ้นสุดส่วนที่ 2 ---


# --- 5. การออกแบบส่วนกรอกข้อมูล (9 ช่องรับข้อมูล) ---

st.divider()
# กลุ่มที่ 1: ข้อมูลจากระบบผู้ป่วยใน (IPD) [cite: 32, 33]
st.markdown("### **🟢 Step 1: ระบบผู้ป่วยใน (IPD)**")
s1_cols = st.columns(4)

with s1_cols[0]:
    st.text_area(
        "1.1 Admission Note",
        height=250,
        placeholder="วางข้อมูลแรกรับและวันนอน...",
        key="s11",  # เชื่อมกับ st.session_state.s11 [cite: 38]
        help="คัดลอกจากหน้า Admission note เพื่อสกัดอาการนำส่งและวันนอน"
    )

with s1_cols[1]:
    st.text_area(
        "1.2 การวินิจฉัย",
        height=250,
        placeholder="วางรหัส ICD-10 ทั้งหมด...",
        key="s12",  # เชื่อมกับ st.session_state.s12 [cite: 41]
        help="คัดลอกจากเมนูการวินิจฉัย (Principal & Comorbidity)"
    )

with s1_cols[2]:
    st.text_area(
        "1.3 Order / Meds",
        height=250,
        placeholder="วางรายการยา Home-Med...",
        key="s13",  # เชื่อมกับ st.session_state.s13 [cite: 43]
        help="คัดลอกจากเมนู Order (เน้นรายการยา Home-Med)"
    )

with s1_cols[3]:
    st.text_area(
        "1.4 Progress Note",
        height=250,
        placeholder="วางบันทึก SOAP ล่าสุด...",
        key="s14",  # เชื่อมกับ st.session_state.s14
        help="คัดลอกจากหน้า Progress note เพื่อสรุปปัญหา"
    )

st.divider()
# กลุ่มที่ 2: ข้อมูลการประเมิน (Assessment) [cite: 34, 35]
st.markdown("### **🔵 Step 2: การประเมิน (Assessment)**")
st.text_area(
    "คัดลอกผลคะแนน 9Q, 8Q, BPRS มาวางที่นี่",
    height=120,
    key="s2",  # เชื่อมกับ st.session_state.s2 [cite: 44]
    help="ดึงจากเมนู Assessment ในระบบผู้ป่วยนอก"
)

st.divider()
# กลุ่มที่ 3: ข้อมูลเวชระเบียน (Registration) [cite: 36, 37]
st.markdown("### **🟠 Step 3: เวชระเบียน (Registration)**")
s3_cols = st.columns(4)

with s3_cols[0]:
    st.text_area(
        "3.1 ข้อมูลทั่วไป", 
        height=200, 
        key="s31", # เชื่อมกับ st.session_state.s31 [cite: 37]
        help="คัดลอกจากหน้า 'ทั่วไป 1' (ชื่อ, อายุ, เลขบัตรประชาชน)"
    )

with s3_cols[1]:
    st.text_area(
        "3.2 ที่อยู่ปัจจุบัน", 
        height=200, 
        key="s32", # เชื่อมกับ st.session_state.s32 [cite: 45]
        help="คัดลอกจากหน้า 'ทั่วไป 2' (ที่อยู่ปัจจุบัน)"
    )

with s3_cols[2]:
    st.text_area(
        "3.3 ผู้ติดต่อ", 
        height=200, 
        key="s33", # เชื่อมกับ st.session_state.s33
        help="คัดลอกจากหน้า 'ผู้ติดต่อ' (ชื่อญาติและเบอร์โทร)"
    )

with s3_cols[3]:
    st.text_area(
        "3.4 สิทธิการรักษา", 
        height=200, 
        key="s34", # เชื่อมกับ st.session_state.s34 [cite: 46]
        help="คัดลอกจากหน้า 'สิทธิการรักษา' (สิทธิ์และ รพ.หลัก)"
    )

# --- สิ้นสุดส่วนที่ 3 ---

# --- 6. ส่วนประมวลผลอัจฉริยะ (Advanced Search & Extract Logic) ---

if st.button("🚀 สกัดข้อมูลและสร้างเอกสาร 062", use_container_width=True):
    # รวบรวมข้อมูลดิบจาก 9 ช่องที่กรอกไว้ในส่วนที่ 3 [cite: 102, 104, 106]
    all_raw_data = f"""
    --- GROUP 1: IPD DATA ---
    {st.session_state.s11}
    {st.session_state.s12}
    {st.session_state.s13}
    {st.session_state.s14}
    --- GROUP 2: ASSESSMENT ---
    {st.session_state.s2}
    --- GROUP 3: REGISTRATION ---
    {st.session_state.s31}
    {st.session_state.s32}
    {st.session_state.s33}
    {st.session_state.s34}
    """
    
    with st.spinner('Gemini 3 Flash กำลังวิเคราะห์ข้อมูลตามตรรกะ PhD และตรวจสอบ Verification Audit...'):
        # ตรรกะการตัดขยะและการสกัดข้อมูลตามสมอเรือ (Keywords Anchor) [cite: 100, 112]
        prompt = f"""
        คุณคือผู้ช่วยวิจัยระดับ PhD ทางการแพทย์ ทำหน้าที่สกัดข้อมูลจากระบบ @ThanHIS ลงแบบฟอร์ม 062 
        โดยต้องปฏิบัติตามกฎเหล็ก "Verification Audit" อย่างเคร่งครัด[cite: 107, 108]:

        1. กฎการตัดขยะข้อมูล (Noise Reduction Rule): 
           Ignore (ละทิ้ง) ข้อมูล Theme Customizer, Navbar, Menu Colors, Light/Dark Mode และ COPYRIGHT ทั้งหมด 

        2. ตรรกะการสกัดข้อมูลรายกลุ่ม (Search & Extract Logic):
           - [HN/AN]: มองหาตัวเลขหลังคำว่า 'HN' หรือ 'AN' [cite: 102]
           - [สิทธิการรักษา]: สกัดข้อความระหว่าง 'สิทธิ์ :' ถึง '(ไม่มีหนังสือส่งตัว)' หรือข้อความถัดไป [cite: 102]
           - [CC]: สกัดข้อความระหว่าง 'Chief Complaint' ถึง 'Present illness' [cite: 102]
           - [วันนอน (LOS)]: ค้นหาตัวเลขหน้าคำว่า 'วัน' ในแถว 'Detox' และ 'Rehab' แล้วนำมาบวกกันเสมอ [cite: 102, 109]
           - [รหัสโรค (DX)]: สกัดรหัส ICD-10 ตัดจุดทศนิยมออก และต้องระบุ "ชื่อโรคเป็นภาษาไทย" กำกับเสมอ เรียงบรรทัดละ 1 โรค [cite: 102, 109]
           - [ยา (MEDS)]: สกัดรายการที่มีคำว่า 'Home-Med' เขียนชื่อยาเป็น UPPERCASE พร้อมวิธีใช้ โดยต้อง "แยกบรรทัดละ 1 ตัวยา" เท่านั้น [cite: 102]
           - [คะแนนประเมิน]: สกัดตัวเลขหลัง 'ผลรวมการประเมินโรคซึมเศร้า' (9Q) และ 'การฆ่าตัวตาย' (8Q) [cite: 104]
           - [สรุปปัญหา (PROGRESS)]: สังเคราะห์จาก SOAP และ Discharge Order ให้เป็นย่อหน้าเดียว ความยาว 2-3 บรรทัด [cite: 102, 109]
           - [ข้อมูลทั่วไป]: สกัดชื่อ-สกุล, อายุ, บัตรประชาชน (13 หลัก), ศาสนา และอาชีพ จากกลุ่มเวชระเบียน [cite: 106, 109]
           - [หน่วยบริการ]: สกัดชื่อโรงพยาบาลหลังคำว่า 'สถานพยาบาลหลัก' [cite: 106]
           - [ผู้ดูแล]: สกัดชื่อ ความสัมพันธ์ และเบอร์โทรศัพท์ [cite: 106]

        3. นโยบายความถูกต้อง (Verification Audit):
           - หากไม่พบข้อมูลใดๆ ให้ระบุ [กรอกด้วยตนเอง] ห้ามเว้นว่างเด็ดขาด [cite: 108]
           - หากพบประวัติจำหน่ายเดิม (เช่น เคยนอน 4 ครั้ง) ให้วิเคราะห์ 'รับไว้ครั้งที่' เป็นลำดับถัดไป (เช่น 5) [cite: 109]

        ข้อมูลดิบสำหรับวิเคราะห์:
        {all_raw_data}

        ตอบกลับในรูปแบบ JSON ที่มี Key ตรงกับ Placeholder: 
        NAME, AGE, HN, ID, EDU, CAREER, RELIGION, STATUS, RIGHTS, LAST_DC, ADMIT_DATE, VISIT_NUM, CC, CONTACT, RELATION, PHONE, NEAR_HOSP, DC_DATE, LOS, ADDRESS, Q9, Q8, MEDS, DX, PROGRESS, POST_SERVICE
        """
        
        try:
            # ใช้ Gemini 3 Flash ประมวลผล
            response = client.models.generate_content(model=MODEL_ID, contents=prompt)
            match = re.search(r'\{.*\}', response.text, re.DOTALL)
            
            if match:
                # เก็บข้อมูล JSON ลงใน Session State เพื่อส่งต่อให้ส่วนที่ 5
                st.session_state.extracted_json = json.loads(match.group())
                st.success("✅ วิเคราะห์ข้อมูลและคำนวณวันนอน (LOS) สำเร็จตามตรรกะ PhD!")
            else:
                st.error("AI ไม่สามารถสร้างรูปแบบ JSON ที่ถูกต้องได้ กรุณาลองใหม่อีกครั้ง")
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {e}")

# --- สิ้นสุดส่วนที่ 4 ---


# --- 7. ฟังก์ชันจัดการไฟล์ Word (ชิดซ้าย + ฟอนต์ 13 + จัดระเบียบรายการ) ---

def fill_pmnidat_doc(data):
    """นำข้อมูลที่สกัดได้ไปวางในไฟล์แม่แบบและจัดรูปแบบ [cite: 49, 83-112]"""
    try:
        # เปิดไฟล์แม่แบบที่คุณหมอเตรียมไว้ [cite: 83]
        doc = Document("PMNIDAT 062 แบบบันทึกข้อมูลเพื่อส่งต่อ.docx")
        
        # เตรียมชุดข้อมูล Mapping โดยเปลี่ยน Key ให้เป็นตัวพิมพ์ใหญ่ตาม Placeholder [cite: 87-103]
        mapping = {f"{{{{{k.upper()}}}}}": str(v) for k, v in data.items()}
        
        def apply_style_and_replace(paragraph):
            for key, value in mapping.items():
                if key in paragraph.text:
                    # แทนที่ Placeholder ด้วยข้อมูลจริง [cite: 53-82]
                    paragraph.text = paragraph.text.replace(key, value)
                    # บังคับจัดรูปแบบชิดซ้าย (Left Alignment) เพื่อความเป็นระเบียบ
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    # ปรับขนาดฟอนต์เป็น 13 pt ตลอดทั้งเอกสาร [cite: 49]
                    for run in paragraph.runs:
                        run.font.size = Pt(13)

        # ดำเนินการแทนที่ข้อมูลทั้งในเนื้อหาปกติและภายในตาราง [cite: 87-112]
        for p in doc.paragraphs: apply_style_and_replace(p)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs: apply_style_and_replace(p)
                            
        # บันทึกไฟล์ลงใน Memory Buffer
        buffer = io.BytesIO()
        doc.save(buffer)
        return buffer.getvalue()
    except Exception as e:
        st.error(f"⚠️ เกิดข้อผิดพลาดในการสร้างไฟล์ Word: {e}")
        return None

# --- 8. การแสดงผลลัพธ์และปุ่มดาวน์โหลด (ต่อจากตรรกะในส่วนที่ 4) ---

# ตรวจสอบว่ามีการสกัดข้อมูลสำเร็จใน Session State หรือไม่
if "extracted_data" in st.session_state and st.session_state.extracted_data:
    # สร้างไฟล์ Word จากข้อมูลที่ AI สกัดได้ล่าสุด
    word_file = fill_pmnidat_doc(st.session_state.extracted_data)
    
    if word_file:
        # บันทึกสถิติการใช้งานลง Log Book
        log_usage(st.session_state.extracted_data.get('name', '[ไม่ระบุชื่อ]'))
        
        st.divider()
        st.balloons() # แสดงความยินดีเมื่อระบบพร้อม
        st.success("🎉 เอกสาร 062 พร้อมสำหรับการดาวน์โหลดแล้ว!")
        
        # ปุ่มดาวน์โหลดไฟล์ฉบับสมบูรณ์
        st.download_button(
            label="💾 ดาวน์โหลดไฟล์ 062 (จัดรูปแบบชิดซ้าย + ฟอนต์ 13)",
            data=word_file,
            file_name=f"Refer_{st.session_state.extracted_data.get('name', '062')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

# --- 9. มาตรการความปลอดภัยและประกาศ PDPA ---
st.divider()
st.info("""
    **ประกาศมาตรการรักษาความปลอดภัยข้อมูล (PDPA Compliance):**
    * ระบบนี้เป็นเพียงเครื่องมือช่วยสกัดข้อมูลทางการแพทย์ ข้อมูลทั้งหมดจะประมวลผลแบบ Real-time [cite: 47-49]
    * **ไม่มีการจัดเก็บข้อมูลผู้ป่วยถาวรบนเซิร์ฟเวอร์** ข้อมูลจะถูกลบทิ้งทันทีเมื่อปิดเบราว์เซอร์หรือรีเฟรชหน้าจอ [cite: 52]
    * โปรดตรวจสอบความถูกต้องของข้อมูล (Verification Audit) อีกครั้งก่อนนำไปใช้งานจริง [cite: 49]
    """)

