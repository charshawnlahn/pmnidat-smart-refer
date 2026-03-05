import streamlit as st
from google import genai
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re
import json
import requests
import datetime

# --- 1. การจัดการวันที่ภาษาไทย (Thai Date Formatting) ---
THAI_MONTHS = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
]

def get_thai_date_parts():
    """ส่งคืน วัน, เดือน, และ พ.ศ. ปัจจุบัน สำหรับเติมอัตโนมัติ"""
    now = datetime.datetime.now()
    return {
        "DAY": str(now.day),
        "MONTH": THAI_MONTHS[now.month - 1],
        "YEAR": str(now.year + 543)
    }

# --- 2. ฟังก์ชันบันทึกสถิติ (Logging) ---
def log_usage(patient_name):
    try:
        url = st.secrets.get("APPS_SCRIPT_URL", "")
        if url:
            requests.post(url, json={"name": patient_name}, timeout=5)
    except:
        pass

# --- 3. การเชื่อมต่อ API ---
try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
    client = genai.Client(api_key=API_KEY)
    MODEL_ID = "gemini-1.5-flash"
except Exception as e:
    st.error(f"❌ ระบบเชื่อมต่อผิดพลาด: {e}")
    st.stop()

# --- 4. ระบบจดจำข้อมูล (Session State) ---
field_keys = ['s11', 's12', 's13', 's14', 's2', 's31', 's32', 's33', 's34']
for key in field_keys:
    if key not in st.session_state:
        st.session_state[key] = ""

if "extracted_json_data" not in st.session_state:
    st.session_state.extracted_json_data = None


# --- สิ้นสุดส่วนที่ 1 ---



# --- 5. ฟังก์ชันโหลดข้อมูลตัวอย่าง (Smart Sample Data Engine) ---
# บรรจุรายละเอียดครบถ้วนเพื่อใช้ทดสอบตรรกะการสกัดข้อมูลระดับ PhD 

def load_test_data():
    """โหลดข้อมูลตัวอย่าง 'นาย ชาย ธัญญารักษ์' เข้าสู่ Session State [cite: 5, 11, 18-20]"""
    
    # 1.1 ข้อมูลจาก Admission Note (IPD) [cite: 5, 11, 56-58]
    st.session_state.s11 = (
        "นาย ชาย ธัญญารักษ์ อายุ 40 ปี 5 เดือน\n"
        "สิทธิ์ : จ่ายตรงกรมบัญชีกลาง\n"
        "Admit Date 01/03/2569\n"
        "จำนวนวัน Detox [5] วัน Rehab [10] วัน\n"
        "CC : เสพสุราซ้ำ ต้องการเข้ารับการบำบัดรักษา\n"
        "เคยมานอน รพ. 4 ครั้ง จำหน่ายวันที่ 25 กันยายน 2568\n"
        "Admit Date ครั้งนี้ 01/03/2569"
    )
    
    # 1.2 ข้อมูลการวินิจฉัย (ICD-10 พร้อมชื่อไทย) [cite: 18, 72-74]
    st.session_state.s12 = (
        "1. F105 - โรคจิตจากสุรา (Alcohol Psychosis)\n"
        "2. I10 - โรคความดันโลหิตสูง (Hypertension)"
    )
    
    # 1.3 ข้อมูลรายการยา (Order / Meds) [cite: 19, 42]
    st.session_state.s13 = (
        "1. AMLODIPINE 5 MG 1x1 pc (เช้า)\n"
        "2. QUETIAPINE 25 MG 1 tab hs (ก่อนนอน)\n"
        "(รายการยา Home-Med ทั้งหมด)"
    )
    
    # 1.4 ข้อมูลบันทึกการติดตามอาการ (Progress Note)  [cite: 44, 102-104]
    st.session_state.s14 = (
        "S: สบายดี กินข้าวได้ นอนหลับได้\n"
        "O: V/S stable, BP 120/80 mmHg\n"
        "A: อาการคงที่ เตรียมจำหน่าย\n"
        "P: Discharge to home"
    )
    
    # 2. ข้อมูลการประเมิน (Assessment Score) [cite: 16, 67-71]
    st.session_state.s2 = "9Q : 5 คะแนน, 8Q : 0 คะแนน, BPRS : 15 คะแนน [cite: 16, 67-71]"
    
    # 3.1 ข้อมูลทั่วไป (Registration 1)  [cite: 7, 37-44, 96]
    st.session_state.s31 = (
        "Hospital Number 690099999 ชื่อ [ชาย] นามสกุล [ธัญญารักษ์]\n"
        "เลขบัตรประชาชน [1-2345-67890-12-3]\n"
        "ศาสนา [พุทธ] อาชีพ [ข้าราชการ] สถานภาพ [สมรส] การศึกษา [ปริญญาตรี]"
    )
    
    # 3.2 ที่อยู่ปัจจุบัน (Registration 2) [cite: 15, 62]
    st.session_state.s32 = "ที่อยู่ปัจจุบัน: เลขที่ 123 ต.หลักหก อ.เมืองปทุมธานี จ.ปทุมธานี"
    
    # 3.3 ผู้ติดต่อ (Contact Info) [cite: 12, 59, 63-64]
    st.session_state.s33 = "คุณ หญิง ธัญญารักษ์ (ภรรยา) เบอร์โทร: 081-234-XXXX"
    
    # 3.4 สิทธิการรักษา (Rights)  [cite: 8, 45-53, 97]
    st.session_state.s34 = "สิทธิหลัก [จ่ายตรง] สถานพยาบาลหลัก [สถาบันบำบัดรักษาและฟื้นฟูผู้ติดยาเสพติดแห่งชาติบรมราชชนนี]"
    
    st.rerun()

# --- 6. ฟังก์ชันล้างข้อมูล (Clear Data) ---
def clear_all_data():
    """ล้างข้อมูลดิบทั้งหมดออกจากระบบเพื่อรักษาความลับคนไข้ (PDPA) [cite: 53-54]"""
    for key in field_keys:
        st.session_state[key] = ""
    st.session_state.extracted_json_data = None
    st.rerun()



# --- สิ้นสุดส่วนที่ 2 ---





# --- 7. การจัดการแถบเมนูข้าง (Sidebar Configuration) ---

with st.sidebar:
    # แสดงโลโก้สถาบันฯ และหัวข้อคู่มือ 
    st.image("https://pmnidathis.dms.go.th/static/media/health.1bfd961f.png", width=100)
    st.header("📖 คู่มือการใช้งานระบบ")
    st.write("ระบบช่วยสกัดข้อมูลเพื่อจัดทำแบบบันทึกการส่งต่อ (PMNIDAT 062)")
    
    # ปุ่มควบคุมหลัก: ใช้คอลัมน์เพื่อความสวยงามและใช้งานง่าย
    st.subheader("🛠️ เครื่องมือระบบ")
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        if st.button("🧬 ตัวอย่างข้อมูล", key="side_load_test", use_container_width=True, help="ใช้ข้อมูลจำลองของ นาย ชาย ธัญญารักษ์ เพื่อทดสอบระบบ"):
            load_test_data()
    with col_t2:
        if st.button("🧹 ล้างข้อมูล", key="side_clear_data", use_container_width=True, help="ลบข้อมูลทุกช่องเพื่อเริ่มเคสใหม่ (PDPA)"):
            clear_all_data()
            
    st.divider()
    
    # คำแนะนำพื้นฐานในการคัดลอกข้อมูลจากต้นทาง
    st.markdown("### **วิธีการคัดลอกข้อมูลจาก @ThanHIS**")
    st.info("""
    1. **คลิกเมาส์ซ้ายค้าง** ที่ต้นข้อความ ลากลงล่างให้ครอบคลุมข้อมูลทั้งหมด
    2. กด **Ctrl+C** (คัดลอก)
    3. มาที่หน้านี้ คลิกช่องที่ต้องการ แล้วกด **Ctrl+V** (วาง)
    """)
    
    # รายละเอียดแต่ละ Step ตามคู่มือฉบับจริง (รักษาข้อความเดิมครบถ้วน) 
    with st.expander("🟢 STEP 1: ระบบผู้ป่วยใน (IPD)", expanded=True):
        st.markdown("""
        **1.1 Admission Note:**
        - ดูข้อมูลคนไข้ → Admission note
        - คัดลอกข้อมูลทั้งหมด (ตั้งแต่ข้อมูลทั่วไปผู้ป่วย ลากลงไปจนสุด)
        
        **1.2 การวินิจฉัย:**
        - กดเปิด "การวินิจฉัย"
        - คัดลอกรหัส ICD-10 ชื่อโรค และประเภท (Principal & Comorbidity)
        
        **1.3 Order / Meds:**
        - กดเมนู **"Order"** - คัดลอก **Order + Medication ทั้งหมด** จากบนถึงล่างสุด
        
        **1.4 Progress Note:**
        - กดเมนู **"Progress note"** - คัดลอก **Progress note ทั้งหมด** จากบนถึงล่างสุด 
        """)

    with st.expander("🔵 STEP 2: การประเมิน (Assessment)"):
        st.markdown("""
        - กดเมนู "Admission note" → กดปุ่ม **"ข้อมูลผู้ป่วยนอก"**
        - เลื่อนลงล่างและกดที่หัวข้อ **Assessment**
        - คัดลอกผลคะแนน **9Q, 8Q, BPRS** ทั้งหมดมาวาง
        """)

    with st.expander("🟠 STEP 3: เวชระเบียน (Registration)"):
        st.markdown("""
        - เข้าระบบผู้ป่วยนอก → เวชระเบียน → ลงทะเบียนผู้ป่วยใหม่
        - ค้นหา HN เพื่อเข้าสู่หน้าข้อมูลหลัก
        
        **3.1 ทั่วไป 1:** - คัดลอก Hospital number พร้อมกับ ชื่อ, อายุ, เลขบัตรประชาชน, ศาสนา
        
        **3.2 ทั่วไป 2:** - กดแสดง **"ที่อยู่ปัจจุบัน"** แล้วคัดลอกทั้งหมด
        
        **3.3 ผู้ติดต่อ:** - คัดลอกชื่อญาติ, ความสัมพันธ์ และเบอร์โทร
        
        **3.4 สิทธิรักษา:** - คัดลอกสิทธิ์และ **"สถานพยาบาลหลัก"**
        """)
        
    st.divider()
    # แสดงข้อมูลรุ่นโปรแกรมและผู้พัฒนา (แก้ไขวงเล็บปิดเรียบร้อย)
    st.caption("PMNIDAT Smart Transfer v3.36")
    st.success("Created by Dr. Charshawn Lahnwong (5 Mar 2026)")



               
# --- สิ้นสุดส่วนที่ 3 ---



               
# --- 8. การออกแบบส่วนหัวข้อหลัก (Application Header) ---

# แสดงชื่อระบบและรุ่นเพื่อความเป็นทางการของหน้าจอหลัก [cite: 31-32]
st.title("🏥 PMNIDAT Smart Transfer")
st.subheader("ผู้ช่วยพิมพ์ 'แบบบันทึกข้อมูลเพื่อส่งต่อ (PMNIDAT 062)' โดยอัตโนมัติ (Version 3.36)")
st.divider()

# --- 9. ส่วนการรับข้อมูล 9 ส่วน (The 9 Input Fields) ---

# กลุ่มที่ 1: ระบบผู้ป่วยใน (IPD) - จัดวาง 4 คอลัมน์สำหรับข้อมูลทางคลินิก [cite: 56-58, 72-74]
st.markdown("### **🟢 Step 1: ระบบผู้ป่วยใน (IPD)**")
s1_cols = st.columns(4)

with s1_cols[0]:
    st.text_area(
        "1.1 Admission Note", 
        height=300,
        placeholder="คัดลอกข้อมูลแรกรับทั้งหมด...",
        key="s11", # แก้ไข: ใช้ key เพียงตัวเดียวและไม่ซ้ำซ้อน
        help="คัดลอกจากเมนู Admission note ในระบบ IPD [cite: 56-58]"
    )

with s1_cols[1]:
    st.text_area(
        "1.2 การวินิจฉัย", 
        height=300,
        placeholder="คัดลอกรหัส ICD-10 ทั้งหมด...",
        key="s12", 
        help="คัดลอกจากเมนูการวินิจฉัยเพื่อสกัดรหัสโรคภาษาอังกฤษ [cite: 18]"
    )

with s1_cols[2]:
    st.text_area(
        "1.3 Order / Meds", 
        height=300,
        placeholder="คัดลอกข้อมูลจากเมนู Order ทั้งหมด...",
        key="s13", 
        help="คัดลอก Order และ Medication เพื่อสกัดยา Home-Med"
    )

with s1_cols[3]:
    st.text_area(
        "1.4 Progress Note", 
        height=300,
        placeholder="คัดลอกบันทึก Progress note ทั้งหมด...",
        key="s14", 
        help="คัดลอกบันทึกการติดตามอาการเพื่อสังเคราะห์ปัญหาที่ส่งต่อ  [cite: 44, 102-104]"
    )

st.divider()

# กลุ่มที่ 2: การประเมิน (Assessment) - ช่องกว้างพิเศษสำหรับคะแนนสุขภาพจิต [cite: 16, 67-71]
st.markdown("### **🔵 Step 2: การประเมิน (Assessment)**")
st.text_area(
    "คัดลอกผลคะแนน 9Q, 8Q, BPRS ทั้งหมดมาวางที่นี่", 
    height=150,
    placeholder="คะแนน 9Q, 8Q, BPRS ...",
    key="s2",
    help="ดึงจากหน้า Assessment เพื่อวิเคราะห์ภาวะสุขภาพจิต [cite: 16, 67-71]"
)

st.divider()

# กลุ่มที่ 3: เวชระเบียน (Registration) - แบ่ง 4 ส่วนตามทะเบียนประวัติคนไข้  [cite: 37-53, 59-64, 96-97]
st.markdown("### **🟠 Step 3: เวชระเบียน (Registration)**")
s3_cols = st.columns(4)

with s3_cols[0]:
    st.text_area(
        "3.1 ข้อมูลทั่วไป", 
        height=200,
        placeholder="HN, ชื่อ, อายุ, เลขบัตรประชาชน...",
        key="s31",
        help="ดึงจากหน้า 'ทั่วไป 1' ในระบบเวชระเบียน  [cite: 37-44, 96]"
    )

with s3_cols[1]:
    st.text_area(
        "3.2 ที่อยู่ปัจจุบัน", 
        height=200,
        placeholder="ที่อยู่ปัจจุบัน...",
        key="s32",
        help="ดึงจากหน้า 'ทั่วไป 2' ในระบบเวชระเบียน"
    )

with s3_cols[2]:
    st.text_area(
        "3.3 ผู้ติดต่อ", 
        height=200,
        placeholder="ชื่อญาติ ความสัมพันธ์ เบอร์โทรศัพท์...",
        key="s33",
        help="ดึงจากหน้า 'ผู้ติดต่อ' เพื่อใช้เป็นข้อมูล Contact [cite: 12, 59, 63-64]"
    )

with s3_cols[3]:
    st.text_area(
        "3.4 สิทธิการรักษา", 
        height=200,
        placeholder="สิทธิการรักษา...",
        key="s34",
        help="คัดลอกจากหน้า 'สิทธิการรักษา' เพื่อใช้ในตรรกะ Checkbox R1-R4  [cite: 45-53, 97]"
    )




# --- สิ้นสุดส่วนที่ 4 ---




# --- 10. ระบบประมวลผลอัจฉริยะ (AI Extraction Engine) ---

# ปุ่มกดสำหรับเริ่มกระบวนการสกัดข้อมูลเชิงลึก
if st.button("🚀 กดเพื่อประมวลผลและสกัดข้อมูลด้วย Gemini 1.5 Flash", key="btn_process_ai", use_container_width=True):
    # 10.1 เตรียมข้อมูลวันที่และรวบรวมข้อมูลดิบจาก 9 ส่วน  [cite: 63-64, 93]
    thai_date_data = get_thai_date_parts()
    today_str = f"{thai_date_data['DAY']} {thai_date_data['MONTH']} {thai_date_data['YEAR']}"
    
    all_raw_data = f"""
    วันที่ปัจจุบัน: {today_str}
    [1.1 Admission]: {st.session_state.s11}
    [1.2 Diagnosis]: {st.session_state.s12}
    [1.3 Medication]: {st.session_state.s13}
    [1.4 Progress]: {st.session_state.s14}
    [2. Assessments]: {st.session_state.s2}
    [3.1 Demographics]: {st.session_state.s31}
    [3.2 Address]: {st.session_state.s32}
    [3.3 Contact]: {st.session_state.s33}
    [3.4 Rights]: {st.session_state.s34}
    """
    
    with st.spinner('Gemini กำลังวิเคราะห์ข้อมูลและตรวจสอบตรรกะ Checkbox (R1-R4)...'):
        # 10.2 Prompt ระดับ PhD สำหรับการสกัดข้อมูลลงไฟล์แม่แบบล่าสุด  [cite: 18, 96-97]
        prompt = f"""
        คุณคือผู้ช่วยวิจัยทางการแพทย์ระดับ PhD จงสกัดข้อมูลจากระบบ @ThanHIS ลงแบบฟอร์ม 062 
        โดยใช้ตรรกะ "Official Template Mirroring" ดังนี้:

        1. [Checkbox Logic (R1-R4)]: วิเคราะห์สิทธิการรักษาแล้วส่งค่า [✓] ลงใน Key ที่ถูกต้องเพียงหนึ่งเดียว และ [ ] ใน Key อื่น:
           - R1: ประกันสุขภาพ (บัตรทอง/UC)
           - R2: ประกันสังคม
           - R3: ท.74 ข้าราชการ/จ่ายตรง
           - R4: อื่นๆ (ระบุรายละเอียดใน R_NOTE)

        2. [Format Logic]: 
           - DX: รหัส ICD-10 และชื่อโรคภาษาไทย ต่อกันในแถวเดียว คั่นด้วยคอมม่า (,) 
           - MEDS: รายการยา Home-Med (UPPERCASE) พร้อมวิธีใช้ ต่อกันในแถวเดียว คั่นด้วยคอมม่า (,)  [cite: 18, 96-97]
           - PROGRESS: สังเคราะห์จาก Progress Note ให้เป็นย่อหน้าเดียว 2-3 บรรทัด

        3. [Calculation Logic]:
           - LOC: คำนวณระยะเวลาในชุมชน โดยนำวันที่ปัจจุบัน ({today_str}) ลบกับวันที่จำหน่ายครั้งล่าสุด
           - LOS: คำนวณวันนอนรวมจากวันที่รับไว้ครั้งนี้จนถึงปัจจุบัน  [cite: 35-36, 84-88]

        ข้อมูลดิบ: {all_raw_data}
        ตอบกลับเป็น JSON ที่มี Key ดังนี้: 
        NAME, AGE, HN, ID, EDU, CAREER, RELIGION, STATUS, R1, R2, R3, R4, R_NOTE, LAST_DC, LOC, ADMIT_DATE, VISIT_NUM, CC, CONTACT, RELATION, PHONE, NEAR_HOSP, LOS, ADDRESS, Q9, Q8, MEDS, DX, PROGRESS, POST_SERVICE
        """
        
        try:
            # เรียกใช้งานโมเดล Gemini
            response = client.models.generate_content(model=MODEL_ID, contents=prompt)
            # สกัด JSON จากข้อความตอบกลับ
            match = re.search(r'\{.*\}', response.text, re.DOTALL)
            
            if match:
                st.session_state.extracted_json_data = json.loads(match.group())
                # ผสานวันที่ปัจจุบันเข้าไปในชุดข้อมูล  [cite: 63-64, 93]
                st.session_state.extracted_json_data.update(thai_date_data)
                st.success("✅ วิเคราะห์ข้อมูลเชิงลึกและสกัดค่าตัวแปร (Variables) สำเร็จ!")
            else:
                st.error("❌ ไม่สามารถประมวลผลโครงสร้างข้อมูลได้ กรุณาตรวจสอบข้อมูลดิบอีกครั้ง")
        except Exception as e:
            st.error(f"⚠️ เกิดข้อผิดพลาดในการประมวลผลด้วย AI: {e}")



# --- สิ้นสุดส่วนที่ 5 ---




# --- 11. ฟังก์ชันจัดการไฟล์ Word (Advanced Template Mapping) ---

def fill_pmnidat_doc(data):
    """ฟังก์ชันนำข้อมูลจากระบบเข้าสู่ไฟล์แม่แบบ .docx โดยรองรับโครงสร้างตารางและ Checkbox  [cite: 61-67, 92-106]"""
    try:
        # 11.1 โหลดไฟล์แม่แบบที่คุณหมอจัดทำล่าสุด (ต้องวางไฟล์ไว้ในโฟลเดอร์เดียวกับ app.py) [cite: 61]
        doc = Document("PMNIDAT 062 แบบบันทึกข้อมูลเพื่อส่งต่อ.docx")
        
        # 11.2 เตรียมชุดข้อมูลสำหรับแทนที่ โดยแปลง Key เป็นตัวพิมพ์ใหญ่ทั้งหมด
        mapping = {f"{{{{{k.upper()}}}}}": str(v) for k, v in data.items()}
        
        def apply_style_and_replace(paragraph):
            """ตรรกะการจัดวาง Alignment และการแทนที่ข้อความเชิงวิชาชีพ"""
            
            # ก. ตรรกะการจัดวางบรรทัด (Alignment Logic): 
            # วันที่หัวเอกสารให้ชิดขวา แต่จุดอื่นรวมถึง 'วันที่จำหน่าย' ให้ชิดซ้ายปกติ  [cite: 63, 67, 92-95]
            if "วันที่" in paragraph.text and "จำหน่าย" not in paragraph.text:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

            # ข. ค้นหาและแทนที่ข้อความ (Replacement Logic)
            for key, value in mapping.items():
                if key in paragraph.text:
                    # แทนที่ Placeholder (เช่น {{R1}} จะกลายเป็น [✓] หรือ [ ] ตามผลวิเคราะห์จาก AI)
                    paragraph.text = paragraph.text.replace(key, value)
                    
                    # ค. กำหนดคุณลักษณะตัวอักษร (Font Properties)
                    for run in paragraph.runs:
                        run.font.size = Pt(13) # บังคับขนาดฟอนต์ 13 pt ตลอดเอกสาร
                        
                        # ง. ตรรกะการเน้นตัวหนา (Bolding):
                        # เน้นตัวหนาเฉพาะข้อมูลทางคลินิก (DX, MEDS, PROGRESS, CC) เพื่อให้อ่านง่าย
                        clinical_keys = ["{{DX}}", "{{MEDS}}", "{{PROGRESS}}", "{{CC}}"]
                        if any(ck.upper() in key for ck in clinical_keys):
                            run.font.bold = True

        # 11.3 ดำเนินการตรวจสอบและแทนที่ในย่อหน้าปกติ  [cite: 61-65, 92-95]
        for p in doc.paragraphs:
            apply_style_and_replace(p)
            
        # 11.4 ดำเนินการตรวจสอบภายในตาราง (Tables) 
        # ไฟล์ใหม่ของคุณหมอใช้ตารางล็อคตำแหน่งข้อมูล จึงต้องวนลูปตรวจสอบทุก Cell
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        apply_style_and_replace(p)
                            
        # 11.5 บันทึกไฟล์ลงในหน่วยความจำชั่วคราว (Memory Buffer) เพื่อพร้อมสำหรับการดาวน์โหลด
        buffer = io.BytesIO()
        doc.save(buffer)
        return buffer.getvalue()
        
    except Exception as e:
        st.error(f"⚠️ ไม่สามารถประมวลผลไฟล์ Word ได้: {e}")
        return None




# --- สิ้นสุดส่วนที่ 6 ---





# --- 12. การแสดงผลลัพธ์และการควบคุมการดาวน์โหลด (Final Execution) ---

# ตรวจสอบว่ามีข้อมูลที่สกัดสำเร็จอยู่ใน Session State หรือไม่
if st.session_state.extracted_json_data:
    # เรียกใช้ฟังก์ชันจัดการไฟล์ Word เพื่อสร้างเอกสารจากแม่แบบ
    word_file_final = fill_pmnidat_doc(st.session_state.extracted_json_data)
    
    if word_file_final:
        # บันทึกสถิติการใช้งาน (อ้างอิงฟังก์ชันในส่วนที่ 1)
        log_usage(st.session_state.extracted_json_data.get('NAME', '[ไม่ระบุชื่อ]'))
        
        st.divider()
        st.balloons() # แสดงเอฟเฟกต์เฉลิมฉลองเมื่อเอกสารพร้อมใช้งาน
        st.success("🎉 ระบบจัดทำเอกสาร PMNIDAT 062 ฉบับสมบูรณ์เรียบร้อยแล้ว!")
        
        # แสดงปุ่มดาวน์โหลดไฟล์ที่สอดคล้องกับชื่อผู้ป่วย
        file_name_output = f"Refer_{st.session_state.extracted_json_data.get('NAME', '062')}.docx"
        
        st.download_button(
            label=f"💾 ดาวน์โหลดไฟล์: {file_name_output}",
            data=word_file_final,
            file_name=file_name_output,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key="btn_download_final",
            use_container_width=True
        )

# --- 13. มาตรการรักษาความปลอดภัยและความเป็นส่วนตัว (PDPA & Safety Footer) ---
# ส่วนประกาศสำคัญเพื่อรักษามาตรฐานความลับของผู้ป่วยตามระเบียบสถาบันฯ

st.divider()
st.info("""
    **🛡️ มาตรการรักษาความปลอดภัยของข้อมูลคนไข้ (PDPA Compliance):**
    
    * **ความปลอดภัยข้อมูล:** ระบบประมวลผลแบบ Real-time บนหน่วยความจำชั่วคราว (RAM) เท่านั้น **ไม่มีการจัดเก็บข้อมูลส่วนบุคคลของผู้ป่วย** ลงในฐานข้อมูลถาวร
    * **ระบบ Session-Based:** ข้อมูลจะถูกลบทิ้งทันทีเมื่อมีการรีเฟรชหน้าจอหรือปิดเบราว์เซอร์ โปรดดาวน์โหลดไฟล์ให้เรียบร้อยก่อนจบการทำงานในแต่ละเคส
    * **การตรวจสอบความถูกต้อง:** เนื่องจากเป็นระบบ AI ช่วยสกัดข้อมูล ผู้ใช้งานต้องตรวจสอบความถูกต้องและลงนามรับรองในเอกสารฉบับจริงทุกครั้ง
    
    ---
    **🔬 พัฒนาระบบโดย:** ดร.นพ.ชาฌาน หลานวงศ์ (หมออาร์ม) | MD-PhD (Pharmacology) 
    แพทย์เวชศาสตร์สารเสพติด สถาบันบำบัดรักษาและฟื้นฟูผู้ติดยาเสพติดแห่งชาติบรมราชชนนี (สบยช.)
    
    **💡 ติดต่อสอบถาม/แจ้งปัญหาการใช้งาน:**
    * **Line:** armlahn | **โทร:** 094-991-4599
    """)

# --- สิ้นสุดการทำงานของระบบ PMNIDAT Smart Transfer v3.36 ---


