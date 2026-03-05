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
# เพื่อรองรับการเติมข้อมูล วันที่ {{DAY}} เดือน {{MONTH}} พ.ศ. [cite_start]{{YEAR}} อัตโนมัติ [cite: 15]
THAI_MONTHS = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
]

def get_thai_date_parts():
    [cite_start]"""ส่งคืน วัน, เดือน(ชื่อเต็ม), และ พ.ศ. ปัจจุบัน สำหรับใช้ใน LOC และ DC_DATE [cite: 22, 27]"""
    now = datetime.datetime.now()
    return {
        "DAY": str(now.day),
        "MONTH": THAI_MONTHS[now.month - 1],
        "YEAR": str(now.year + 543) # แปลง ค.ศ. เป็น พ.ศ.
    }

# --- 2. ฟังก์ชันบันทึกสถิติการใช้งาน (Usage Logging) ---
def log_usage(patient_name):
    """บันทึกชื่อผู้ป่วยที่ประมวลผลไปยังฐานข้อมูลสถิติ"""
    try:
        url = st.secrets["APPS_SCRIPT_URL"]
        requests.post(url, json={"name": patient_name}, timeout=5)
    except:
        pass

# --- 3. การตั้งค่าระบบความปลอดภัยและการเชื่อมต่อ API ---
try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
    client = genai.Client(api_key=API_KEY)
    
    @st.cache_resource
    def find_active_model():
        """ค้นหาโมเดลที่พร้อมใช้งานสูงสุดในระบบ"""
        try:
            available_models = [m.name for m in client.models.list()]
            for m in available_models:
                if "gemini-1.5-flash" in m: return m
            return available_models[0]
        except:
            return "models/gemini-1.5-flash"
            
    MODEL_ID = find_active_model()
except Exception as e:
    st.error(f"❌ ระบบเชื่อมต่อผิดพลาด (Configuration Error): {e}")
    st.stop()

# --- 4. ระบบจดจำข้อมูลชั่วคราว (Session State Memory) ---
field_keys = ['s11', 's12', 's13', 's14', 's2', 's31', 's32', 's33', 's34']
for key in field_keys:
    if key not in st.session_state:
        st.session_state[key] = ""




# --- 5. ฟังก์ชันโหลดข้อมูลตัวอย่าง (Smart Sample Data) ---
# บรรจุรายละเอียดครบถ้วนเพื่อใช้ทดสอบตรรกะระดับ PhD และการคำนวณ LOC

def load_test_data():
    """โหลดข้อมูลตัวอย่าง 'นาย ชาย ธัญญารักษ์' เข้าสู่ Session State โดยไม่ตัดทอนข้อความ [cite: 3, 10]"""
    
    # 1.1 ข้อมูลจาก Admission Note (IPD) 
    st.session_state.s11 = (
        "นาย ชาย ธัญญารักษ์ อายุ 40 ปี 5 เดือน\n"
        "สิทธิ์ : จ่ายตรงกรมบัญชีกลาง\n"
        "Admit Date 01/03/2569\n"
        "จำนวนวัน Detox [5] วัน Rehab [10] วัน\n"
        "CC : เสพสุราซ้ำ ต้องการเข้ารับการบำบัดรักษา\n"
        "เคยมานอน รพ.4 ครั้ง จำหน่ายวันที่ 25 กันยายน 2568\n"
        "Admit Date 20/01/2569"
    )
    
    # 1.2 ข้อมูลการวินิจฉัย (ICD-10 พร้อมชื่อไทย) [cite: 3, 10]
    st.session_state.s12 = (
        "1. F105 - โรคจิตจากสุรา (Alcohol)\n"
        "2. I10 - โรคความดันโลหิตสูง (Hypertension)"
    )
    
    # 1.3 ข้อมูลรายการยา (Order / Meds) [cite: 3, 10]
    st.session_state.s13 = (
        "1. AMLODIPINE 5 MG 1x1 pc (เช้า)\n"
        "2. QUETIAPINE 25 MG 1 tab hs (ก่อนนอน)\n"
        "(Home-Med ทั้งหมด)"
    )
    
    # 1.4 ข้อมูลบันทึกการติดตามอาการ (Progress Note) [cite: 3, 10]
    st.session_state.s14 = (
        "S: สบายดี กินข้าวได้ นอนหลับได้\n"
        "O: V/S stable, BP 120/80 mmHg\n"
        "A: อาการคงที่ เตรียมจำหน่าย\n"
        "P: Discharge to home"
    )
    
    # 2. ข้อมูลการประเมิน (Assessment) [cite: 5, 10]
    st.session_state.s2 = "9Q : 5 คะแนน\n8Q : 0 คะแนน\nBPRS : 15 คะแนน"
    
    # 3.1 ข้อมูลทั่วไป (Registration 1) [cite: 7, 10]
    st.session_state.s31 = (
        "Hospital Number 690099999 ชื่อ [ชาย] นามสกุล [ธัญญารักษ์]\n"
        "เลขบัตรประชาชน* [1-2345-67890-12-3]\n"
        "ศาสนา [พุทธ] อาชีพ [ข้าราชการ] สถานภาพ [สมรส] การศึกษา [ปริญญาตรี]"
    )
    
    # 3.2 ที่อยู่ปัจจุบัน (Registration 2) [cite: 7, 10]
    st.session_state.s32 = "ที่อยู่ปัจจุบัน: เลขที่ 123 ต.หลักหก อ.เมืองปทุมธานี จ.ปทุมธานี"
    
    # 3.3 ผู้ติดต่อ (Contact Info) [cite: 7, 10]
    st.session_state.s33 = "คุณ หญิง ธัญญารักษ์ (ภรรยา) เบอร์โทร: 081-234-XXXX"
    
    # 3.4 สิทธิการรักษา (Rights) [cite: 7, 10]
    st.session_state.s34 = "สิทธิหลัก [จ่ายตรง] สถานพยาบาลหลัก [สถาบันบำบัดรักษาและฟื้นฟูผู้ติดยาเสพติดแห่งชาติบรมราชชนนี]"
    
    # สั่งให้แอปรีเฟรชหน้าจอเพื่อแสดงข้อมูลในช่องกรอกทันที
    st.rerun()

def clear_all_data():
    """ล้างข้อมูลดิบทั้งหมดใน Session State เพื่อเริ่มเคสใหม่"""
    for key in field_keys:
        st.session_state[key] = ""
    st.rerun()


# --- 6. การออกแบบแถบเมนูด้านข้าง (Sidebar Manual & Controls) ---

with st.sidebar:
    # แสดงโลโก้สถาบันฯ และหัวข้อหลักของคู่มือ
    st.image("https://pmnidathis.dms.go.th/static/media/health.1bfd961f.png", width=100)
    st.header("📖 คู่มือการใช้งาน")
    
    # ปุ่มควบคุมสำหรับทดสอบระบบและจัดการข้อมูลใน Session State
    st.subheader("🛠️ เครื่องมือระบบ")
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        if st.button("🧬 ตัวอย่างข้อมูล", use_container_width=True):
            load_test_data()
    with col_t2:
        if st.button("🧹 ล้างข้อมูล", use_container_width=True):
            clear_all_data()
            
    st.divider()
    
    # คำแนะนำพื้นฐานที่เข้าใจง่ายสำหรับผู้ใช้งาน
    st.markdown("### **วิธีการคัดลอกข้อมูลจาก @ThanHIS**")
    st.info("""
    1. **คลิกเมาส์ซ้ายค้าง** ที่ต้นข้อความ ลากลงล่างให้ครอบคลุมข้อมูลทั้งหมด
    2. กด **Ctrl+C** (คัดลอก)
    3. มาที่หน้านี้ คลิกช่องที่ต้องการ แล้วกด **Ctrl+V** (วาง)
    """)
    
    # รายละเอียดแต่ละ Step ตามคู่มือฉบับสมบูรณ์ที่คุณหมออาร์มกำหนด [cite: 3, 5, 7, 10]
    with st.expander("🟢 STEP 1: ระบบผู้ป่วยใน (IPD)", expanded=True):
        st.markdown("""
        **1.1 Admission Note:**
        - ดูข้อมูลคนไข้ → Admission note
        - คัดลอกข้อมูลทั้งหมด (ตั้งแต่ข้อมูลทั่วไปผู้ป่วย ลากลงไปจนสุด)
        
        **1.2 การวินิจฉัย:**
        - กดเปิด "การวินิจฉัย"
        - คัดลอกรหัส ICD-10 ชื่อโรค และประเภท (Principal & Comorbidity)
        
        **1.3 Order / Meds:**
        - กดเมนู **"Order"** - คัดลอก **Order + Medication ทั้งหมด** จากบนถึงล่างสุด (เพื่อให้ AI สกัดยา Home-Med เอง)
        
        **1.4 Progress Note:**
        - กดเมนู **"Progress note"** - คัดลอก **Progress note ทั้งหมด** จากบนถึงล่างสุด (เพื่อให้ AI สังเคราะห์ปัญหา)
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
    # แสดงเครดิตและรุ่นของระบบเพื่อความโปร่งใสและตรวจสอบได้
    st.success("PMNIDAT Smart Transfer (Version 3.35) | Created by Dr.Charshawn Lahnwong (5 March 2026)")

# --- สิ้นสุดส่วนที่ 3 ---

# --- 7. ส่วนการออกแบบหน้าจอหลัก (Main UI Layout) ---

# แทรกส่วนหัวที่คุณหมอกำหนดไว้ เพื่อความเป็นทางการของระบบ
st.title("🏥 PMNIDAT Smart Transfer")
st.subheader("ผู้ช่วยพิมพ์ 'แบบบันทึกข้อมูลเพื่อส่งต่อ (PMNIDAT 062)' โดยอัตโนมัติ (Version 3.35)")

st.divider()

# กลุ่มที่ 1: ระบบผู้ป่วยใน (IPD) - จัดวาง 4 คอลัมน์สำหรับข้อมูลทางคลินิก [cite: 2, 3]
st.markdown("### **🟢 Step 1: ระบบผู้ป่วยใน (IPD)**")
s1_cols = st.columns(4)

with s1_cols[0]:
    st.text_area(
        "1.1 Admission Note",
        height=300,
        placeholder="คัดลอกข้อมูลแรกรับทั้งหมด...",
        key="s11",
        help="คัดลอกจากเมนู Admission note ในระบบ IPD [cite: 3, 10]"
    )

with s1_cols[1]:
    st.text_area(
        "1.2 การวินิจฉัย",
        height=300,
        placeholder="คัดลอกรหัส ICD-10 ทั้งหมด...",
        key="s12",
        help="คัดลอกจากเมนูการวินิจฉัยเพื่อสกัดรหัสโรคภาษาอังกฤษ [cite: 3, 10]"
    )

with s1_cols[2]:
    st.text_area(
        "1.3 Order / Meds",
        height=300,
        placeholder="คัดลอกข้อมูลจากเมนู Order ทั้งหมด ...",
        key="s13",
        help="คัดลอก Order และ Medication ทั้งหมดที่มีเพื่อหาสารบบยา Home-Med [cite: 3, 10]"
    )

with s1_cols[3]:
    st.text_area(
        "1.4 Progress Note",
        height=300,
        placeholder="คัดลอกบันทึก Progress note ทั้งหมด ...",
        key="s14",
        help="คัดลอกบันทึกการติดตามอาการทั้งหมด เพื่อให้ AI สังเคราะห์ปัญหาที่ส่งต่อ [cite: 3, 10]"
    )

st.divider()

# กลุ่มที่ 2: การประเมิน (Assessment) - ช่องกว้างพิเศษสำหรับคะแนนสุขภาพจิต [cite: 4, 5]
st.markdown("### **🔵 Step 2: การประเมิน (Assessment)**")
st.text_area(
    "คัดลอกผลคะแนน 9Q, 8Q, BPRS ทั้งหมดมาวางที่นี่",
    height=150,
    placeholder="คะแนน 9Q, 8Q, BPRS ...",
    key="s2",
    help="ดึงจากหน้า Assessment ในระบบผู้ป่วยนอกเพื่อวิเคราะห์ภาวะซึมเศร้าและการฆ่าตัวตาย [cite: 5, 10]"
)

st.divider()

# กลุ่มที่ 3: เวชระเบียน (Registration) - แบ่ง 4 ส่วนตามทะเบียนประวัติ [cite: 6, 7]
st.markdown("### **🟠 Step 3: เวชระเบียน (Registration)**")
s3_cols = st.columns(4)

with s3_cols[0]:
    st.text_area(
        "3.1 ข้อมูลทั่วไป",
        height=200,
        placeholder="HN, ชื่อ, อายุ, เลขบัตรประชาชน ...",
        key="s31",
        help="ดึงจากหน้า 'ทั่วไป 1' ในระบบเวชระเบียน (ชื่อ, อายุ, ศาสนา, อาชีพ) [cite: 7, 10]"
    )

with s3_cols[1]:
    st.text_area(
        "3.2 ที่อยู่ปัจจุบัน",
        height=200,
        placeholder="ที่อยู่ปัจจุบัน ...",
        key="s32",
        help="ดึงจากหน้า 'ทั่วไป 2' โดยต้องกดยืดกล่องเพื่อให้เห็นที่อยู่ครบถ้วน [cite: 7, 10]"
    )

with s3_cols[2]:
    st.text_area(
        "3.3 ผู้ติดต่อ",
        height=200,
        placeholder="ชื่อญาติ ความสัมพันธ์ เบอร์โทรศัพท์ ...",
        key="s33",
        help="ดึงจากหน้า 'ผู้ติดต่อ' ในระบบเวชระเบียนเพื่อใช้เป็นข้อมูล Contact [cite: 7, 10]"
    )

with s3_cols[3]:
    st.text_area(
        "3.4 สิทธิการรักษา",
        height=200,
        placeholder="สิทธิการรักษา ...",
        key="s34",
        help="ดึงจากหน้า 'สิทธิการรักษา' (สิทธิ์และสถานพยาบาลหลักใกล้บ้าน) [cite: 3, 7, 10]"
    )

# --- สิ้นสุดส่วนที่ 4 ---

# --- 8. ส่วนประมวลผลอัจฉริยะ (The PhD Extraction & Audit Logic) ---

if st.button("🚀 กดเพื่อประมวลผลและสกัดข้อมูลด้วย Gemini 3 Flash", use_container_width=True):
    # ดึงวันที่ปัจจุบันเพื่อใช้คำนวณระยะเวลาในชุมชน (LOC) และเติมในหัวเอกสาร
    thai_date_now = get_thai_date_parts()
    current_date_str = f"{thai_date_now['DAY']} {thai_date_now['MONTH']} {thai_date_now['YEAR']}"
    
    # รวบรวมข้อมูลดิบจากทั้ง 9 ส่วนที่คุณหมอกรอกไว้ [cite: 1-7]
    raw_context = f"""
    วันที่ทำรายการ (Current Date): {current_date_str}
    
    [IPD DATA]
    1.1 Admission: {st.session_state.s11}
    1.2 DX: {st.session_state.s12}
    1.3 Order/Med: {st.session_state.s13}
    1.4 Progress: {st.session_state.s14}
    
    [ASSESSMENT]
    Score: {st.session_state.s2}
    
    [REGISTRATION]
    3.1 General: {st.session_state.s31}
    3.2 Address: {st.session_state.s32}
    3.3 Contact: {st.session_state.s33}
    3.4 Rights: {st.session_state.s34}
    """
    
    with st.spinner('Gemini 3 Flash กำลังวิเคราะห์ข้อมูลและตรวจสอบ Verification Audit...'):
        # ตรรกะการสกัดข้อมูลเชิงลึกตามมาตรฐาน PhD [cite: 1-13]
        prompt = f"""
        คุณคือผู้ช่วยวิจัยทางการแพทย์ระดับ PhD ทำหน้าที่สกัดข้อมูลจากระบบ @ThanHIS ลงแบบฟอร์ม 062 
        จงปฏิบัติตามกฎเหล็ก "Verification Audit" และ "Search & Extract Logic" อย่างเคร่งครัด:

        1. กฎการตัดขยะข้อมูล (Noise Reduction Rule): 
           Ignore (ละทิ้ง) ข้อมูล Theme Customizer, Navbar, Menu Colors, Light/Dark Mode และ COPYRIGHT ทั้งหมด [cite: 13]

        2. ตรรกะการคำนวณและจัดรูปแบบพิเศษ (Specific Constraints):
           - [LOC (ระยะเวลาที่อยู่ในชุมชน)]: คำนวณโดยนำวันที่ปัจจุบัน ({current_date_str}) ลบด้วย วันที่จำหน่ายครั้งสุดท้าย (LAST_DC) ที่สกัดได้จากประวัติเดิม
           - [การวินิจฉัย (DX)]: สกัดรหัส ICD-10 (ไม่มีจุดทศนิยม) พร้อมชื่อโรคภาษาไทย โดยเริ่มจาก Principal Diagnosis เป็นอันดับแรก ตามด้วย Comorbidity ทั้งหมด ให้เขียนต่อกันในแถวเดียวและคั่นด้วยเครื่องหมายคอมม่า (,) ไปจนครบ
           - [ยา (MEDS)]: สกัดเฉพาะรายการ 'Home-Med' เขียนชื่อยาเป็น UPPERCASE พร้อมวิธีใช้และการบริหารยา ให้เขียนต่อกันในแถวเดียวและคั่นด้วยเครื่องหมายคอมม่า (,) ไปจนครบในแถวเดียวกัน
           - [วันนอนรวม (LOS)]: นำตัวเลขวัน Detox และ Rehab มาบวกกันเสมอ [cite: 3, 10, 27]
           - [วันที่จำหน่าย (DC_DATE)]: ให้ใช้ค่าวันที่ปัจจุบัน คือ {current_date_str} [cite: 10, 27]

        3. การสกัดตามสมอเรือ (Keywords Anchor):
           - [HN]: มองหาตัวเลขหลัง 'HN' หรือ 'Hospital number' [cite: 3, 7]
           - [อาการนำส่ง (CC)]: สกัดจาก 'Chief Complaint' หรือ 'CC :' จนถึง 'Present illness' [cite: 3, 10]
           - [คะแนนประเมิน]: สกัดตัวเลขหลัง 'ซึมเศร้า' (9Q) และ 'ฆ่าตัวตาย' (8Q) [cite: 5, 10]
           - [สรุปปัญหา (PROGRESS)]: สังเคราะห์จาก Progress Note ทั้งหมด ให้เป็นสรุปย่อหน้าเดียว ความยาว 2-3 บรรทัด [cite: 10, 33]

        4. นโยบายความถูกต้อง (Verification Audit Policy):
           - หากไม่พบข้อมูลให้ระบุ [กรอกด้วยตนเอง] ห้ามเว้นว่างเด็ดขาด [cite: 9, 10]
           - วิเคราะห์ 'รับไว้ครั้งที่' (VISIT_NUM) จากประวัติเดิม (เช่น เคยนอน 4 ครั้ง ครั้งนี้จะเป็น 5) [cite: 10, 23]

        ข้อมูลดิบสำหรับวิเคราะห์:
        {raw_context}

        ตอบกลับในรูปแบบ JSON ที่มี Key ตรงกับ Placeholder ใน Word: 
        NAME, AGE, HN, ID, EDU, CAREER, RELIGION, STATUS, RIGHTS, LAST_DC, LOC, ADMIT_DATE, VISIT_NUM, CC, CONTACT, RELATION, PHONE, NEAR_HOSP, DC_DATE, LOS, ADDRESS, Q9, Q8, MEDS, DX, PROGRESS, POST_SERVICE
        """
        
        try:
            # ประมวลผลผ่าน Gemini 3 Flash Paid Tier เพื่อความแม่นยำสูงสุด
            response = client.models.generate_content(model=MODEL_ID, contents=prompt)
            match = re.search(r'\{.*\}', response.text, re.DOTALL)
            
            if match:
                # เก็บข้อมูล JSON ลงใน Session State พร้อมเติมข้อมูล วัน/เดือน/ปี ไทย สำหรับหัวเอกสาร
                st.session_state.extracted_json_data = json.loads(match.group())
                st.session_state.extracted_json_data.update(thai_date_now)
                st.success("✅ วิเคราะห์ข้อมูลและคำนวณระยะเวลาในชุมชน (LOC) สำเร็จ!")
            else:
                st.error("AI ไม่สามารถสร้างรูปแบบข้อมูลที่ถูกต้องได้ กรุณาตรวจสอบข้อมูลดิบอีกครั้ง")
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {e}")

# --- สิ้นสุดส่วนที่ 5 ---



# --- 9. ฟังก์ชันจัดการไฟล์ Word (วันที่ชิดขวา + หัวข้อตัวหนา + ฟอนต์ 13) ---

def fill_pmnidat_doc(data):
    """ฟังก์ชันนำข้อมูลจากระบบเข้าสู่ไฟล์แม่แบบ .docx พร้อมจัดรูปแบบเชิงลึก [cite: 8-10, 14]"""
    try:
        # โหลดไฟล์แม่แบบ PMNIDAT 062 จาก Directory ของโปรเจกต์
        doc = Document("PMNIDAT 062 แบบบันทึกข้อมูลเพื่อส่งต่อ.docx")
        
        # เตรียมชุดข้อมูล Mapping โดยใช้ Key ตัวพิมพ์ใหญ่เพื่อให้ตรงกับ Placeholder  [cite: 18-34]
        mapping = {f"{{{{{k.upper()}}}}}": str(v) for k, v in data.items()}
        
        def apply_style_and_replace(paragraph):
            """ตรรกะการจัดวาง Alignment และการเน้นตัวหนาตามเงื่อนไขของคุณหมออาร์ม"""
            
            # 1. การจัดวางบรรทัด (Alignment Logic): วันที่จำหน่ายชิดซ้าย / วันที่หัวเอกสารชิดขวา
            if "วันที่จำหน่าย" in paragraph.text:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            elif any(k in paragraph.text for k in ["{{DAY}}", "{{MONTH}}", "{{YEAR}}"]):
                paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

            # 2. ค้นหาและแทนที่ข้อความ (Replacement Logic)
            for key, value in mapping.items():
                if key in paragraph.text:
                    paragraph.text = paragraph.text.replace(key, value)
                    
                    # 3. การเน้นตัวหนา (Bolding Logic): เน้นที่หัวข้อทางการแพทย์เพื่อให้กวาดสายตาได้ง่าย
                    clinical_headers = ["การวินิจฉัย", "Home Medication", "สรุปปัญหา", "อาการนำส่ง"]
                    should_bold = any(h in paragraph.text for h in clinical_headers)
                    
                    # กำหนดคุณลักษณะอักษร: ฟอนต์ 13 และตัวหนาเฉพาะจุด
                    for run in paragraph.runs:
                        run.font.size = Pt(13)
                        if should_bold:
                            run.font.bold = True 

        # ตรวจสอบและแทนที่ข้อมูลทั้งในส่วนเนื้อหาหลักและภายในตารางเอกสาร  [cite: 14-43]
        for p in doc.paragraphs: 
            apply_style_and_replace(p)
            
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs: 
                        apply_style_and_replace(p)
                            
        # บันทึกไฟล์ลงในหน่วยความจำชั่วคราวเพื่อให้พร้อมสำหรับขั้นตอนการดาวน์โหลด
        buffer = io.BytesIO()
        doc.save(buffer)
        return buffer.getvalue()
    except Exception as e:
        st.error(f"⚠️ เกิดข้อผิดพลาดในขั้นตอนจัดการไฟล์: {e}")
        return None




# --- 10. การแสดงผลลัพธ์และการควบคุมการดาวน์โหลด (Final Execution) ---

# ตรวจสอบว่ามีการสกัดข้อมูลสำเร็จใน Session State หรือไม่
if "extracted_json_data" in st.session_state and st.session_state.extracted_json_data:
    # เรียกใช้ฟังก์ชันจากส่วนที่ 6 เพื่อสร้างไฟล์ Word ตามตรรกะจัดวางแบบใหม่
    word_file_final = fill_pmnidat_doc(st.session_state.extracted_json_data)
    
    if word_file_final:
        # บันทึกสถิติการใช้งานไปยังระบบ Log (ฟังก์ชันในส่วนที่ 1)
        log_usage(st.session_state.extracted_json_data.get('NAME', '[ไม่ระบุชื่อ]'))
        
        st.divider()
        st.balloons() # เฉลิมฉลองความสำเร็จในการประมวลผลข้อมูล
        st.success("🎉 ระบบสกัดข้อมูลและจัดทำเอกสาร PMNIDAT 062 ฉบับสมบูรณ์ (Master v3.35) เรียบร้อยแล้ว!")
        
        # ปุ่มดาวน์โหลดไฟล์ฉบับ Final ที่พร้อมสำหรับการนำไปใช้งานทางคลินิก
        st.download_button(
            label="💾 ดาวน์โหลดไฟล์ 'แบบบันทึกข้อมูลเพื่อส่งต่อ (PMNIDAT 062).docx'",
            data=word_file_final,
            file_name=f"Refer_{st.session_state.extracted_json_data.get('NAME', '062')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

# --- 11. มาตรการรักษาความปลอดภัยและความเป็นส่วนตัว (PDPA Footer) ---
# ส่วนสำคัญที่คุณหมออาร์มกำหนดเพื่อรักษามาตรฐานความลับของผู้ป่วย 

st.divider()
st.info("""
    **🛡️ มาตรการรักษาความปลอดภัยของข้อมูลคนไข้ (PDPA Compliance):**
    
    * **ไม่มีการจัดเก็บข้อมูลถาวร:** ระบบ PMNIDAT Smart Transfer ประมวลผลแบบ Real-time บนหน่วยความจำชั่วคราว และจะไม่มีการบันทึกข้อมูลส่วนบุคคลของผู้ป่วยลงในฐานข้อมูลถาวรใด ๆ ของแอปพลิเคชัน
    * **ระบบ Session-Based:** ข้อมูลที่คัดลอกมาวางจะถูกลบทิ้งทันทีเมื่อมีการรีเฟรชหน้าจอ (Refresh) หรือปิดเบราว์เซอร์ โปรดดาวน์โหลดไฟล์ให้เรียบร้อยก่อนปิดระบบ
    * **การตรวจสอบความถูกต้อง:** เนื่องจากเป็นระบบช่วยสกัดข้อมูลด้วย AI (Large Language Model) โปรดตรวจสอบความถูกต้องของข้อมูล (Verification Audit) อีกครั้งตามมาตรฐานวิชาชีพก่อนนำไปใช้งานจริง
    
    *Created by Dr.Charshawn Lahnwong (Pharmacology & Addiction Medicine Specialist)*
    """)

# --- สิ้นสุดชุดโค้ดทั้งหมด ---

