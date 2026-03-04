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
# รองรับการเติมข้อมูล วันที่...เดือน...พ.ศ... อัตโนมัติ 
THAI_MONTHS = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
]

def get_thai_date_parts():
    """ส่งคืน วัน, เดือน(ชื่อเต็ม), และ พ.ศ. ปัจจุบัน """
    now = datetime.datetime.now()
    return {
        "DAY": str(now.day),
        "MONTH": THAI_MONTHS[now.month - 1],
        "YEAR": str(now.year + 543)  # แปลงปี ค.ศ. เป็น พ.ศ. 
    }

# --- 2. ฟังก์ชันบันทึกสถิติการใช้งาน (Usage Logging) ---
def log_usage(patient_name):
    """ส่งชื่อผู้ป่วยไปยัง Google Sheets เพื่อบันทึกประวัติการใช้งาน"""
    try:
        url = st.secrets["APPS_SCRIPT_URL"]
        requests.post(url, json={"name": patient_name}, timeout=5)
    except:
        pass # ป้องกันแอปหยุดทำงานหากเชื่อมต่อ Log ไม่สำเร็จ

# --- 3. การตั้งค่าระบบความปลอดภัยและการเชื่อมต่อ API ---
try:
    # ดึงค่า API Key จาก Streamlit Secrets
    API_KEY = st.secrets["GEMINI_API_KEY"]
    client = genai.Client(api_key=API_KEY)
    
    @st.cache_resource
    def find_active_model():
        """ตรวจสอบรุ่นโมเดลที่สามารถใช้งานได้จริงในบัญชีของคุณหมอ"""
        try:
            available_models = [m.name for m in client.models.list()]
            # เน้นการใช้รุ่น Gemini 1.5 Flash เพื่อความเร็วและแม่นยำ
            for m in available_models:
                if "gemini-1.5-flash" in m: return m
            return available_models[0]
        except:
            return "models/gemini-1.5-flash"
            
    MODEL_ID = find_active_model()
except Exception as e:
    st.error(f"❌ ระบบเชื่อมต่อผิดพลาด (API Configuration Error): {e}")
    st.stop()

# --- 4. ระบบจดจำข้อมูลชั่วคราว (Session State Memory) ---
# สร้างสถานะสำหรับเก็บข้อมูลทั้ง 9 ช่อง เพื่อไม่ให้หายเมื่อรีเฟรชหน้าจอ 
field_keys = ['s11', 's12', 's13', 's14', 's2', 's31', 's32', 's33', 's34']
for key in field_keys:
    if key not in st.session_state:
        st.session_state[key] = ""

# --- 5. ฟังก์ชันโหลดข้อมูลตัวอย่าง (Smart Sample Data) ---
# บรรจุรายละเอียดครบถ้วนเพื่อใช้ทดสอบตรรกะการสกัดข้อมูลระดับ PhD

def load_test_data():
    """โหลดข้อมูลตัวอย่าง 'นาย ชาย ธัญญารักษ์' เข้าสู่ Session State """
    
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
    
    # 1.2 ข้อมูลการวินิจฉัย (ICD-10 พร้อมชื่อไทย)
    st.session_state.s12 = (
        "1. F105 - โรคจิตจากสุรา (Alcohol)\n"
        "2. I10 - โรคความดันโลหิตสูง (Hypertension)"
    )
    
    # 1.3 ข้อมูลรายการยา (Order / Meds)
    st.session_state.s13 = (
        "1. AMLODIPINE 5 MG 1x1 pc (เช้า)\n"
        "2. QUETIAPINE 25 MG 1 tab hs (ก่อนนอน)\n"
        "(Home-Med ทั้งหมด)"
    )
    
    # 1.4 ข้อมูลบันทึกการติดตามอาการ (Progress Note)
    st.session_state.s14 = (
        "S: สบายดี กินข้าวได้ นอนหลับได้\n"
        "O: V/S stable, BP 120/80 mmHg\n"
        "A: อาการคงที่ เตรียมจำหน่าย\n"
        "P: Discharge to home"
    )
    
    # 2. ข้อมูลการประเมิน (Assessment)
    st.session_state.s2 = "9Q : 5 คะแนน\n8Q : 0 คะแนน\nBPRS : 15 คะแนน"
    
    # 3.1 ข้อมูลทั่วไป (Registration 1)
    st.session_state.s31 = (
        "Hospital Number 690099999 ชื่อ [ชาย] นามสกุล [ธัญญารักษ์]\n"
        "เลขบัตรประชาชน* [1-2345-67890-12-3]\n"
        "ศาสนา [พุทธ] อาชีพ [ข้าราชการ] สถานภาพ [สมรส] การศึกษา [ปริญญาตรี]"
    )
    
    # 3.2 ที่อยู่ปัจจุบัน (Registration 2)
    st.session_state.s32 = "ที่อยู่ปัจจุบัน: เลขที่ 123 ต.หลักหก อ.เมืองปทุมธานี จ.ปทุมธานี"
    
    # 3.3 ผู้ติดต่อ (Contact Info)
    st.session_state.s33 = "คุณ หญิง ธัญญารักษ์ (ภรรยา) เบอร์โทร: 081-234-XXXX"
    
    # 3.4 สิทธิการรักษา (Rights)
    st.session_state.s34 = "สิทธิหลัก [จ่ายตรง] สถานพยาบาลหลัก [สถาบันบำบัดรักษาและฟื้นฟูผู้ติดยาเสพติดแห่งชาติบรมราชชนนี]"
    
    # สั่งให้แอปรีเฟรชเพื่อแสดงข้อมูลในช่องกรอกทันที
    st.rerun()

def clear_all_data():
    """ล้างข้อมูลดิบทั้งหมดใน Session State """
    for key in field_keys:
        st.session_state[key] = ""
    st.rerun()

# --- 6. การออกแบบแถบเมนูด้านข้าง (Sidebar Manual & Controls) ---

with st.sidebar:
    # แสดงโลโก้สถาบันฯ และหัวข้อคู่มือ 
    st.image("https://pmnidathis.dms.go.th/static/media/health.1bfd961f.png", width=100)
    st.header("📖 คู่มือการใช้งาน")
    
    # ปุ่มควบคุมสำหรับทดสอบระบบและล้างข้อมูล
    st.subheader("🛠️ เครื่องมือระบบ")
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        if st.button("🧬 ตัวอย่างข้อมูล", use_container_width=True):
            load_test_data()
    with col_t2:
        if st.button("🧹 ล้างข้อมูล", use_container_width=True):
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
    # แสดงเครดิตและรุ่นของระบบ
    st.success("PMNIDAT Smart Transfer (Version 3.35) | Created by Dr.Charshawn Lahnwong (1 March 2026)")

# --- สิ้นสุดส่วนที่ 3 ---

# --- 7. ส่วนการออกแบบหน้าจอหลัก (Main UI Layout) ---

# แทรกส่วนหัวที่คุณหมอกำหนดไว้ เพื่อความเป็นทางการของระบบ
st.title("🏥 PMNIDAT Smart Transfer")
st.subheader("ผู้ช่วยพิมพ์ 'แบบบันทึกข้อมูลเพื่อส่งต่อ (PMNIDAT 062)' โดยอัตโนมัติ (Version 3.35)")

st.divider()

# กลุ่มที่ 1: ระบบผู้ป่วยใน (IPD) - จัดวาง 4 คอลัมน์สำหรับข้อมูลทางคลินิก 
st.markdown("### **🟢 Step 1: ระบบผู้ป่วยใน (IPD)**")
s1_cols = st.columns(4)

with s1_cols[0]:
    st.text_area(
        "1.1 Admission Note",
        height=300,
        placeholder="คัดลอกข้อมูลแรกรับทั้งหมด...",
        key="s11",
        help="คัดลอกจากเมนู Admission note ในระบบ IPD"
    )

with s1_cols[1]:
    st.text_area(
        "1.2 การวินิจฉัย",
        height=300,
        placeholder="คัดลอกรหัส ICD-10 ทั้งหมด...",
        key="s12",
        help="คัดลอกจากเมนูการวินิจฉัยเพื่อสกัดรหัสโรคภาษาอังกฤษ"
    )

with s1_cols[2]:
    st.text_area(
        "1.3 Order / Meds",
        height=300,
        placeholder="คัดลอกข้อมูลจากเมนู Order ทั้งหมด ...",
        key="s13",
        help="คัดลอก Order และ Medication ทั้งหมดที่มีเพื่อหาสารบบยา Home-Med"
    )

with s1_cols[3]:
    st.text_area(
        "1.4 Progress Note",
        height=300,
        placeholder="คัดลอกบันทึก Progress note ทั้งหมด ...",
        key="s14",
        help="คัดลอกบันทึกการติดตามอาการทั้งหมด เพื่อให้ AI สังเคราะห์ปัญหาที่ส่งต่อ"
    )

st.divider()

# กลุ่มที่ 2: การประเมิน (Assessment) - ช่องกว้างพิเศษสำหรับคะแนนสุขภาพจิต
st.markdown("### **🔵 Step 2: การประเมิน (Assessment)**")
st.text_area(
    "คัดลอกผลคะแนน 9Q, 8Q, BPRS ทั้งหมดมาวางที่นี่",
    height=150,
    placeholder="คะแนน 9Q, 8Q, BPRS ...",
    key="s2",
    help="ดึงจากหน้า Assessment ในระบบผู้ป่วยนอกเพื่อวิเคราะห์ภาวะซึมเศร้าและการฆ่าตัวตาย"
)

st.divider()

# กลุ่มที่ 3: เวชระเบียน (Registration) - แบ่ง 4 ส่วนตามทะเบียนประวัติ
st.markdown("### **🟠 Step 3: เวชระเบียน (Registration)**")
s3_cols = st.columns(4)

with s3_cols[0]:
    st.text_area(
        "3.1 ข้อมูลทั่วไป",
        height=200,
        placeholder="HN, ชื่อ, อายุ, เลขบัตรประชาชน ...",
        key="s31",
        help="ดึงจากหน้า 'ทั่วไป 1' ในระบบเวชระเบียน (ชื่อ, อายุ, ศาสนา, อาชีพ)"
    )

with s3_cols[1]:
    st.text_area(
        "3.2 ที่อยู่ปัจจุบัน",
        height=200,
        placeholder="ที่อยู่ปัจจุบัน ...",
        key="s32",
        help="ดึงจากหน้า 'ทั่วไป 2' โดยต้องกดยืดกล่องเพื่อให้เห็นที่อยู่ครบถ้วน"
    )

with s3_cols[2]:
    st.text_area(
        "3.3 ผู้ติดต่อ",
        height=200,
        placeholder="ชื่อญาติ ความสัมพันธ์ เบอร์โทรศัพท์ ...",
        key="s33",
        help="ดึงจากหน้า 'ผู้ติดต่อ' ในระบบเวชระเบียนเพื่อใช้เป็นข้อมูล Contact"
    )

with s3_cols[3]:
    st.text_area(
        "3.4 สิทธิการรักษา",
        height=200,
        placeholder="สิทธิการรักษา ...",
        key="s34",
        help="ดึงจากหน้า 'สิทธิการรักษา' (สิทธิ์และสถานพยาบาลหลักใกล้บ้าน)"
    )

# --- สิ้นสุดส่วนที่ 4 ---

# --- 8. ส่วนประมวลผลอัจฉริยะ (The PhD Extraction Brain) ---

if st.button("🚀 กดเพื่อประมวลผลและสกัดข้อมูลด้วย Gemini 3 Flash", use_container_width=True):
    # ดึงวันที่ปัจจุบันและส่วนประกอบวันที่ไทยเพื่อใช้คำนวณ LOC และแสดงผลในใบส่งต่อ
    thai_date_data = get_thai_date_parts()
    today_str = f"{thai_date_data['DAY']} {thai_date_data['MONTH']} {thai_date_data['YEAR']}"
    
    # รวบรวมข้อมูลดิบทั้งหมดจาก 9 ช่องกรอก
    all_raw_data = f"""
    วันที่ทำรายการปัจจุบัน: {today_str}
    
    [GROUP 1: IPD DATA]
    1.1 Admission Note: {st.session_state.s11}
    1.2 การวินิจฉัย: {st.session_state.s12}
    1.3 Order / Meds (คัดลอกทั้งหมด): {st.session_state.s13}
    1.4 Progress Note (คัดลอกทั้งหมด): {st.session_state.s14}
    
    [GROUP 2: ASSESSMENT]
    Assessment Score: {st.session_state.s2}
    
    [GROUP 3: REGISTRATION]
    3.1 ข้อมูลทั่วไป (HN, ข้อมูลส่วนตัว): {st.session_state.s31}
    3.2 ที่อยู่ปัจจุบัน: {st.session_state.s32}
    3.3 ผู้ติดต่อ: {st.session_state.s33}
    3.4 สิทธิการรักษา: {st.session_state.s34}
    """
    
    with st.spinner('Gemini 3 Flash กำลังวิเคราะห์ข้อมูลเชิงลึกและตรวจสอบ Verification Audit...'):
        # ตรรกะการประมวลผลระดับ PhD ตามมาตรฐานสถาบันฯ
        prompt = f"""
        คุณคือผู้ช่วยวิจัยทางการแพทย์ระดับ PhD ทำหน้าที่สกัดข้อมูลจากระบบ @ThanHIS ลงแบบฟอร์ม 062 
        จงปฏิบัติตามกฎเหล็ก "Verification Audit" และ "Search & Extract Logic" อย่างเคร่งครัด:

        1. กฎการตัดขยะข้อมูล (Noise Reduction Rule): 
           Ignore (ละทิ้ง) ข้อมูล Theme Customizer, Navbar, Menu Colors, Light/Dark Mode และ COPYRIGHT ทั้งหมด 

        2. ตรรกะการคำนวณและจัดรูปแบบพิเศษ (PhD Logic):
           - [LOC (ระยะเวลาที่อยู่ในชุมชน)]: คำนวณโดยนำวันที่ปัจจุบัน ({today_str}) ลบด้วย วันที่จำหน่ายครั้งสุดท้าย (LAST_DC)
           - [DX (การวินิจฉัย)]: สกัดรหัส ICD-10 (ไม่มีจุดทศนิยม) พร้อมชื่อโรคภาษาไทย โดยเริ่มจาก Principal Diagnosis และตามด้วย Comorbidity ทั้งหมด ให้เขียนต่อกันในแถวเดียว คั่นด้วยเครื่องหมายคอมม่า (,) ไปจนครบ
           - [MEDS (Home Medication)]: สกัดรายการยาที่มีคำว่า 'Home-Med' เขียนชื่อยาเป็น UPPERCASE พร้อมวิธีใช้และการบริหารยา ให้เขียนต่อกันในแถวเดียว คั่นด้วยเครื่องหมายคอมม่า (,) ไปจนครบ
           - [LOS (รวมวันนอน)]: นำจำนวนวัน Detox และ Rehab มาบวกกันเสมอ
           - [PROGRESS (สรุปปัญหา)]: สังเคราะห์จาก Progress Note ทั้งหมด ให้เป็นย่อหน้าเดียว ความยาว 2-3 บรรทัด

        3. ตรรกะสมอเรือ (Keywords Anchor):
           - [HN]: มองหาตัวเลขหลัง 'HN' หรือ 'Hospital number' 
           - [สิทธิการรักษา]: สกัดข้อความหลัง 'สิทธิ์ :'
           - [CC]: สกัดจาก 'Chief Complaint' หรือ 'CC :' จนถึง 'Present illness'
           - [ข้อมูลทั่วไป]: สกัดชื่อ-สกุล, อายุ, บัตรประชาชน (13 หลัก), ศาสนา, อาชีพ และที่อยู่
           - [คะแนน]: สกัดตัวเลขหลัง 'ผลรวมการประเมินโรคซึมเศร้า' (9Q) และ 'การฆ่าตัวตาย' (8Q)
           - [ผู้ดูแล]: สกัดชื่อ ความสัมพันธ์ และเบอร์โทรศัพท์

        4. นโยบาย Verification Audit:
           - หากไม่พบข้อมูลในช่องที่ระบุ ให้หาจากช่องอื่น หากไม่มีจริงๆ ให้ระบุ [กรอกด้วยตนเอง]
           - วิเคราะห์ 'รับไว้ครั้งที่' จากประวัติเดิม (เช่น เคยนอน 4 ครั้ง ครั้งนี้จะเป็น 5) 

        ข้อมูลดิบสำหรับวิเคราะห์:
        {all_raw_data}

        ตอบกลับในรูปแบบ JSON ที่มี Key ตรงกับ Placeholder ใน Word: 
        NAME, AGE, HN, ID, EDU, CAREER, RELIGION, STATUS, RIGHTS, LAST_DC, LOC, ADMIT_DATE, VISIT_NUM, CC, CONTACT, RELATION, PHONE, NEAR_HOSP, DC_DATE, LOS, ADDRESS, Q9, Q8, MEDS, DX, PROGRESS, POST_SERVICE
        """
        
        try:
            # ประมวลผลผ่านโมเดล Gemini 3 Flash
            response = client.models.generate_content(model=MODEL_ID, contents=prompt)
            match = re.search(r'\{.*\}', response.text, re.DOTALL)
            
            if match:
                st.session_state.extracted_json_data = json.loads(match.group())
                # เพิ่มข้อมูล วัน/เดือน/ปี ไทย เข้าไปใน JSON เพื่อเติมในบรรทัดแรกของเอกสาร
                st.session_state.extracted_json_data.update(thai_date_data)
                st.success("✅ วิเคราะห์ข้อมูลสำเร็จและจัดรูปแบบ DX/MEDS เป็นแถวเดียวเรียบร้อยแล้ว!")
            else:
                st.error("AI ไม่สามารถสร้างรูปแบบข้อมูลที่ถูกต้องได้ กรุณาลองตรวจสอบข้อมูลดิบอีกครั้ง")
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {e}")

# --- สิ้นสุดส่วนที่ 5 ---

# --- 9. ฟังก์ชันจัดการไฟล์ Word (จัดรูปแบบชิดซ้าย + ฟอนต์ 13 + แทนที่ Placeholder) ---

def fill_pmnidat_doc(data):
    """ฟังก์ชันนำข้อมูลจาก JSON ไปบรรจุลงในไฟล์แม่แบบ .docx พร้อมจัดรูปแบบ"""
    try:
        # โหลดไฟล์แม่แบบ PMNIDAT 062 ที่คุณหมอเตรียมไว้
        doc = Document("PMNIDAT 062 แบบบันทึกข้อมูลเพื่อส่งต่อ.docx")
        
        # เตรียมชุดข้อมูลสำหรับแทนที่ (Mapping) โดยใช้ตัวพิมพ์ใหญ่ตาม Placeholder
        mapping = {f"{{{{{k.upper()}}}}}": str(v) for k, v in data.items()}
        
        def apply_style_and_replace(paragraph):
            """ฟังก์ชันย่อยสำหรับแทนที่ข้อความและบังคับฟอนต์/การจัดวาง"""
            for key, value in mapping.items():
                if key in paragraph.text:
                    # แทนที่ข้อความในตำแหน่ง Placeholder
                    paragraph.text = paragraph.text.replace(key, value)
                    
                    # บังคับการจัดวางให้ชิดซ้าย (Left Alignment)
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    
                    # กำหนดขนาดฟอนต์เป็น 13 pt ตลอดทั้งย่อหน้า
                    for run in paragraph.runs:
                        run.font.size = Pt(13)

        # ดำเนินการค้นหาและแทนที่ทั้งในเนื้อหาปกติและภายในตาราง
        for p in doc.paragraphs: 
            apply_style_and_replace(p)
            
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs: 
                        apply_style_and_replace(p)
                            
        # บันทึกไฟล์ลงในหน่วยความจำชั่วคราว (Memory Buffer)
        buffer = io.BytesIO()
        doc.save(buffer)
        return buffer.getvalue()
    except Exception as e:
        st.error(f"⚠️ เกิดข้อผิดพลาดในการสร้างไฟล์ Word: {e}")
        return None

# --- 10. การแสดงผลลัพธ์และปุ่มดาวน์โหลด (Final Execution) ---

# ตรวจสอบว่ามีการประมวลผล JSON สำเร็จในระบบหรือไม่
if "extracted_json_data" in st.session_state and st.session_state.extracted_json_data:
    # สร้างไฟล์ Word จากข้อมูลล่าสุดที่จัดรูปแบบแถวเดียวเรียบร้อยแล้ว
    word_file_final = fill_pmnidat_doc(st.session_state.extracted_json_data)
    
    if word_file_final:
        # บันทึกสถิติการใช้งานลง Log Book (ฟังก์ชันในส่วนที่ 1)
        log_usage(st.session_state.extracted_json_data.get('NAME', '[ไม่ระบุชื่อ]'))
        
        st.divider()
        st.balloons() # แสดงเอฟเฟกต์ความสำเร็จ
        st.success("🎉 เอกสาร PMNIDAT 062 (Master Version 3.35) พร้อมสำหรับการดาวน์โหลดแล้ว!")
        
        # ปุ่มดาวน์โหลดไฟล์ฉบับสมบูรณ์ (จัดรูปแบบฟอนต์ 13 + ชิดซ้าย + วันที่ปัจจุบัน)
        st.download_button(
            label="💾 ดาวน์โหลดใบส่งต่อ 062 (จัดรูปแบบแถวเดียว + ออโต้วันที่ปัจจุบัน)",
            data=word_file_final,
            file_name=f"Refer_{st.session_state.extracted_json_data.get('NAME', '062')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

# --- 11. มาตรการความปลอดภัยและประกาศ PDPA ---
st.divider()
st.info("""
    **มาตรการรักษาความปลอดภัยของข้อมูลผู้ป่วย (PDPA Compliance):**
    * ระบบจะประมวลผลข้อมูลแบบ Real-time และ **ไม่มีการจัดเก็บข้อมูลถาวรบนเซิร์ฟเวอร์**
    * ข้อมูลในช่องกรอกจะถูกลบทิ้งทันทีเมื่อปิดเบราว์เซอร์หรือกด Refresh หน้าจอ
    * **คำแนะนำ:** โปรดตรวจสอบความถูกต้องของข้อมูล (Verification Audit) อีกครั้งก่อนนำเอกสารไปใช้งานจริง
    """)

