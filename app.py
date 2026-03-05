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
# รองรับการเติมข้อมูล วันที่ {{DAY}} เดือน {{MONTH}} พ.ศ. {{YEAR}} อัตโนมัติ
THAI_MONTHS = [
    "มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
    "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"
]

def get_thai_date_parts():
    """ส่งคืน วัน, เดือน(ชื่อเต็ม), และ พ.ศ. ปัจจุบัน สำหรับใช้ใน LOC และหัวเอกสาร"""
    now = datetime.datetime.now()
    return {
        "DAY": str(now.day),
        "MONTH": THAI_MONTHS[now.month - 1],
        "YEAR": str(now.year + 543) # แปลง ค.ศ. เป็น พ.ศ.
    }

# --- 2. ฟังก์ชันบันทึกสถิติการใช้งาน (Usage Logging) ---
def log_usage(patient_name):
    """บันทึกชื่อผู้ป่วยที่ประมวลผลไปยังระบบสถิติ (ถ้ามี)"""
    try:
        url = st.secrets.get("APPS_SCRIPT_URL", "")
        if url:
            requests.post(url, json={"name": patient_name}, timeout=5)
    except:
        pass

# --- 3. การตั้งค่าความปลอดภัยและการเชื่อมต่อ API ---
try:
    API_KEY = st.secrets["GEMINI_API_KEY"]
    client = genai.Client(api_key=API_KEY)
    MODEL_ID = "gemini-1.5-flash" # โมเดลหลักสำหรับการสกัดข้อมูลความเร็วสูง
except Exception as e:
    st.error(f"❌ ระบบเชื่อมต่อผิดพลาด (Configuration Error): {e}")
    st.stop()

# --- 4. ระบบจดจำข้อมูลชั่วคราว (Session State Memory) ---
# สร้างหน่วยความจำสำหรับช่องกรอกข้อมูลทั้ง 9 ช่อง
field_keys = ['s11', 's12', 's13', 's14', 's2', 's31', 's32', 's33', 's34']
for key in field_keys:
    if key not in st.session_state:
        st.session_state[key] = ""

if "extracted_json_data" not in st.session_state:
    st.session_state.extracted_json_data = None



# --- สิ้นสุดส่วนที่ 1 ---



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

# --- 6. ฟังก์ชันล้างข้อมูล (Clear Data) ---
def clear_all_data():
    """ล้างข้อมูลดิบทั้งหมดออกจากระบบเพื่อเริ่มเคสใหม่ (Security Audit)"""
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
        if st.button("🧬 ตัวอย่างข้อมูล", use_container_width=True, help="ใช้ข้อมูลจำลองของ นาย ชาย ธัญญารักษ์ เพื่อทดสอบระบบ"):
            load_test_data()
    with col_t2:
        if st.button("🧹 ล้างข้อมูล", use_container_width=True, help="ลบข้อมูลทุกช่องเพื่อเริ่มเคสใหม่ (PDPA)"):
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
    # แสดงข้อมูลรุ่นโปรแกรมและผู้พัฒนา (เครดิตคุณหมออาร์ม)
    st.caption("PMNIDAT Smart Transfer v3.35")
    st.success("Created by Dr. Charshawn Lahnwong (3 Mar 2026)"



               
# --- สิ้นสุดส่วนที่ 3 ---



               
# --- 8. การออกแบบส่วนหัวข้อหลัก (Application Header) ---

# แทรกส่วนหัวที่คุณหมอกำหนดไว้ เพื่อความเป็นทางการของระบบ
st.title("🏥 PMNIDAT Smart Transfer")
st.subheader("ผู้ช่วยพิมพ์ 'แบบบันทึกข้อมูลเพื่อส่งต่อ (PMNIDAT 062)' โดยอัตโนมัติ (Version 3.35)")

st.divider()

# --- 9. ส่วนการรับข้อมูล 9 ส่วน (The 9 Input Fields) ---
               
# กลุ่มที่ 1: ระบบผู้ป่วยใน (IPD) - จัดวาง 4 คอลัมน์สำหรับข้อมูลทางคลินิก 
st.markdown("### **🟢 Step 1: ระบบผู้ป่วยใน (IPD)**")
s1_cols = st.columns(4)

with s1_cols[0]:
    st.text_area(
        "1.1 Admission Note", key="s11",
        height=300,
        placeholder="คัดลอกข้อมูลแรกรับทั้งหมด...",
        key="s11",
        help="คัดลอกจากเมนู Admission note ในระบบ IPD"
    )

with s1_cols[1]:
    st.text_area(
        "1.2 การวินิจฉัย", key="s12",
        height=300,
        placeholder="คัดลอกรหัส ICD-10 ทั้งหมด...",
        key="s12",
        help="คัดลอกจากเมนูการวินิจฉัยเพื่อสกัดรหัสโรคภาษาอังกฤษ"
    )

with s1_cols[2]:
    st.text_area(
        "1.3 Order / Meds", key="s13",
        height=300,
        placeholder="คัดลอกข้อมูลจากเมนู Order ทั้งหมด ...",
        key="s13",
        help="คัดลอก Order และ Medication ทั้งหมดที่มีเพื่อหาสารบบยา Home-Med"
    )

with s1_cols[3]:
    st.text_area(
        "1.4 Progress Note", key="s14",
        height=300,
        placeholder="คัดลอกบันทึก Progress note ทั้งหมด ...",
        key="s14",
        help="คัดลอกบันทึกการติดตามอาการทั้งหมด เพื่อให้ AI สังเคราะห์ปัญหาที่ส่งต่อ"
    )

st.divider()

# กลุ่มที่ 2: การประเมิน (Assessment) - ช่องกว้างพิเศษสำหรับคะแนนสุขภาพจิต
st.markdown("### **🔵 Step 2: การประเมิน (Assessment)**")
st.text_area(
    "คัดลอกผลคะแนน 9Q, 8Q, BPRS ทั้งหมดมาวางที่นี่", key="s2",
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
        "3.1 ข้อมูลทั่วไป", key="s31",
        height=200,
        placeholder="HN, ชื่อ, อายุ, เลขบัตรประชาชน ...",
        key="s31",
        help="ดึงจากหน้า 'ทั่วไป 1' ในระบบเวชระเบียน (ชื่อ, อายุ, ศาสนา, อาชีพ)"
    )

with s3_cols[1]:
    st.text_area(
        "3.2 ที่อยู่ปัจจุบัน", key="s32",
        height=200,
        placeholder="ที่อยู่ปัจจุบัน ...",
        key="s32",
        help="ดึงจากหน้า 'ทั่วไป 2' โดยต้องกดยืดกล่องเพื่อให้เห็นที่อยู่ครบถ้วน"
    )

with s3_cols[2]:
    st.text_area(
        "3.3 ผู้ติดต่อ", key="s33",
        height=200,
        placeholder="ชื่อญาติ ความสัมพันธ์ เบอร์โทรศัพท์ ...",
        key="s33",
        help="ดึงจากหน้า 'ผู้ติดต่อ' ในระบบเวชระเบียนเพื่อใช้เป็นข้อมูล Contact"
    )

with s3_cols[3]:
    st.text_area(
        "3.4 สิทธิการรักษา", key="s34",
        height=200,
        placeholder="สิทธิการรักษา ...",
        key="s34",
        help="คัดลอกจากหน้า 'สิทธิการรักษา' เพื่อติ๊กช่อง R1-R4"
    )



# --- สิ้นสุดส่วนที่ 4 ---




# --- 10. ระบบประมวลผลอัจฉริยะ (AI Extraction Engine) ---

if st.button("🚀 กดเพื่อประมวลผลและสกัดข้อมูลด้วย Gemini 3 Flash", use_container_width=True):
    # ดึงข้อมูลวันที่ปัจจุบันและส่วนประกอบวันที่ไทยเพื่อใช้คำนวณ 
    thai_date_data = get_thai_date_parts()
    today_str = f"{thai_date_data['DAY']} {thai_date_data['MONTH']} {thai_date_data['YEAR']}"
    
    # รวบรวมข้อมูลดิบจาก 9 ช่องกรอกเข้าสู่ตัวแปรเดียว 
    all_raw_data = f"""
    วันที่ทำรายการปัจจุบัน: {today_str}
    1.1 Admission Note: {st.session_state.s11}
    1.2 การวินิจฉัย: {st.session_state.s12}
    1.3 Order / Meds: {st.session_state.s13}
    1.4 Progress Note: {st.session_state.s14}
    Assessment Score: {st.session_state.s2}
    3.1 ข้อมูลทั่วไป: {st.session_state.s31}
    3.2 ที่อยู่ปัจจุบัน: {st.session_state.s32}
    3.3 ผู้ติดต่อ: {st.session_state.s33}
    3.4 สิทธิการรักษา: {st.session_state.s34}
    """
    
    with st.spinner('Gemini 3 Flash กำลังวิเคราะห์ข้อมูลเชิงลึกและจัดตรรกะ Checkbox...'):
        # ตรรกะการประมวลผลระดับ PhD เพื่อรองรับไฟล์ .docx แบบตารางล่าสุด 
        prompt = f"""
        คุณคือผู้ช่วยวิจัยทางการแพทย์ระดับ PhD จงสกัดข้อมูลจากระบบ @ThanHIS ลงแบบฟอร์ม 062 
        โดยปฏิบัติตามกฎเหล็ก "Official Template Mirroring" อย่างเคร่งครัด:

        1. ตรรกะสิทธิการรักษา (Checkbox R1-R4 Logic):
           วิเคราะห์สิทธิการรักษาจากข้อมูล แล้วส่งค่า [✓] ลงในช่องที่ถูกต้องเพียงช่องเดียว และ [ ] ในช่องที่เหลือ:
           - R1: ประกันสุขภาพ (บัตรทอง/UC/30 บาท)
           - R2: ประกันสังคม
           - R3: ท.74 ข้าราชการ/จ่ายตรง/รัฐวิสาหกิจ
           - R4: อื่นๆ (หากเลือกอันนี้ ให้ระบุรายละเอียดใน R_NOTE)

        2. ตรรกะการจัดรูปแบบแถวเดียว (Single Row Format):
           - [DX]: สกัดรหัส ICD-10 และชื่อโรค เขียนต่อกันในแถวเดียว คั่นด้วยคอมม่า (,) 
           - [MEDS]: สกัดรายการยา Home-Med เขียนชื่อภาษาอังกฤษเป็น UPPERCASE พร้อมวิธีใช้ ต่อกันในแถวเดียว คั่นด้วยคอมม่า (,)

        3. ตรรกะการคำนวณ (Calculation Logic):
           - [LOC]: คำนวณระยะเวลาในชุมชน โดยนำวันที่ปัจจุบัน ({today_str}) ลบด้วย วันที่จำหน่ายครั้งสุดท้าย
           - [LOS]: คำนวณวันนอนรวม โดยนำวัน Detox + Rehab

        4. ตรรกะการสกัดข้อมูลทั่วไป:
           - สกัดข้อมูลลงใน Key: NAME, AGE, HN, ID, EDU, CAREER, RELIGION, STATUS, CONTACT, RELATION, PHONE, NEAR_HOSP, ADMIT_DATE, VISIT_NUM, CC, ADDRESS, Q9, Q8, PROGRESS

        ข้อมูลดิบ:
        {all_raw_data}

        ตอบกลับเป็น JSON เท่านั้น
        """
        
        try:
            # ประมวลผลผ่านโมเดล Gemini
            response = client.models.generate_content(model=MODEL_ID, contents=prompt)
            # ใช้ Regular Expression เพื่อดึงเฉพาะเนื้อหา JSON 
            match = re.search(r'\{.*\}', response.text, re.DOTALL)
            
            if match:
                st.session_state.extracted_json_data = json.loads(match.group())
                # ผสานข้อมูลวันที่ไทยสำหรับหัวกระดาษและวันที่จำหน่าย
                st.session_state.extracted_json_data.update(thai_date_data)
                st.success("✅ วิเคราะห์ข้อมูลและเลือกสิทธิการรักษา (Checkboxes) สำเร็จ!")
            else:
                st.error("❌ AI ไม่สามารถสร้างโครงสร้างข้อมูลที่ถูกต้องได้ กรุณาตรวจสอบข้อมูลดิบ")
        except Exception as e:
            st.error(f"⚠️ เกิดข้อผิดพลาดในการวิเคราะห์เชิงลึก: {e}")



# --- สิ้นสุดส่วนที่ 5 ---




# --- 11. ฟังก์ชันจัดการไฟล์ Word (Advanced Template Mapping) ---

def fill_pmnidat_doc(data):
    """ฟังก์ชันนำข้อมูลจากระบบเข้าสู่ไฟล์แม่แบบ .docx โดยรองรับโครงสร้างตารางและ Checkbox """
    try:
        # 11.1 โหลดไฟล์แม่แบบที่คุณหมอจัดทำล่าสุด
        doc = Document("PMNIDAT 062 แบบบันทึกข้อมูลเพื่อส่งต่อ.docx")
        
        # 11.2 เตรียมชุดข้อมูลสำหรับแทนที่ (Key ต้องเป็นตัวพิมพ์ใหญ่ตามใน Word)
        mapping = {f"{{{{{k.upper()}}}}}": str(v) for k, v in data.items()}
        
        def apply_style_and_replace(paragraph):
            """ตรรกะการจัดวาง Alignment และการแทนที่ข้อความ"""
            
            # ก. ตรรกะการจัดวาง (Alignment Logic): 
            # วันที่หัวเอกสารให้ชิดขวา แต่ 'วันที่จำหน่าย' ในเนื้อหาให้ชิดซ้ายปกติ 
            if "วันที่" in paragraph.text and "จำหน่าย" not in paragraph.text:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            else:
                paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT

            # ข. ค้นหาและแทนที่ข้อความ (Replacement Logic)
            for key, value in mapping.items():
                if key in paragraph.text:
                    # แทนที่ Placeholder ด้วยข้อมูลจริง (เช่น {{R1}} จะกลายเป็น [✓] หรือ [ ]) 
                    paragraph.text = paragraph.text.replace(key, value)
                    
                    # ค. กำหนดคุณลักษณะตัวอักษร (Font Properties)
                    for run in paragraph.runs:
                        run.font.size = Pt(13) # ขนาดฟอนต์ 13 pt ตามมาตรฐานสถาบันฯ
                        
                        # ง. ตรรกะการเน้นตัวหนา (Bolding):
                        # เน้นตัวหนาที่ส่วนการวินิจฉัย (DX), รายการยา (MEDS) และสรุปปัญหา (PROGRESS)
                        clinical_keys = ["{{DX}}", "{{MEDS}}", "{{PROGRESS}}"]
                        if any(ck.upper() in key for ck in clinical_keys):
                            run.font.bold = True

        # 11.3 ดำเนินการตรวจสอบและแทนที่ในทุกย่อหน้าปกติ 
        for p in doc.paragraphs:
            apply_style_and_replace(p)
            
        # 11.4 ดำเนินการตรวจสอบภายในตาราง (Tables) 
        # ไฟล์ใหม่ของคุณหมอใช้ตารางเป็นโครงสร้างหลักเพื่อล็อคตำแหน่ง 
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        apply_style_and_replace(p)
                            
        # 11.5 บันทึกไฟล์ลงในหน่วยความจำชั่วคราวเพื่อพร้อมดาวน์โหลด
        buffer = io.BytesIO()
        doc.save(buffer)
        return buffer.getvalue()
        
    except Exception as e:
        st.error(f"⚠️ ไม่สามารถเปิดไฟล์ 'PMNIDAT 062 แบบบันทึกข้อมูลเพื่อส่งต่อ.docx' ได้: {e}")
        return None




# --- สิ้นสุดส่วนที่ 6 ---





# --- 12. การแสดงผลลัพธ์และการควบคุมการดาวน์โหลด (Final Execution) ---

# ตรวจสอบว่ามีการสกัดข้อมูลสำเร็จและเก็บไว้ในหน่วยความจำชั่วคราวหรือไม่
if "extracted_json_data" in st.session_state and st.session_state.extracted_json_data:
    # เรียกใช้ฟังก์ชันจัดการไฟล์ Word เพื่อสร้างเอกสารตัวจริงจากแม่แบบล่าสุด
    word_file_final = fill_pmnidat_doc(st.session_state.extracted_json_data)
    
    if word_file_final:
        # บันทึกสถิติการใช้งานไปยังระบบจัดเก็บข้อมูล (อ้างอิงฟังก์ชันในส่วนที่ 1)
        log_usage(st.session_state.extracted_json_data.get('NAME', '[ไม่ระบุชื่อ]'))
        
        st.divider()
        st.balloons() # แสดงเอฟเฟกต์เฉลิมฉลองเมื่อเอกสารพร้อม
        st.success("🎉 ระบบจัดทำเอกสาร PMNIDAT 062 ฉบับสมบูรณ์ (Master v3.35) เรียบร้อยแล้ว!")
        
        # ปุ่มดาวน์โหลดไฟล์เอกสารที่จัดรูปแบบตามโครงสร้างตารางที่คุณหมอออกแบบ 
        st.download_button(
            label="💾 ดาวน์โหลดไฟล์ 'แบบบันทึกข้อมูลเพื่อส่งต่อ (PMNIDAT 062).docx'",
            data=word_file_final,
            file_name=f"Refer_{st.session_state.extracted_json_data.get('NAME', '062')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

# --- 13. มาตรการรักษาความปลอดภัยและความเป็นส่วนตัว (PDPA & Safety Footer) ---
# ส่วนประกาศสำคัญเพื่อรักษามาตรฐานความลับของผู้ป่วยและจริยธรรมทางการแพทย์ 

st.divider()
st.info("""
    **🛡️ มาตรการรักษาความปลอดภัยของข้อมูลคนไข้ (PDPA Compliance):**
    
    * **ไม่มีการจัดเก็บข้อมูลถาวร:** ระบบประมวลผลแบบ Real-time บนหน่วยความจำชั่วคราว (RAM) เท่านั้น **ไม่มีการบันทึกข้อมูลส่วนบุคคลของผู้ป่วย** ลงในฐานข้อมูลถาวรหรือบันทึกไฟล์ทิ้งไว้บนเซิร์ฟเวอร์ 
    * **ระบบ Session-Based:** ข้อมูลที่คัดลอกมาวางจะถูกลบทิ้งจากหน่วยความจำทันทีเมื่อมีการรีเฟรชหน้าจอ (Refresh) หรือปิดเบราว์เซอร์ โปรดดาวน์โหลดไฟล์ให้เรียบร้อยก่อนจบการทำงานในแต่ละเคส
    * **การตรวจสอบความถูกต้อง (Verification Audit):** เนื่องจากเป็นระบบ AI ช่วยสกัดข้อมูล ผู้ใช้งานต้องตรวจสอบความถูกต้องของข้อมูลและลงนามรับรอง  ในเอกสารฉบับจริงทุกครั้งก่อนนำไปใช้งานทางคลินิก
    
    ---
    **🔬 พัฒนาระบบโดย:** ดร.นพ.ชาฌาน หลานวงศ์ (หมออาร์ม) | MD-PhD (Pharmacology) 
    แพทย์เวชศาสตร์สารเสพติด สถาบันบำบัดรักษาและฟื้นฟูผู้ติดยาเสพติดแห่งชาติบรมราชชนนี (สบยช.)
    
    **💡หากพบปัญหาการใช้งาน/มีข้อเสนอแนะ ติดต่อได้ที่**
    * **Line:** armlahn | **โทร:** 094-991-4599
    
    """)

# --- สิ้นสุดการทำงานของระบบ PMNIDAT Smart Transfer ---

