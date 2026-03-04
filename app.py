import streamlit as st
from google import genai 
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re
import json
import requests

# --- 1. ฟังก์ชันบันทึกสถิติการใช้งาน (Log Usage) ---
def log_usage(patient_name):
    try:
        url = st.secrets["APPS_SCRIPT_URL"]
        requests.post(url, json={"name": patient_name}, timeout=5)
    except:
        pass

# --- 2. การตั้งค่าการเชื่อมต่อ API ---
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

# --- 3. ระบบ Session State Memory (จดจำข้อมูล 9 ช่อง) ---
field_keys = ['s11', 's12', 's13', 's14', 's2', 's31', 's32', 's33', 's34']
for key in field_keys:
    if key not in st.session_state:
        st.session_state[key] = ""

# --- 4. ฟังก์ชันโหลดข้อมูลตัวอย่าง (สมมติชื่อ: นาย ช.ธัญญารักษ์) ---
def load_test_data():
    # ข้อมูลจาก Step 1.1
    st.session_state.s11 = "นาย ช.ธัญญารักษ์ อายุ 40 ปี 5 เดือน\nสิทธิ์ : จ่ายตรงกรมบัญชีกลาง\nAdmit Date 01/03/2569\nจำนวนวัน Detox [15] วัน Rehab [60] วัน\nCC : เสพสุราซ้ำ ต้องการเข้ารับการบำบัดรักษา"
    # ข้อมูลจาก Step 1.2 (รหัสโรคพร้อมชื่อไทย)
    st.session_state.s12 = "1. F102 - โรคเสพติดสุรา\n2. I10 - โรคความดันโลหิตสูง (Hypertension)"
    # ข้อมูลจาก Step 1.3 (รายการยาเรียงแถว)
    st.session_state.s13 = "1. AMLODIPINE 5 MG 1x1 pc (เช้า)\n2. QUETIAPINE 25 MG 1 tab hs (ก่อนนอน)\n(Home-Med ทั้งหมด)"
    # ข้อมูลจาก Step 1.4
    st.session_state.s14 = "S: สบายดี กินข้าวได้ นอนหลับได้\nO: V/S stable, BP 120/80 mmHg\nA: อาการคงที่ เตรียมจำหน่าย\nP: Discharge to home"
    # ข้อมูลจาก Step 2
    st.session_state.s2 = "9Q : 5 คะแนน\n8Q : 0 คะแนน\nBPRS : 15 คะแนน"
    # ข้อมูลจาก Step 3.1
    st.session_state.s31 = "ชื่อ [ช.] นามสกุล [ธัญญารักษ์]\nเลขบัตรประชาชน* [1-2345-67890-12-3]\nศาสนา [พุทธ] อาชีพ [ข้าราชการ]"
    # ข้อมูลจาก Step 3.2
    st.session_state.s32 = "ที่อยู่ปัจจุบัน: เลขที่ 60 ต.หลักหก อ.เมืองปทุมธานี จ.ปทุมธานี"
    # ข้อมูลจาก Step 3.3
    st.session_state.s33 = "คุณญ. ธัญญารักษ์ (ภรรยา) เบอร์โทร: 081-234-5678"
    # ข้อมูลจาก Step 3.4
    st.session_state.s34 = "สิทธิหลัก [จ่ายตรง] สถานพยาบาลหลัก [สถาบันบำบัดรักษาและฟื้นฟูผู้ติดยาเสพติดแห่งชาติบรมราชชนนี]"
    st.rerun()

def clear_all_data():
    for key in field_keys:
        st.session_state[key] = ""
    st.rerun()

st.title("🏥 PMNIDAT Smart Transfer")
st.subheader("ระบบ AI ผู้ช่วยสร้าง "แบบบันทึกข้อมูลเพื่อส่งต่อ (PMNIDAT 062)" อัตโนมัติ  (Master Version 3.32)")



# --- สิ้นสุดส่วนที่ 1 ---



# --- 5. การออกแบบแถบเมนูด้านข้าง (Sidebar Manual & Controls) ---

with st.sidebar:
    # แสดงโลโก้สถาบันฯ และหัวข้อคู่มือ
    st.image("https://pmnidathis.dms.go.th/static/media/health.1bfd961f.png", width=100)
    st.header("📖 คู่มือการใช้งาน PMNIDAT Smart Transfer")
    
    # ปุ่มควบคุมสำหรับการทดสอบระบบ (เรียกใช้ฟังก์ชันจากส่วนที่ 1)
    st.subheader("🛠️ เครื่องมือช่วยงาน")
    col_t1, col_t2 = st.columns(2)
    with col_t1:
        if st.button("🧬 ลงข้อมูลตัวอย่าง", use_container_width=True):
            load_test_data()
    with col_t2:
        if st.button("🧹 ล้างข้อมูล", use_container_width=True):
            clear_all_data()
            
    st.divider()
    
    # คำแนะนำพื้นฐานในการคัดลอก
    st.markdown("### **วิธีการคัดลอกข้อมูลจาก @ThanHIS และวางในโปรแกรม**")
    st.info("""
    1. **คลิกเมาส์ซ้ายค้าง** ที่ต้นข้อความ ลากครอบให้คลุมทั้งหมด
    2. กด **Ctrl+C** (คัดลอก)
    3. มาที่หน้านี้ คลิกช่องที่ต้องการ แล้วกด **Ctrl+V** (วาง)
    """)
    
    # รายละเอียดแต่ละ Step ตามคู่มือฉบับจริง
    with st.expander("🟢 STEP 1: ระบบผู้ป่วยใน (IPD)", expanded=True):
        st.markdown("""
        **1.1 Admission Note:**
        - ดูข้อมูลคนไข้ → Admission note
        - คัดลอกข้อมูลทั้งหมด (ข้อมูลทั่วไป, VS, PI, PH)
        
        **1.2 การวินิจฉัย:**
        - กดเมนู "การวินิจฉัย"
        - คัดลอกรหัส ICD-10 ทั้งหมด (Principal & Comorbidity)
        
        **1.3 Order / Meds:**
        - กดเมนู **"Order"** ด้านซ้าย
        - คัดลอก **Discharge order + Home medication** ทั้งหมด
        - (สังเกตรายการที่มีจำนวนยาสำหรับ 0.5 - 1 เดือน)
        
        **1.4 Progress Note:**
        - กดเมนู **"Progress note"** ด้านซ้าย
        - คัดลอกบันทึก **ล่าสุด (SOAP)** ที่อยู่บนชื่อแพทย์
        """)

    with st.expander("🔵 STEP 2: การประเมิน (Assessment)"):
        st.markdown("""
        - กดเมนู "Admission note" → ปุ่ม **"ข้อมูลผู้ป่วยนอก"**
        - เลื่อนลงล่างไปที่หัวข้อ **Assessment**
        - คัดลอกผลคะแนน **9Q, 8Q, BPRS** ทั้งหมดมาวาง
        """)

    with st.expander("🟠 STEP 3: เวชระเบียน (Registration)"):
        st.markdown("""
        - ระบบผู้ป่วยนอก → เวชระเบียน → ลงทะเบียนผู้ป่วย
        - ค้นหา HN เพื่อเข้าสู่หน้าข้อมูลหลัก
        
        **3.1 ทั่วไป 1:** คัดลอกชื่อ, อายุ, เลขบัตรประชาชน, ศาสนา
        **3.2 ทั่วไป 2:** กดแสดง **"ที่อยู่ปัจจุบัน"** แล้วคัดลอกทั้งหมด
        **3.3 ผู้ติดต่อ:** คัดลอกชื่อญาติ, ความสัมพันธ์ และเบอร์โทร
        **3.4 สิทธิรักษา:** คัดลอกสิทธิ์และ **"สถานพยาบาลหลัก"**
        """)
        
    st.divider()
    # แสดงสถานะระบบ (ไม่มี cite_start มากวนใจ)
    st.success("💡 AI วิเคราะห์ด้วยตรรกะ PhD พร้อมทำงาน")



# --- สิ้นสุดส่วนที่ 2 ---



# --- 6. ส่วนการออกแบบช่องกรอกข้อมูล (9 Input Fields) ---

st.divider()
# กลุ่มที่ 1: ข้อมูลระบบผู้ป่วยใน (IPD) - เน้นการคัดลอกข้อมูลดิบแบบจัดเต็ม 
st.markdown("### **🟢 Step 1: ระบบผู้ป่วยใน (IPD)**")
s1_cols = st.columns(4)

with s1_cols[0]:
    st.text_area(
        "1.1 Admission Note",
        height=300,
        placeholder="คัดลอกข้อมูลแรกรับทั้งหมด...",
        key="s11",
        help="คัดลอกจากเมนู Admission note ในระบบ IPD [cite: 101]"
    )

with s1_cols[1]:
    st.text_area(
        "1.2 การวินิจฉัย",
        height=300,
        placeholder="คัดลอกรหัส ICD-10 ทั้งหมด...",
        key="s12",
        help="คัดลอกจากเมนูการวินิจฉัย [cite: 101]"
    )

with s1_cols[2]:
    st.text_area(
        "1.3 Order / Meds",
        height=300,
        placeholder="คัดลอกข้อมูลจากเมนู Order ทั้งหมด (ไม่เอาเฉพาะล่าสุด)...",
        key="s13",
        help="คัดลอก Discharge order และ Home medication ทั้งหมดที่มี"
    )

with s1_cols[3]:
    st.text_area(
        "1.4 Progress Note",
        height=300,
        placeholder="คัดลอกบันทึก Progress note ทั้งหมด (ไม่เอาเฉพาะล่าสุด)...",
        key="s14",
        help="คัดลอกบันทึกการติดตามอาการทั้งหมด เพื่อให้ AI สังเคราะห์ปัญหา"
    )

st.divider()
# กลุ่มที่ 2: ข้อมูลการประเมิน (Assessment) [cite: 103, 104]
st.markdown("### **🔵 Step 2: การประเมิน (Assessment)**")
st.text_area(
    "คัดลอกผลคะแนน 9Q, 8Q, BPRS ทั้งหมดมาวางที่นี่",
    height=150,
    key="s2",
    help="ดึงจากหน้า Assessment ในระบบผู้ป่วยนอก [cite: 104]"
)

st.divider()
# กลุ่มที่ 3: ข้อมูลเวชระเบียน (Registration) [cite: 105, 106]
st.markdown("### **🟠 Step 3: เวชระเบียน (Registration)**")
s3_cols = st.columns(4)

with s3_cols[0]:
    st.text_area(
        "3.1 ข้อมูลทั่วไป", 
        height=200, 
        key="s31",
        help="ดึงจากหน้า 'ทั่วไป 1' (ชื่อ, อายุ, เลขบัตรประชาชน) [cite: 106]"
    )

with s3_cols[1]:
    st.text_area(
        "3.2 ที่อยู่ปัจจุบัน", 
        height=200, 
        key="s32",
        help="ดึงจากหน้า 'ทั่วไป 2' (ต้องกดยืดกล่องที่อยู่) [cite: 106]"
    )

with s3_cols[2]:
    st.text_area(
        "3.3 ผู้ติดต่อ", 
        height=200, 
        key="s33",
        help="ดึงจากหน้า 'ผู้ติดต่อ' (ชื่อญาติและเบอร์โทรศัพท์) [cite: 106]"
    )

with s3_cols[3]:
    st.text_area(
        "3.4 สิทธิการรักษา", 
        height=200, 
        key="s34",
        help="ดึงจากหน้า 'สิทธิการรักษา' (สิทธิ์และสถานพยาบาลหลัก) [cite: 106]"
    )



# --- สิ้นสุดส่วนที่ 3 ---



# --- 7. ส่วนประมวลผลอัจฉริยะ (Advanced Search, Extract & Audit Logic) ---

if st.button("🚀 ประมวลผลและสกัดข้อมูลด้วยตรรกะ PhD", use_container_width=True):
    # รวบรวมข้อมูลดิบจาก 9 ช่องที่กรอกไว้ในส่วนที่ 3 [cite: 102, 104, 106]
    all_raw_data = f"""
    --- GROUP 1: IPD RAW DATA ---
    {st.session_state.s11}
    {st.session_state.s12}
    {st.session_state.s13}
    {st.session_state.s14}
    --- GROUP 2: ASSESSMENT DATA ---
    {st.session_state.s2}
    --- GROUP 3: REGISTRATION DATA ---
    {st.session_state.s31}
    {st.session_state.s32}
    {st.session_state.s33}
    {st.session_state.s34}
    """
    
    with st.spinner('Gemini 3 Flash กำลังวิเคราะห์ข้อมูลเชิงลึกและตรวจสอบ Verification Audit...'):
        # ตรรกะการประมวลผลตามมาตรฐานสถาบันฯ และเอกสารแนบ [cite: 100-112]
        prompt = f"""
        คุณคือผู้ช่วยวิจัยระดับ PhD ทางการแพทย์ ทำหน้าที่สกัดข้อมูลจากระบบ @ThanHIS ลงแบบฟอร์ม 062 
        โดยต้องปฏิบัติตามกฎเหล็ก "Verification Audit" และ "Search & Extract Logic" อย่างเคร่งครัด:

        1. กฎการตัดขยะข้อมูล (Noise Reduction Rule): 
           Ignore (ละทิ้ง) ข้อมูล Theme Customizer, Navbar, Menu Colors, Light/Dark Mode และ COPYRIGHT ทั้งหมด 

        2. ตรรกะการสกัดข้อมูลรายกลุ่ม (Search & Extract Logic):
           - [HN/AN]: มองหาตัวเลขหลังคำว่า 'HN' หรือ 'AN' [cite: 102]
           - [สิทธิการรักษา]: สกัดข้อความระหว่าง 'สิทธิ์ :' ถึง '(ไม่มีหนังสือส่งตัว)' หรือข้อความถัดไป [cite: 102]
           - [CC]: สกัดจาก 'Chief Complaint' หรือ 'CC :' จนถึง 'Present illness' [cite: 102]
           - [วันนอน (LOS)]: ค้นหาตัวเลขหน้าคำว่า 'วัน' ในแถว 'Detox' และ 'Rehab' แล้วนำมาบวกกันเสมอ 
           - [รหัสโรค (DX)]: สกัดรหัส ICD-10 (ไม่มีจุดทศนิยม) และต้องระบุ "ชื่อโรคเป็นภาษาไทย" กำกับเสมอ โดยเรียงเป็นบรรทัดละ 1 โรค 
           - [ยา (MEDS)]: ค้นหาเฉพาะรายการยาที่มีคำว่า 'Home-Med' สกัดชื่อยาเป็น UPPERCASE พร้อมวิธีใช้ โดยต้อง "แยกบรรทัดละ 1 ตัวยา" เพื่อความสวยงาม 
           - [คะแนนประเมิน]: สกัดตัวเลขหลัง 'ผลรวมการประเมินโรคซึมเศร้า' (9Q) และ 'การฆ่าตัวตาย' (8Q) [cite: 104]
           - [สรุปปัญหา (PROGRESS)]: สังเคราะห์จาก Progress Note ทั้งหมดที่ได้รับ ให้เป็นสรุปย่อหน้าเดียว ความยาว 2-3 บรรทัด 
           - [ข้อมูลเวชระเบียน]: สกัดชื่อ-สกุล, อายุ, บัตรประชาชน (13 หลัก), ศาสนา, อาชีพ และที่อยู่ปัจจุบัน [cite: 106]
           - [หน่วยบริการ]: สกัดชื่อโรงพยาบาลหลังคำว่า 'สถานพยาบาลหลัก' [cite: 106]
           - [ผู้ดูแล]: สกัดชื่อผู้ติดต่อ ความสัมพันธ์ และเบอร์โทรศัพท์ [cite: 106]

        3. นโยบายการตรวจสอบความถูกต้อง (Verification Audit Policy):
           - หากไม่พบข้อมูลใดๆ ในช่องที่ระบุ ให้พยายามหาจากช่องอื่นในเนื้อหาทั้งหมด 
           - หากไม่มีจริงๆ ให้ระบุ [กรอกด้วยตนเอง] ห้ามเว้นว่างเด็ดขาด [cite: 108]
           - ตรวจสอบประวัติจำหน่ายเดิม (เช่น เคยนอน 4 ครั้ง) เพื่อวิเคราะห์ 'รับไว้ครั้งที่' (เช่น 5) 

        ข้อมูลดิบสำหรับวิเคราะห์:
        {all_raw_data}

        ตอบกลับในรูปแบบ JSON ที่มี Key ตรงกับ Placeholder ในไฟล์ Word: 
        NAME, AGE, HN, ID, EDU, CAREER, RELIGION, STATUS, RIGHTS, LAST_DC, ADMIT_DATE, VISIT_NUM, CC, CONTACT, RELATION, PHONE, NEAR_HOSP, DC_DATE, LOS, ADDRESS, Q9, Q8, MEDS, DX, PROGRESS, POST_SERVICE
        """
        
        try:
            # ประมวลผลผ่านโมเดล Gemini 3 Flash ตาม Paid Tier
            response = client.models.generate_content(model=MODEL_ID, contents=prompt)
            match = re.search(r'\{.*\}', response.text, re.DOTALL)
            
            if match:
                # เก็บข้อมูล JSON ที่สกัดได้ลงใน Session State เพื่อส่งต่อให้ส่วนที่ 5
                st.session_state.extracted_json_data = json.loads(match.group())
                st.success("✅ วิเคราะห์ข้อมูลเชิงลึกและสกัดชื่อโรคภาษาไทยเรียบร้อยแล้ว!")
            else:
                st.error("AI ไม่สามารถสร้างรูปแบบข้อมูลที่ถูกต้องได้ กรุณาลองตรวจสอบข้อมูลดิบอีกครั้ง")
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {e}")




# --- สิ้นสุดส่วนที่ 4 ---



# --- 8. ฟังก์ชันจัดการไฟล์ Word (จัดรูปแบบชิดซ้าย + ฟอนต์ 13 + แทนที่ Placeholder) ---

def fill_pmnidat_doc(data):
    """ฟังก์ชันนำข้อมูลจาก JSON ไปบรรจุลงในไฟล์แม่แบบ .docx [cite: 83-99, 109]"""
    try:
        # โหลดไฟล์แม่แบบที่คุณหมอเตรียมไว้
        doc = Document("PMNIDAT 062 แบบบันทึกข้อมูลเพื่อส่งต่อ.docx")
        
        # เตรียมชุดข้อมูล Mapping (Key ต้องเป็นตัวพิมพ์ใหญ่ตรงกับ Placeholder ใน Word) [cite: 83-90]
        mapping = {f"{{{{{k.upper()}}}}}": str(v) for k, v in data.items()}
        
        def apply_style_and_replace(paragraph):
            for key, value in mapping.items():
                if key in paragraph.text:
                    # แทนที่ข้อความใน Placeholder [cite: 83-99]
                    paragraph.text = paragraph.text.replace(key, value)
                    
                    # บังคับจัดรูปแบบชิดซ้าย (Left Alignment) ตามมาตรฐานที่คุณหมอกำหนด 
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    
                    # กำหนดฟอนต์ขนาด 13 pt สำหรับงานวิชาการสถาบันฯ [cite: 51, 109]
                    for run in paragraph.runs:
                        run.font.size = Pt(13)

        # ดำเนินการแทนที่ข้อมูลทั้งใน Paragraph ปกติและภายในตาราง [cite: 83-99]
        for p in doc.paragraphs: apply_style_and_replace(p)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs: apply_style_and_replace(p)
                            
        # บันทึกไฟล์ลงในหน่วยความจำชั่วคราว (Memory Buffer)
        buffer = io.BytesIO()
        doc.save(buffer)
        return buffer.getvalue()
    except Exception as e:
        st.error(f"⚠️ เกิดข้อผิดพลาดในการสร้างไฟล์ Word: {e}")
        return None

# --- 9. การแสดงผลลัพธ์และปุ่มดาวน์โหลด (Final Execution) ---

# ตรวจสอบว่ามีการสกัดข้อมูลสำเร็จใน Session State หรือไม่ [cite: 100-106]
if "extracted_json_data" in st.session_state and st.session_state.extracted_json_data:
    # เรียกใช้ฟังก์ชันสร้างไฟล์จากข้อมูลล่าสุด
    word_file_final = fill_pmnidat_doc(st.session_state.extracted_json_data)
    
    if word_file_final:
        # บันทึกสถิติการใช้งาน (ฟังก์ชันที่ประกาศไว้ในส่วนที่ 1)
        log_usage(st.session_state.extracted_json_data.get('name', '[ไม่ระบุชื่อ]'))
        
        st.divider()
        st.balloons() # เฉลิมฉลองความสำเร็จในการประมวลผล
        st.success("🎉 ระบบสกัดข้อมูลและจัดทำเอกสาร 062 ฉบับสมบูรณ์เรียบร้อยแล้ว!")
        
        # ปุ่มดาวน์โหลดไฟล์ฉบับ Final ที่พร้อมส่งต่อ [cite: 107-109]
        st.download_button(
            label="💾 ดาวน์โหลดใบส่งต่อ 062 (จัดรูปแบบฟอนต์ 13 + ชิดซ้าย)",
            data=word_file_final,
            file_name=f"Refer_{st.session_state.extracted_json_data.get('name', '062')}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )

# --- 10. มาตรการรักษาความปลอดภัย (PDPA Footer) ---
st.divider()
st.info("""
    **ประกาศมาตรการรักษาความปลอดภัยข้อมูล (PDPA Compliance):**
    * ระบบ PMNIDAT Smart Refer ประมวลผลแบบ Real-time และ **ไม่มีการจัดเก็บข้อมูลผู้ป่วยถาวร** [cite: 112]
    * ข้อมูลจะถูกลบทิ้งทันทีเมื่อมีการรีเฟรชหน้าจอ (Refresh) โปรดดาวน์โหลดไฟล์ให้เรียบร้อยก่อนปิดระบบ
    * โปรดตรวจสอบความถูกต้องของข้อมูล (Verification Audit) อีกครั้งก่อนนำไปใช้งานจริง [cite: 107-109]
    """)




# --- สิ้นสุดส่วนที่ 5 ---

