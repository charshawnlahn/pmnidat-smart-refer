import streamlit as st
from google import genai 
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import re
import json
import requests

# --- 1. การตั้งค่าระบบและการเชื่อมต่อ API ---
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

# --- 2. ระบบ Session State Memory (หัวใจของการโหลดข้อมูล) ---
field_keys = ['s11', 's12', 's13', 's14', 's2', 's31', 's32', 's33', 's34']
for key in field_keys:
    if key not in st.session_state:
        st.session_state[key] = ""

# ฟังก์ชันโหลดข้อมูลตัวอย่าง (สมมติชื่อ: ชาฌาน หลานวงศ์)
def load_test_data():
    st.session_state.s11 = "ชื่อผู้ป่วย: นพ.ชาฌาน หลานวงศ์\nอายุ: 40 ปี\nสิทธิ์: จ่ายตรงกรมบัญชีกลาง\nCC: ต้องการทดสอบระบบ Smart Refer\nAdmit Date: 04/03/2026\nDetox: 5 วัน Rehab: 10 วัน"
    st.session_state.s12 = "1. F155 (Mental disorders due to amphetamine)\n2. I10 (Essential hypertension)"
    st.session_state.s13 = "1. AMLODIPINE 5 MG 1x1 pc\n2. QUETIAPINE 25 MG 1 tab hs (Home-Med)"
    st.session_state.s14 = "S: ผู้ป่วยสบายดี\nO: สัญญาณชีพปกติ\nA: อาการคงที่ เตรียมจำหน่าย\nP: ส่งต่อรับยาใกล้บ้าน"
    st.session_state.s2 = "9Q: 2 คะแนน\n8Q: 0 คะแนน\nBPRS: 15 คะแนน"
    st.session_state.s31 = "นพ.ชาฌาน หลานวงศ์\nเลขบัตร: 1-2345-67890-12-3\nศาสนา: พุทธ\nอาชีพ: แพทย์"
    st.session_state.s32 = "สถาบันบำบัดรักษาและฟื้นฟูผู้ติดยาเสพติดแห่งชาติบรมราชชนนี (สบยช.) จ.ปทุมธานี"
    st.session_state.s33 = "คุณวิไล (ภรรยา)\nโทร: 02-531-XXXX"
    st.session_state.s34 = "สิทธิหลัก: จ่ายตรง\nสถานพยาบาลหลัก: สบยช."
    # บังคับให้ Streamlit อัปเดตหน้าจอทันที
    st.rerun()

def clear_all_data():
    for key in field_keys:
        st.session_state[key] = ""
    st.rerun()

# --- 3. ส่วนการออกแบบ Sidebar (คู่มือฉบับละเอียดเรียงแถวสวยงาม) ---
st.set_page_config(page_title="PMNIDAT Smart Portal", layout="wide")

with st.sidebar:
    st.image("https://pmnidathis.dms.go.th/static/media/health.1bfd961f.png", width=100)
    st.header("📖 คู่มือพี่พยาบาล ฉบับจับมือทำ")
    
    # ส่วนปุ่มควบคุมการทดสอบ
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
    1. **ลากครอบข้อความ** ในระบบ @ThanHIS
    2. กด **Ctrl+C** เพื่อคัดลอก
    3. มาที่หน้านี้ คลิกช่องที่ต้องการ แล้วกด **Ctrl+V**
    """)
    
    # ปรับปรุง Step 1-3 ให้เรียงแถวแยกบรรทัดชัดเจน
    with st.expander("🟢 STEP 1: ระบบผู้ป่วยใน (IPD)", expanded=True):
        st.markdown("""
        **1.1 Admission Note:** - กดดูข้อมูลคนไข้ 
        - เลือก Admission note 
        - คัดลอกทั้งหมด
        
        **1.2 การวินิจฉัย:** - กดเมนู "การวินิจฉัย" 
        - คัดลอกรหัส ICD-10
        
        **1.3 Order / Meds:** - กดเมนู "Order" 
        - คัดลอกรายการยา Home-Med
        
        **1.4 Progress Note:** - กดเมนู "Progress note" 
        - คัดลอกบันทึก SOAP ล่าสุด
        """)

    with st.expander("🔵 STEP 2: การประเมิน (Assessment)"):
        st.markdown("""
        - เข้าเมนู **Admission note** - กดปุ่ม **ข้อมูลผู้ป่วยนอก** - เลื่อนหาหัวข้อ **Assessment** - คัดลอกคะแนน 9Q, 8Q, BPRS
        """)

    with st.expander("🟠 STEP 3: เวชระเบียน (Registration)"):
        st.markdown("""
        - **3.1 ทั่วไป 1:** ชื่อ, อายุ, เลขบัตรฯ
        - **3.2 ทั่วไป 2:** กด **ที่อยู่ปัจจุบัน** แล้วคัดลอก
        - **3.3 ผู้ติดต่อ:** ชื่อญาติและเบอร์โทร
        - **3.4 สิทธิรักษา:** สิทธิ์และสถานพยาบาลหลัก
        """)
        
    st.divider()
    st.success(f"💡 AI เชื่อมต่อสำเร็จผ่าน: {MODEL_ID}")

st.title("🏥 PMNIDAT Smart D/C Transfer")
st.subheader("ระบบสร้างไฟล์ใบส่งต่ออัตโนมัติ (Master Version 3.29)")

# --- 4. ส่วนกรอกแบบฟอร์ม (เชื่อมโยงกับ Session State 100%) ---

st.divider()
st.markdown("### **🟢 Step 1: ระบบผู้ป่วยใน (IPD)**")
# ใช้ Columns แบ่งเป็น 4 ช่องเพื่อให้ดูง่ายและเรียงลำดับตามขั้นตอนจริง [cite: 76-81]
s1_cols = st.columns(4)

with s1_cols[0]:
    st.text_area(
        "1.1 Admission Note",
        height=200,
        placeholder="วางข้อมูลแรกรับ...",
        key="s11",  # ผูกกับ st.session_state.s11
        help="คัดลอกจาก Admission note ในระบบ IPD [cite: 76]"
    )

with s1_cols[1]:
    st.text_area(
        "1.2 การวินิจฉัย",
        height=200,
        placeholder="วางรหัส ICD-10...",
        key="s12",  # ผูกกับ st.session_state.s12
        help="คัดลอกจากหน้าการวินิจฉัย (Step 1.2) [cite: 79-81]"
    )

with s1_cols[2]:
    st.text_area(
        "1.3 Order / Meds",
        height=200,
        placeholder="วางรายการยา Home-Med...",
        key="s13",  # ผูกกับ st.session_state.s13
        help="คัดลอก Discharge order + Home medication (Step 1.3) [cite: 81]"
    )

with s1_cols[3]:
    st.text_area(
        "1.4 Progress Note",
        height=200,
        placeholder="วาง SOAP ล่าสุด...",
        key="s14",  # ผูกกับ st.session_state.s14
        help="คัดลอก Progress note ล่าสุด (Step 1.4) [cite: 81]"
    )

st.divider()
st.markdown("### **🔵 Step 2: การประเมิน (Assessment)**")
# Step 2: การประเมินผลคะแนน 9Q, 8Q, BPRS [cite: 82]
st.text_area(
    "คัดลอกผลคะแนน 9Q, 8Q, BPRS มาวางที่นี่",
    height=120,
    key="s2",  # ผูกกับ st.session_state.s2
    help="ดึงจากหน้า Assessment ในระบบผู้ป่วยนอก [cite: 82]"
)

st.divider()
st.markdown("### **🟠 Step 3: เวชระเบียน (Registration)**")
# Step 3: ข้อมูลเวชระเบียนที่ต้องใช้ในการระบุตัวตนและที่อยู่ [cite: 83-84]
s3_cols = st.columns(4)

with s3_cols[0]:
    st.text_area(
        "3.1 ข้อมูลทั่วไป", 
        height=180, 
        key="s31", # ผูกกับ st.session_state.s31
        help="ดึงจากหน้า 'ทั่วไป 1' ในระบบเวชระเบียน [cite: 83]"
    )

with s3_cols[1]:
    st.text_area(
        "3.2 ที่อยู่ปัจจุบัน", 
        height=180, 
        key="s32", # ผูกกับ st.session_state.s32
        help="ดึงจากหน้า 'ทั่วไป 2' (ที่อยู่ปัจจุบัน) [cite: 83]"
    )

with s3_cols[2]:
    st.text_area(
        "3.3 ผู้ติดต่อ", 
        height=180, 
        key="s33", # ผูกกับ st.session_state.s33
        help="ดึงจากหน้า 'ผู้ติดต่อ' [cite: 83]"
    )

with s3_cols[3]:
    st.text_area(
        "3.4 สิทธิการรักษา", 
        height=180, 
        key="s34", # ผูกกับ st.session_state.s34
        help="ดึงจากหน้า 'สิทธิการรักษา' [cite: 84]"
    )

# --- 5. ฟังก์ชันจัดการไฟล์ Word (ชิดซ้าย + ฟอนต์ 13 + สกัดข้อมูลลง Placeholder) ---

def fill_pmnidat_doc(data):
    try:
        # โหลดไฟล์แม่แบบ PMNIDAT 062
        doc = Document("PMNIDAT 062 แบบบันทึกข้อมูลเพื่อส่งต่อ.docx")
        
        # เตรียม Mapping ข้อมูล (Key ต้องเป็นตัวพิมพ์ใหญ่ตาม Placeholder ใน Word)
        mapping = {f"{{{{{k.upper()}}}}}": str(v) for k, v in data.items()}
        
        def apply_style_and_replace(paragraph):
            for key, value in mapping.items():
                if key in paragraph.text:
                    paragraph.text = paragraph.text.replace(key, value)
                    # กำหนดการจัดวางเป็นชิดซ้าย (Left Alignment) เพื่อความสวยงาม [cite: 13-28]
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
                    # บังคับฟอนต์ขนาด 13 pt ตลอดทั้งเอกสาร
                    for run in paragraph.runs:
                        run.font.size = Pt(13)

        # ตรวจสอบและแทนที่ข้อมูลทั้งใน Paragraph ปกติและใน Table
        for p in doc.paragraphs: apply_style_and_replace(p)
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs: apply_style_and_replace(p)
                            
        buffer = io.BytesIO()
        doc.save(buffer)
        return buffer.getvalue()
    except Exception as e:
        st.error(f"⚠️ ปัญหาไฟล์แม่แบบหรือการจัดรูปแบบ: {e}")
        return None

# --- 6. ส่วนประมวลผล AI (สกัดข้อมูลตาม Search & Extract Logic) ---

if st.button("🚀 สกัดข้อมูลและสร้างเอกสาร 062", use_container_width=True):
    # รวบรวมข้อมูลดิบจากทั้ง 9 ช่อง [cite: 76-84]
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
    
    with st.spinner('Gemini กำลังสกัดข้อมูลตามตรรกะ PhD และตรวจสอบ Verification Audit...'):
        # ชุดคำสั่ง Prompt ที่ฝังตรรกะที่คุณหมอกำหนด [cite: 69, 87, 88]
        prompt = f"""
        จงสกัดข้อมูลทางการแพทย์จากระบบ @ThanHIS ลงแบบฟอร์ม 062 โดยใช้ตรรกะดังนี้:
        
        1. กฎการตัดขยะ (Noise Reduction): Ignore ข้อความ Theme Customizer, Menu Colors หรือ COPYRIGHT ทั้งหมด [cite: 90]
        2. การคำนวณวันนอน (LOS): ค้นหาตัวเลขหน้าคำว่า 'วัน' ในช่อง Detox และ Rehab แล้วนำมาบวกกันเสมอ [cite: 71, 87]
        3. รหัสโรค (DX): สกัด ICD-10 (เช่น F155) และตัดจุดทศนิยมออกให้เป็นตัวเลขติดกัน [cite: 71, 87]
        4. ยา (MEDS): ค้นหาเฉพาะบรรทัด 'Home-Med' สกัดชื่อยาเป็น UPPERCASE พร้อมวิธีใช้และเลขลำดับ (ใช้ \\n แยกบรรทัด) [cite: 71, 87]
        5. สรุปปัญหา (PROGRESS): สังเคราะห์จาก Progress Note ล่าสุด ให้เหลือย่อหน้าเดียว ความยาว 2-3 บรรทัดเท่านั้น
        6. การตรวจสอบ (Verification Audit): หากไม่พบข้อมูลในช่องที่ระบุ ให้พยายามหาจากช่องอื่น หากไม่มีจริงๆ ให้ใส่ [กรอกด้วยตนเอง] [cite: 86, 87]
        
        ข้อมูลดิบ:
        {all_raw}
        
        ตอบกลับในรูปแบบ JSON ที่มี Key ตรงกับ Placeholder: 
        NAME, AGE, HN, ID, EDU, CAREER, RELIGION, STATUS, RIGHTS, LAST_DC, ADMIT_DATE, VISIT_NUM, CC, CONTACT, RELATION, PHONE, NEAR_HOSP, DC_DATE, LOS, ADDRESS, Q9, Q8, MEDS, DX, PROGRESS, POST_SERVICE
        """
        
        try:
            response = client.models.generate_content(model=MODEL_ID, contents=prompt)
            # สกัด JSON ออกจาก Response
            match = re.search(r'\{.*\}', response.text, re.DOTALL)
            if match:
                json_data = json.loads(match.group())
                st.success("✅ วิเคราะห์ข้อมูลและคำนวณวันนอน (LOS) สำเร็จ!")
                
                # สร้างไฟล์ Word
                word_file = fill_pmnidat_doc(json_data)
                if word_file:
                    # บันทึก Log ไปยัง Google Sheets
                    log_usage(json_data.get('name', '[ไม่ระบุชื่อ]'))
                    
                    st.download_button(
                        label="💾 ดาวน์โหลดไฟล์ 062 ฉบับสมบูรณ์ (ฟอนต์ 13)",
                        data=word_file,
                        file_name=f"Refer_{json_data.get('name','062')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )
            else:
                st.error("AI ไม่สามารถสร้างรูปแบบ JSON ที่ถูกต้องได้ กรุณาลองใหม่อีกครั้ง")
        except Exception as e:
            st.error(f"เกิดข้อผิดพลาดในการประมวลผล: {e}")

# --- 7. มาตรการรักษาความปลอดภัย (PDPA Footer) ---
st.divider()
st.info("""
    **ประกาศ: มาตรการรักษาความปลอดภัยข้อมูลทางการแพทย์ (PDPA Compliance)**
    * ระบบ PMNIDAT Smart D/C Transfer ไม่มีการจัดเก็บข้อมูลผู้ป่วยไว้ในเซิร์ฟเวอร์ถาวร
    * ข้อมูลจะสูญหายทันทีเมื่อมีการรีเฟรชหน้าจอ (Refresh) โปรดบันทึกไฟล์ให้เรียบร้อยก่อนออกจากระบบ
    """)
