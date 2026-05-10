import streamlit as st
from datetime import datetime
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from io import BytesIO
import json
import os

# --- الإعدادات الأساسية للصفحة ---
st.set_page_config(page_title="نظام استلامات العظم - أدوار الحداثة", page_icon="🏗️", layout="wide")

# ==========================================
# 🎨 تحويل واجهة التطبيق بالكامل للغة العربية (RTL)
# ==========================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap');

/* تطبيق الخط العربي والاتجاه على كامل التطبيق */
html, body, [class*="css"], .stApp {
    direction: rtl;
    text-align: right;
    font-family: 'Cairo', sans-serif !important;
}

/* محاذاة النصوص داخل جميع الحقول والأزرار لليمين */
p, div, input, label, h1, h2, h3, h4, h5, h6, span, li, textarea {
    text-align: right !important;
}

/* تعديل القائمة الجانبية لتكون ملائمة للغة العربية */
[data-testid="stSidebar"] {
    direction: rtl;
}

/* تعديل عرض أزرار الاختيار (Radio Buttons) */
div[role="radiogroup"] {
    flex-direction: row-reverse;
}

/* تحسين شكل الفواصل البصرية */
hr {
    border-top: 2px solid #f0f2f6;
}
</style>
""", unsafe_allow_html=True)

# ==========================================
# نظام قواعد البيانات (للمشاريع والترقيم التراكمي)
# ==========================================
DB_FILE = "projects_bone_works_db.json"

def load_db():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except: return {"projects": [], "counters": {}}
    return {"projects": [], "counters": {}}

def save_db(db):
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(db, f, ensure_ascii=False, indent=4)

def get_next_visit_id(project):
    db = load_db()
    return db["counters"].get(project, 0) + 1

def update_visit_counter(project):
    db = load_db()
    if project not in db["projects"]:
        db["projects"].append(project)
    if project not in db["counters"]:
        db["counters"][project] = 0
    db["counters"][project] += 1
    save_db(db)

# ==========================================
# قاعدة بيانات نماذج أعمال العظم (أول 5 نماذج بدقة تامة وبدون اختصار)
# ==========================================
INSPECTIONS_DATA = {
    "استلام اعمال الحفر": [
        "مراجعة الميزانية الشبكية للأرض الطبيعية قبل الحفر.",
        "مراجعة إحداثيات حدود الحفر مع حدود المبنى طبقا لتقرير الجسات.",
        "مراجعة منسوب التأسيس مع اللوحات وعمق الحفر من تقرير التربة وتحديد أماكن الروبير.",
        "مراجعة تطهير قاع وجوانب الحفر واستواء الحفر.",
        "التأكد من أن نوع تربة قاع الحفر مطابقة لتقرير التربة ولا توجد اختلافات."
    ],
    "استلام اعمال الاحلال": [
        "التأكد من منسوب طبقة الإحلال (منسوب التأسيس).",
        "التأكد من سمك طبقة الإحلال طبقاً لتوصيات الاستشاري.",
        "التأكد من طريقة الدمك المستخدمة (ميكانيكي / ستاتيكي / هزاز) طبقا للمواصفات.",
        "التأكد من الوزن المستخدم في الدمك.",
        "التأكد من نتيجة اختبار الدمك (Sand Cone) قبل البدء في الطبقة التالية.",
        "مطابقة مواد طبقة الإحلال للعينة المعتمدة."
    ],
    "استلام نجارة القواعد العادية": [
        "التأكد من نظافة منسوب التأسيس مع تمام دمك السطح ورشه بالماء.",
        "توقيع أركان القواعد بجهاز مساحي (Total Station) ومطابقتها مع الرسومات.",
        "مطابقة المحاور الإنشائية وصحة توقيع زواياها.",
        "مراجعة أبعاد القواعد وارتفاعاتها وحدود جوانبها.",
        "التقفيل الجيد لجوانب القواعد مع بعضها وتسديد الفتحات.",
        "مراجعة التقويات والتأكد من إتمامها بطريقة صحيحة.",
        "التأكد من وضع شرب صب القواعد (منسوب الظهر) بميزان القامة."
    ],
    "استلام نجارة القواعد الخرسانية المسلحة والسملات": [
        "مراجعة أبعاد القواعد المسلحة والسملات وتربيعها.",
        "التأكد من نظافة القوالب الخشبية من الداخل.",
        "التأكد من دهان القوالب الخشبية بمادة مانعة للالتصاق.",
        "مراجعة أماكن الفتحات ومسارات الصحي والكهرباء.",
        "مراجعة التقويات (الزراجين والشكالات والأحزمة) لمنع الحركة أثناء الصب.",
        "مراجعة شرب الصب وتحديد الارتفاعات المطلوبة للخرسانة."
    ],
    "استلام حديد تسليح الاساسات": [
        "التأكد من أقطار وأعداد وأطوال حديد التسليح ومطابقتها للمخططات.",
        "التأكد من أطوال التراكب (الوصلات) وأماكنها الإنشائية الصحيحة.",
        "مراجعة تربيط الحديد بالسلك المجلفن بشكل جيد.",
        "مراجعة تركيب البسكويت (الغطاء الخرساني) بالسمك المطلوب وتوزيعه.",
        "التأكد من نظافة حديد التسليح من الصدأ أو الزيوت أو الشحوم.",
        "التأكد من تثبيت أشاير الأعمدة في أماكنها الصحيحة والمحافظة على رأسيتها."
    ]
}

# --- دوال تنسيق الوورد (RTL للغة العربية) ---
def set_rtl(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p_pr = paragraph._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    p_pr.append(bidi)

def set_table_rtl(table):
    tblPr = table._tbl.tblPr
    bidiVisual = OxmlElement('w:bidiVisual')
    bidiVisual.set(qn('w:val'), '1')
    tblPr.append(bidiVisual)

# --- دالة التقرير (Word) ---
def create_docx(data):
    doc = Document()
    
    header = doc.add_heading(f"نموذج {data['inspection_type']}", level=1)
    set_rtl(header)
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # البيانات الأساسية
    h2 = doc.add_heading('البيانات الأساسية للمشروع:', level=2)
    set_rtl(h2)
    p = doc.add_paragraph()
    set_rtl(p)
    p.add_run(f"اسم المشروع: {data['project']}\n").bold = True
    p.add_run(f"الاستشاري: {data['consultant']}\n")
    p.add_run(f"ممثل الاستشاري: {data['engineer']}\n")
    p.add_run(f"ممثل المقاول: {data['contractor']}\n")
    p.add_run(f"رقم الزيارة (الاستلام): {data['inspection_no']}\n")
    p.add_run(f"تاريخ الفحص: {data['date']}\n")

    # جدول الإدارة (مرتب من اليمين لليسار بسبب RTL)
    h3 = doc.add_heading('مراجعة النقاط الإدارية:', level=2)
    set_rtl(h3)
    t1 = doc.add_table(rows=1, cols=2)
    t1.style = 'Table Grid'
    set_table_rtl(t1)
    t1.rows[0].cells[0].text = 'بند المراجعة'
    t1.rows[0].cells[1].text = 'الحالة'
    for key, value in data['general_checks'].items():
        row = t1.add_row().cells
        row[0].text = key
        row[1].text = "تم" if value else "غير مطبق"

    # جدول الفحص الفني (مرتب من اليمين لليسار: البند - النتيجة - الملاحظات)
    h3_tech = doc.add_heading('بنود المراجعة الفنية:', level=2)
    set_rtl(h3_tech)
    t2 = doc.add_table(rows=1, cols=3)
    t2.style = 'Table Grid'
    set_table_rtl(t2)
    t2.rows[0].cells[0].text = 'بند الفحص'
    t2.rows[0].cells[1].text = 'النتيجة'
    t2.rows[0].cells[2].text = 'الملاحظات'
    
    for key, value in data['tech_checks'].items():
        row = t2.add_row().cells
        row[0].text = key
        row[1].text = value['result']
        row[2].text = value['note']

    # التقييم
    h4 = doc.add_heading('القرار النهائي وملاحظات المهندس:', level=2)
    set_rtl(h4)
    p_final = doc.add_paragraph()
    set_rtl(p_final)
    p_final.add_run(f"حالة الاعتماد: {data['status']}\n").bold = True
    p_final.add_run(f"توجيهات عامة: {data['comments']}")

    # إدراج الصور (حتى 10 صور)
    if data['photos']:
        doc.add_page_break()
        h_img = doc.add_heading('التوثيق الفوتوغرافي للموقع:', level=1)
        set_rtl(h_img)
        for idx, photo in enumerate(data['photos']):
            try:
                doc.add_picture(BytesIO(photo['image'].getvalue()), width=Inches(4.5))
                p_img = doc.add_paragraph(f"صورة ({idx+1}): {photo['desc']}")
                set_rtl(p_img)
            except: pass

    target = BytesIO()
    doc.save(target)
    return target.getvalue()

# --- واجهة الاستخدام (Main App UI) ---
db_data = load_db()

with st.sidebar:
    st.header("⚙️ بيانات المشروع (حقول إلزامية)")
    
    proj_options = db_data["projects"] + ["+ إضافة مشروع جديد..."]
    selected_proj = st.selectbox("اختر المشروع:", proj_options)
    
    if selected_proj == "+ إضافة مشروع جديد...":
        project_name = st.text_input("أدخل اسم المشروع الجديد:")
    else:
        project_name = selected_proj
        
    consultant = st.text_input("الاستشاري", "شركة ادوار الحداثة للاستشارات الهندسية")
    engineer = st.text_input("ممثل الاستشاري", "م.سلمان تاره")
    contractor = st.text_input("ممثل المقاول", placeholder="اكتب اسم ممثل المقاول هنا")
    date_val = st.date_input("تاريخ الفحص", datetime.now()).strftime('%Y-%m-%d')

st.subheader("🎯 اختيار نموذج استلام العظم")
stage = st.selectbox("المرحلة المراد استلامها:", list(INSPECTIONS_DATA.keys()))

next_visit_no = get_next_visit_id(project_name) if project_name else 1
st.info(f"🔢 زيارة المشروع رقم: **{next_visit_no}**")

tab1, tab2, tab3 = st.tabs(["📝 بنود المراجعة", "📸 إرفاق الصور (10 صور)", "📤 إصدار التقرير"])

with tab1:
    st.write("### البنود الإدارية")
    gen_checks = {k: st.checkbox(k, value=True) for k in ["التأكد من اعتماد المواد", "التأكد من اعتماد المخططات التنفيذية", "التأكد من اعتماد طريقة التنفيذ", "توافر شروط السلامة المهنية في الموقع"]}
    
    st.divider()
    st.write("### الفحص الفني الميداني")
    tech_results = {}
    for item in INSPECTIONS_DATA[stage]:
        st.markdown(f"**🔹 {item}**")
        res = st.radio("النتيجة:", ["مطابق", "غير مطابق", "غير مطبق"], horizontal=True, key=f"res_{item}", label_visibility="collapsed")
        note = st.text_input("ملاحظة المهندس:", key=f"note_{item}", placeholder="أضف تعليقاً على هذا البند (اختياري)...")
        tech_results[item] = {"result": res, "note": note if note else "-"}
        st.markdown("---")

with tab2:
    st.write("### توثيق الموقع (اختياري - حتى 10 صور)")
    photos_list = []
    for i in range(1, 11):
        with st.expander(f"إرفاق صورة رقم {i}"):
            f = st.file_uploader(f"التقط أو ارفع صورة {i}", type=['jpg','jpeg','png'], key=f"img_{i}")
            d = st.text_input(f"وصف الصورة {i}", key=f"d_{i}")
            if f:
                st.image(f, use_column_width=True)
                photos_list.append({"image": f, "desc": d if d else "-"})

with tab3:
    st.write("### الاعتماد النهائي")
    final_no = st.text_input("رقم الزيارة للتثبيت في التقرير:", value=str(next_visit_no))
    decision = st.radio("القرار النهائي للاستلام:", ["مقبول", "مقبول مع ملاحظات يتم تلافيها", "مرفوض ويعاد التقديم"])
    general_notes = st.text_area("توجيهات عامة لممثل المقاول:")

    if st.button("✅ تأكيد البيانات وإصدار التقرير", use_container_width=True):
        if not project_name or not consultant or not engineer or not contractor:
            st.error("⚠️ يرجى التأكد من تعبئة جميع بيانات المشروع في القائمة الجانبية (اسم المشروع، الاستشاري، ممثل الاستشاري، ممثل المقاول) فهي حقول إلزامية.")
        else:
            update_visit_counter(project_name)
            final_data = {
                "inspection_type": stage, "project": project_name, "consultant": consultant,
                "engineer": engineer, "contractor": contractor, "inspection_no": final_no,
                "date": date_val, "general_checks": gen_checks, "tech_checks": tech_results,
                "photos": photos_list, "status": decision, "comments": general_notes
            }
            report = create_docx(final_data)
            
            file_name = f"{stage} - {project_name}.docx"
            
            st.success(f"🎉 تم حفظ بيانات الزيارة بنجاح للمشروع: {project_name}")
            st.download_button("📥 تحميل التقرير (Word)", data=report, file_name=file_name, use_container_width=True)
