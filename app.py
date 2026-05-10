import streamlit as st
from datetime import datetime
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches
from io import BytesIO
import json
import os

# --- الإعدادات الأساسية للصفحة ---
st.set_page_config(page_title="نظام ACE للاستلامات الهندسية", page_icon="🏗️", layout="centered")

# ==========================================
# نظام الترقيم التلقائي (قاعدة البيانات)
# ==========================================
DB_FILE = "inspections_history.json"

def load_db():
    if os.path.exists(DB_FILE):
        try:
            with open(DB_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except: return {}
    return {}

def get_next_id(project, stage):
    db = load_db()
    current_count = db.get(project, {}).get(stage, 0)
    return current_count + 1

def update_db(project, stage):
    db = load_db()
    if project not in db: db[project] = {}
    db[project][stage] = db[project].get(stage, 0) + 1
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(db, f, ensure_ascii=False, indent=4)

# ==========================================
# قاعدة بيانات الاستلامات (12+ نموذج)
# ==========================================
INSPECTIONS_DATA = {
    "أعمال الحفر": ["مراجعة الميزانية الشبكية للأرض", "مطابقة الإحداثيات مع حدود المبنى", "مراجعة منسوب التأسيس مع الروبير", "تطهير واستواء قاع الحفر", "مطابقة التربة لتقرير الجسات"],
    "أعمال الإحلال": ["منسوب طبقة الإحلال", "سمك الطبقات وتوصيات الاستشاري", "طريقة الدمك والوزن المستخدم", "نتيجة اختبار الدمك (Sand Cone)"],
    "فحص واستلام الخنزيرة": ["أبعاد الخنزيرة أكبر من الحفر", "استقامة الأضلاع بالخيط", "تقوية الأضلاع بالخوابير", "أفقية الأضلاع بميزان المياه", "صحة الزوايا (فيثاغورث)", "توقيع المحاور"],
    "القواعد العادية": ["نظافة قاع الحفر ورشه", "توقيع الأركان بالـ Total Station", "أبعاد القواعد وارتفاعاتها", "تسديد الفتحات والتقوية", "منسوب ظهر القواعد"],
    "القواعد المسلحة والسملات": ["الأبعاد والتربيع", "نظافة النجارة والدهان بالزيت", "أماكن الفتحات ومسارات الصحي", "التقويات والزراجين"],
    "حديد تسليح الأساسات": ["الأقطار والأعداد حسب المخطط", "أطوال الوصلات وأماكنها", "التربيط بسلك مجلفن", "تركيب البسكويت (الغطاء الخرساني)", "نظافة الحديد من الصدأ"],
    "نجارة الأعمدة": ["قطاعات الأعمدة وأماكنها", "رأسية الأعمدة بميزان الخيط", "متانة الأحزمة والتقويات", "منسوب نهاية الصب"],
    "تسليح الأعمدة والحوائط": ["عدد وأقطار الأسياخ", "أقفال الكانات وتبادلها", "تكثيف الكانات في الثلثين", "طول إشارة العمود"],
    "نجارة الأسقف (تحت السقف)": ["توقيع المحاور والكمرات", "أفقية السطح وقوة الشدة", "نظافة السطح الخشبي", "أبعاد السقوط للكمرات"],
    "تسليح الأسقف": ["حديد الفرش والغطاء", "حديد الكمرات المكسح والعدل", "وصلات التراكب وأماكنها", "بسكويت البلاطات والكمرات"],
    "أعمال العزل للاساسات": ["نظافة الأسطح ورقبة الزجاجة", "طبقة البرايمر والتغطية", "لحام رولات العزل والركوب", "حماية العزل بعد التنفيذ"],
    "أعمال الردم": ["جودة تربة الردم الموردة", "الردم على طبقات (25-30 سم)", "الرش بالماء والدمك جيداً", "اختبارات التربة لكل طبقة"]
}

# --- دالة إنشاء ملف Word (الترتيب: الملاحظات -> النتيجة -> البند) ---
def create_docx(data):
    doc = Document()
    
    header = doc.add_heading(f"نموذج فحص واستلام - {data['inspection_type']}", level=1)
    header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # البيانات الأساسية
    doc.add_heading('البيانات الأساسية:', level=2)
    p = doc.add_paragraph()
    p.add_run(f"المشروع: {data['project']}\n").bold = True
    p.add_run(f"الاستشاري: {data['consultant']}\n")
    p.add_run(f"المهندس المستلم: {data['engineer']}\n")
    p.add_run(f"مندوب المقاول: {data['contractor']}\n")
    p.add_run(f"رقم الطلب: {data['inspection_no']}\n")
    p.add_run(f"التاريخ: {data['date']}\n")

    # جدول الإدارة (عكس: الحالة -> البند)
    doc.add_heading('مراجعة النقاط الإدارية:', level=2)
    t1 = doc.add_table(rows=1, cols=2)
    t1.style = 'Table Grid'
    t1.rows[0].cells[0].text = 'الحالة'
    t1.rows[0].cells[1].text = 'البند'
    for key, value in data['general_checks'].items():
        row = t1.add_row().cells
        row[0].text = "تم" if value else "غير مطبق"
        row[1].text = key

    # جدول الفحص الفني (عكس: الملاحظات -> النتيجة -> البند)
    doc.add_heading('بنود المراجعة الفنية:', level=2)
    t2 = doc.add_table(rows=1, cols=3)
    t2.style = 'Table Grid'
    t2.rows[0].cells[0].text = 'الملاحظات'
    t2.rows[0].cells[1].text = 'النتيجة'
    t2.rows[0].cells[2].text = 'بند المراجعة'
    for key, value in data['tech_checks'].items():
        row = t2.add_row().cells
        row[0].text = value['note']
        row[1].text = value['result']
        row[2].text = key

    # التقييم النهائي
    doc.add_heading('النتيجة النهائية:', level=2)
    doc.add_paragraph(f"القرار: {data['status']}")
    doc.add_paragraph(f"ملاحظات عامة: {data['comments']}")

    # الصور
    if data['photos']:
        doc.add_page_break()
        doc.add_heading('التوثيق الفوتوغرافي:', level=1)
        for idx, photo in enumerate(data['photos']):
            doc.add_picture(BytesIO(photo['image'].getvalue()), width=Inches(4.5))
            doc.add_paragraph(f"صورة ({idx+1}): {photo['desc']}")

    target = BytesIO()
    doc.save(target)
    return target.getvalue()

# --- واجهة المستخدم ---
st.title("🏗️ نظام الاستلامات الهندسية الشامل")

with st.sidebar:
    st.header("⚙️ إعدادات الطلب")
    proj = st.text_input("اسم المشروع", "مشروع ادوار الحداثة")
    cons = st.text_input("الاستشاري", "محرم باخوم")
    eng = st.text_input("المهندس المستلم", "م. سلمان تارة")
    cont = st.text_input("مندوب المقاول")
    dt = st.date_input("التاريخ", datetime.now()).strftime('%Y-%m-%d')

stage = st.selectbox("🎯 اختر مرحلة الاستلام:", list(INSPECTIONS_DATA.keys()))
next_no = get_next_id(proj, stage)
st.success(f"الرقم التلقائي القادم لهذه المرحلة: **{next_no}**")

tab1, tab2, tab3 = st.tabs(["📋 بنود الفحص", "📸 التوثيق الصوري", "💾 الاعتماد والإصدار"])

with tab1:
    st.subheader("تحقق ما قبل الاستلام (إداري)")
    g_checks = {k: st.checkbox(k, value=True) for k in ["اعتماد المواد", "اعتماد المخططات", "طريقة التنفيذ", "السلامة المهنية"]}
    
    st.markdown("---")
    st.subheader("الفحص الفني الدقيق")
    t_results = {}
    
    # التعديل الجديد ليكون متوافقاً ومريحاً جداً لشاشة الجوال
    for item in INSPECTIONS_DATA[stage]:
        st.markdown(f"🔹 **{item}**") # اسم البند يظهر بوضوح وبخط عريض
        
        # اختيار النتيجة بشكل أفقي ومريح للأصابع
        res = st.radio("النتيجة:", ["مطابق", "غير مطابق", "غير مطبق"], horizontal=True, key=f"r_{item}", label_visibility="collapsed")
        
        # خانة الملاحظات أسفل النتيجة مباشرة
        note = st.text_input("الملاحظات:", placeholder="أضف ملاحظة للبند (اختياري)...", key=f"n_{item}")
        
        t_results[item] = {"result": res, "note": note if note else "-"}
        st.divider() # خط فاصل رمادي أنيق بين كل بند والآخر

with tab2:
    st.subheader("إرفاق حتى 8 صور مع الوصف")
    photos = []
    for i in range(1, 9):
        with st.expander(f"📷 إرفاق صورة رقم {i}"):
            file = st.file_uploader(f"ارفع أو التقط صورة", type=['jpg','png','jpeg'], key=f"f{i}")
            desc = st.text_input(f"وصف الصورة (اختياري)", key=f"d{i}")
            if file: 
                st.image(file, use_column_width=True)
                photos.append({"image": file, "desc": desc if desc else "-"})

with tab3:
    st.subheader("الاعتماد النهائي")
    final_no = st.text_input("رقم طلب الاستلام النهائي بالتقرير", value=str(next_no))
    stat = st.radio("القرار:", ["1 - يعتمد", "2 - يعتمد بملاحظات", "3 - يعاد التقديم", "4 - مرفوض"])
    comm = st.text_area("تعليق المهندس العام")
    
    if st.button("✅ حفظ البيانات وإصدار التقرير", use_container_width=True):
        if not proj or not stage:
            st.error("يرجى إدخال اسم المشروع أولاً")
        else:
            update_db(proj, stage)
            data = {"inspection_type": stage, "project": proj, "consultant": cons, "engineer": eng, "contractor": cont, "inspection_no": final_no, "date": dt, "general_checks": g_checks, "tech_checks": t_results, "photos": photos, "status": stat, "comments": comm}
            doc_out = create_docx(data)
            
            # التعديل الجديد: اسم الملف يشمل اسم الفحص واسم المشروع
            safe_proj_name = proj.replace("/", "-") # لتجنب مشاكل الأسماء في الويندوز
            file_name = f"{stage} - {safe_proj_name}.docx"
            
            st.download_button(
                label="📥 تحميل التقرير (Word)", 
                data=doc_out, 
                file_name=file_name, 
                use_container_width=True
            )
