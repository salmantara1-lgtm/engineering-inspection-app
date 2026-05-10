import streamlit as st
from datetime import datetime
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches
from io import BytesIO
import json
import os

# --- الإعدادات الأساسية للصفحة ---
st.set_page_config(page_title="نظام الاستلامات الهندسية", page_icon="🏗️", layout="wide")

# ==========================================
# نظام قاعدة البيانات المبسط (لتتبع أرقام الاستلامات التلقائية)
# ==========================================
DB_FILE = "inspections_db.json"

def load_db():
    if os.path.exists(DB_FILE):
        with open(DB_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    return {}

def save_to_db(project, insp_type):
    db = load_db()
    if project not in db:
        db[project] = {}
    if insp_type not in db[project]:
        db[project][insp_type] = 0
    db[project][insp_type] += 1
    
    with open(DB_FILE, "w", encoding="utf-8") as f:
        json.dump(db, f, ensure_ascii=False, indent=4)
        
    return db[project][insp_type]

def get_next_number(project, insp_type):
    if not project:
        return 1
    db = load_db()
    return db.get(project, {}).get(insp_type, 0) + 1

# ==========================================
# قاعدة بيانات الاستلامات 
# ==========================================
INSPECTIONS_DATA = {
    "أعمال الحفر": [
        "مراجعة الميزانية الشبكية للأرض الطبيعية قبل الحفر.",
        "مراجعة إحداثيات حدود الحفر مع حدود المبنى طبقا لتقرير الجسات.",
        "مراجعة منسوب التأسيس مع اللوحات وعمق الحفر من تقرير التربة.",
        "مراجعة تطهير قاع وجوانب الحفر واستواء الحفر.",
        "التأكد من أن نوع تربة قاع الحفر مطابقة لتقرير التربة."
    ],
    "أعمال الإحلال": [
        "التأكد من منسوب طبقة الإحلال (منسوب التأسيس).",
        "التأكد من سمك طبقة الإحلال طبقاً لتوصيات الاستشاري.",
        "التأكد من طريقة الدمك المستخدمة طبقا للمواصفات.",
        "التأكد من نتيجة اختبار الدمك قبل البدء في الطبقة التالية."
    ],
    "فحص واستلام الخنزيرة": [
        "أبعاد الخنزيرة أكبر من أبعاد قاع الحفر بمسافة تمنع تأثرها.",
        "استقامة أضلاع الخنزيرة بخيط الشد.",
        "أفقية الأضلاع بميزان المياه.",
        "مراجعة الزوايا المحصورة بين الأضلاع (نظرية فيثاغورث).",
        "مراجعة المحاور على الخنزيرة."
    ]
}

# --- دالة إنشاء ملف Word ---
def create_docx(data):
    doc = Document()
    
    heading = doc.add_heading(f"نموذج فحص واستلام - {data['inspection_type']}", level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # البيانات الأساسية (تمت إضافة الاستشاري والمهندس)
    doc.add_heading('البيانات الأساسية:', level=2)
    p = doc.add_paragraph()
    p.add_run(f"المشروع: ").bold = True
    p.add_run(f"{data['project']}\n")
    p.add_run(f"الاستشاري: ").bold = True
    p.add_run(f"{data['consultant']}\n")
    p.add_run(f"رقم طلب الاستلام: ").bold = True
    p.add_run(f"{data['inspection_no']}\n")
    p.add_run(f"التاريخ: ").bold = True
    p.add_run(f"{data['date']}\n")
    p.add_run(f"المهندس المستلم: ").bold = True
    p.add_run(f"{data['engineer']}\n")
    p.add_run(f"مندوب المقاول: ").bold = True
    p.add_run(f"{data['contractor']}\n")

    # جدول المراجعة العامة
    doc.add_heading('مراجعة النقاط قبل الاستلام:', level=2)
    t1 = doc.add_table(rows=1, cols=2)
    t1.style = 'Table Grid'
    t1.rows[0].cells[0].text = 'البند'
    t1.rows[0].cells[1].text = 'الحالة'
    for key, value in data['general_checks'].items():
        row = t1.add_row().cells
        row[0].text = key
        row[1].text = "تم" if value else "غير مطبق"

    # جدول الفحص الفني (3 أعمدة: البند - النتيجة - الملاحظات)
    doc.add_heading('بنود المراجعة الفنية:', level=2)
    t2 = doc.add_table(rows=1, cols=3)
    t2.style = 'Table Grid'
    t2.rows[0].cells[0].text = 'البند'
    t2.rows[0].cells[1].text = 'النتيجة'
    t2.rows[0].cells[2].text = 'الملاحظات'
    
    for key, value in data['tech_checks'].items():
        row = t2.add_row().cells
        row[0].text = key
        row[1].text = value['result']
        row[2].text = value['note'] # الملاحظة الخاصة بالبند

    # التقييم النهائي
    doc.add_heading('النتيجة والملاحظات العامة:', level=2)
    doc.add_paragraph(f"التقييم النهائي: {data['status']}")
    doc.add_paragraph(f"ملاحظات عامة: {data['comments']}")

    # إضافة الصور
    if data['photos']:
        doc.add_page_break()
        doc.add_heading('توثيق الموقع (الصور):', level=1)
        for idx, photo_data in enumerate(data['photos']):
            try:
                image_stream = BytesIO(photo_data['image'].getvalue())
                doc.add_picture(image_stream, width=Inches(4.5))
                p_desc = doc.add_paragraph()
                p_desc.add_run(f"صورة ({idx+1}): ").bold = True
                p_desc.add_run(photo_data['desc'])
                doc.add_paragraph("\n")
            except Exception:
                pass

    target = BytesIO()
    doc.save(target)
    return target.getvalue()

# --- واجهة المستخدم (UI) ---
st.title("🏗️ نظام الاستلامات الهندسية الاحترافي")

# القائمة الجانبية (تمت إضافة حقول الاستشاري والمهندس)
with st.sidebar:
    st.header("📋 بيانات الطلب")
    project_name = st.text_input("اسم المشروع", "مشروع ادوار الحداثة")
    consultant_name = st.text_input("اسم الاستشاري", "محرم باخوم")
    engineer_name = st.text_input("المهندس المستلم (اسمك)")
    contractor_name = st.text_input("اسم المقاول / المندوب")
    date_str = st.date_input("تاريخ الفحص", datetime.now()).strftime('%Y-%m-%d')

# اختيار نوع الاستلام
st.markdown("### 🔍 اختر مرحلة الاستلام المراد فحصها:")
selected_inspection = st.selectbox("قائمة الاستلامات المتاحة:", list(INSPECTIONS_DATA.keys()))

# حساب وعرض رقم الاستلام التلقائي
auto_insp_num = get_next_number(project_name, selected_inspection)
st.info(f"🔢 رقم الاستلام التلقائي المقترح لهذه المرحلة في هذا المشروع هو: **{auto_insp_num}**")

st.markdown("---")

tab1, tab2, tab3, tab4 = st.tabs(["📋 المراجعة الإدارية", "🛠️ الفحص الفني (والملاحظات)", "📸 الصور", "✅ الاعتماد"])

with tab1:
    st.subheader("تحقق ما قبل الاستلام")
    g_checks = {
        "اعتماد المواد المستخدمة في البند": st.checkbox("اعتماد المواد", value=True),
        "اعتماد رسومات الورشة (Shop Drawings)": st.checkbox("اعتماد المخططات التنفيذية", value=True),
        "اعتماد طريقة التنفيذ (Method Statement)": st.checkbox("اعتماد طريقة التنفيذ"),
        "توافر شروط السلامة المهنية بالموقع": st.checkbox("شروط السلامة والأمان")
    }

with tab2:
    st.subheader(f"بنود فحص: {selected_inspection}")
    current_tech_items = INSPECTIONS_DATA[selected_inspection]
    
    t_results = {}
    for item in current_tech_items:
        st.markdown(f"🔹 **{item}**")
        # النتيجة
        res = st.radio("النتيجة:", ["مطابق", "غير مطابق", "غير مطبق"], horizontal=True, key=f"r_{item}", label_visibility="collapsed")
        # حقل الملاحظة الخاص بالبند
        note = st.text_input("ملاحظة على البند:", key=f"n_{item}", placeholder="اكتب ملاحظة إن وجدت...", label_visibility="collapsed")
        
        # حفظ النتيجة والملاحظة معاً كقاموس
        t_results[item] = {"result": res, "note": note if note else "-"}
        st.markdown("---")

with tab3:
    st.subheader("إرفاق الصور والملاحظات (اختياري)")
    captured_photos = []
    col1, col2 = st.columns(2)
    for i in range(1, 5): # قللنا العدد لـ 4 للشكل العام، يمكنك زيادته لـ 8
        current_col = col1 if i % 2 != 0 else col2
        with current_col:
            with st.expander(f"📷 إرفاق صورة رقم {i}"):
                img_file = st.file_uploader(f"اختر أو التقط صورة", type=['png', 'jpg', 'jpeg'], key=f"file_{i}")
                img_desc = st.text_input(f"ملاحظة للصورة:", key=f"desc_{i}")
                if img_file is not None:
                    st.image(img_file, use_column_width=True)
                    captured_photos.append({"image": img_file, "desc": img_desc if img_desc else "-"})

with tab4:
    st.subheader("التقييم النهائي واعتماد التقرير")
    
    # خانة لتعديل رقم الاستلام يدوياً لو أراد المهندس تجاوزه الترقيم التلقائي
    final_insp_no = st.text_input("رقم طلب الاستلام لاعتماده بالتقرير:", value=str(auto_insp_num))
    
    final_status = st.radio("القرار النهائي للمهندس:", [
        "1 - لا توجد ملاحظات (يعتمد البند)",
        "2 - يعتمد مع تنفيذ الملاحظات",
        "3 - يعاد التقديم بعد إتمام الملاحظات",
        "4 - مرفوض"
    ])
    notes = st.text_area("ملاحظات عامة / أسباب الرفض (إن وجدت)")
    
    st.markdown("---")
    
    if st.button("✅ حفظ وإصدار التقرير (Word)", use_container_width=True):
        if not project_name or not engineer_name:
            st.error("⚠️ يرجى كتابة 'اسم المشروع' و 'اسم المهندس المستلم' في القائمة الجانبية.")
        else:
            # تحديث قاعدة البيانات وحفظ الرقم التلقائي الجديد
            saved_no = save_to_db(project_name, selected_inspection)
            
            form_data = {
                "inspection_type": selected_inspection,
                "project": project_name,
                "consultant": consultant_name,
                "engineer": engineer_name,
                "contractor": contractor_name,
                "inspection_no": final_insp_no,
                "date": date_str,
                "general_checks": g_checks,
                "tech_checks": t_results,
                "photos": captured_photos,
                "status": final_status,
                "comments": notes
            }
            
            with st.spinner("جاري إنشاء التقرير..."):
                doc_bytes = create_docx(form_data)
                
            st.success(f"🎉 تم حفظ الاستلام برقم {saved_no} وتجهيز التقرير بنجاح!")
            st.download_button(
                label="📥 اضغط هنا لتحميل تقرير Word الشامل",
                data=doc_bytes,
                file_name=f"Inspection_{selected_inspection}_{final_insp_no}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
