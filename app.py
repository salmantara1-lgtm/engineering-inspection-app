import streamlit as st
from datetime import datetime
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from io import BytesIO

# --- الإعدادات الأساسية للصفحة ---
st.set_page_config(page_title="ACE - نظام الاستلامات الهندسية", page_icon="🏗️", layout="centered")

# --- دالة إنشاء ملف Word ---
def create_docx(data):
    doc = Document()
    
    # عنوان التقرير
    heading = doc.add_heading('نموذج فحص واستلام الخنزيرة (ACE/QA/019-003)', level=1)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # بيانات المشروع الأساسية
    doc.add_heading('البيانات الأساسية:', level=2)
    p = doc.add_paragraph()
    p.add_run(f"المشروع: ").bold = True
    p.add_run(f"{data['project']}\n")
    p.add_run(f"رقم طلب الاستلام: ").bold = True
    p.add_run(f"{data['inspection_no']}\n")
    p.add_run(f"التاريخ: ").bold = True
    p.add_run(f"{data['date']}\n")
    p.add_run(f"مندوب المقاول: ").bold = True
    p.add_run(f"{data['contractor']}\n")

    # جدول بنود المراجعة العامة
    doc.add_heading('مراجعة النقاط قبل الاستلام:', level=2)
    table_general = doc.add_table(rows=1, cols=2)
    table_general.style = 'Table Grid'
    hdr_cells = table_general.rows[0].cells
    hdr_cells[0].text = 'البند'
    hdr_cells[1].text = 'الحالة'
    
    for key, value in data['general_checks'].items():
        row_cells = table_general.add_row().cells
        row_cells[0].text = key
        row_cells[1].text = "تم" if value else "غير مطبق"

    # جدول الفحص الفني
    doc.add_heading('بنود المراجعة الفنية:', level=2)
    table_tech = doc.add_table(rows=1, cols=2)
    table_tech.style = 'Table Grid'
    hdr_cells_tech = table_tech.rows[0].cells
    hdr_cells_tech[0].text = 'بند المراجعة'
    hdr_cells_tech[1].text = 'النتيجة'
    
    for key, value in data['tech_checks'].items():
        row = table_tech.add_row().cells
        row[0].text = key
        row[1].text = value

    # التقييم النهائي والملاحظات
    doc.add_heading('النتيجة والملاحظات:', level=2)
    doc.add_paragraph(f"التقييم النهائي: {data['status']}")
    doc.add_paragraph(f"ملاحظات: {data['comments']}")
    
    # حفظ الملف في الذاكرة
    target = BytesIO()
    doc.save(target)
    return target.getvalue()

# --- واجهة المستخدم (UI) ---
st.title("🏗️ نظام استلامات ACE")
st.markdown("مشروع: **تاج سلطان** | الاستشاري: **محرم باخوم**")

# القائمة الجانبية للبيانات الأساسية
with st.sidebar:
    st.header("📋 بيانات الطلب")
    project_name = st.text_input("اسم المشروع", "تاج سلطان")
    insp_no = st.text_input("رقم طلب الاستلام")
    contractor = st.text_input("مندوب المقاول")
    date_str = st.date_input("التاريخ", datetime.now()).strftime('%Y-%m-%d')

# تقسيم الشاشة إلى تبويبات لتسهيل الاستخدام على الجوال
tab1, tab2, tab3 = st.tabs(["مراجعة إدارية", "فحص الخنزيرة", "التقرير النهائي"])

with tab1:
    st.subheader("تحقق ما قبل الاستلام")
    g_checks = {
        "اعتماد المواد المستخدمة [1]": st.checkbox("اعتماد المواد"),
        "اعتماد رسومات الورشة [2]": st.checkbox("اعتماد رسومات الورشة"),
        "اعتماد كيفية التنفيذ [3]": st.checkbox("اعتماد طريقة التنفيذ (Method Statement)"),
        "مراجعة رسومات التصميم [6]": st.checkbox("مراجعة المخططات"),
        "شروط السلامة المهنية [7]": st.checkbox("شروط السلامة والأمان"),
        "إزالة المخلفات [8]": st.checkbox("إزالة المخلفات")
    }

with tab2:
    st.subheader("بنود الفحص الفني")
    tech_items = [
        "أبعاد الخنزيرة أكبر من قاع الحفر [1]",
        "استقامة الأضلاع بخيط الشد [2]",
        "تقوية الأضلاع بالخوابير والشكالات [3]",
        "أفقية الأضلاع بميزان المياه [4]",
        "أفقية زوايا الالتقاء [5]",
        "مراجعة الزوايا (نظرية فيثاغورث) [6]",
        "عدم وجود حركة في الزوايا [7]",
        "مراجعة المحاور على الخنزيرة [8]"
    ]
    t_results = {}
    for item in tech_items:
        t_results[item] = st.selectbox(item, ["مطابق", "غير مطابق", "غير مطبق"])

with tab3:
    st.subheader("التقييم والاعتماد")
    final_status = st.radio("درجة التقييم:", [
        "1 - لا توجد ملاحظات (يعتمد)",
        "2 - يتم تنفيذ الملاحظات",
        "3 - يعاد التقديم بعد إتمام الملاحظات",
        "4 - غير مقبول"
    ])
    notes = st.text_area("ملاحظات عامة لمهندس الموقع")
    
    # ميزة الكاميرا لالتقاط صورة (اختياري في التقرير حالياً)
    st.camera_input("📸 توثيق الموقع (اختياري)")

    st.markdown("---")
    if st.button("✅ إصدار واعتماد التقرير", use_container_width=True):
        if not insp_no or not contractor:
            st.error("⚠️ يرجى إكمال رقم الطلب واسم المقاول في القائمة الجانبية أولاً.")
        else:
            form_data = {
                "project": project_name,
                "inspection_no": insp_no,
                "contractor": contractor,
                "date": date_str,
                "general_checks": g_checks,
                "tech_checks": t_results,
                "status": final_status,
                "comments": notes
            }
            
            doc_bytes = create_docx(form_data)
            
            st.success("🎉 تم تجهيز التقرير بنجاح!")
            st.download_button(
                label="📥 تحميل ملف Word وإرساله",
                data=doc_bytes,
                file_name=f"Inspection_Khanezira_{insp_no}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
