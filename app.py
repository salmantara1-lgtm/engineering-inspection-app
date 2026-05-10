import streamlit as st
from datetime import datetime
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches
from io import BytesIO

# --- الإعدادات الأساسية للصفحة ---
st.set_page_config(page_title="نظام الاستلامات الهندسية الشامل", page_icon="🏗️", layout="wide")

# ==========================================
# قاعدة بيانات الاستلامات (مستخرجة من ملف الـ PDF)
# يمكنك نسخ نفس التنسيق لإضافة المزيد من البنود لاحقاً
# ==========================================
INSPECTIONS_DATA = {
    "أعمال الحفر": [
        "مراجعة الميزانية الشبكية للأرض الطبيعية قبل الحفر.",
        "مراجعة إحداثيات حدود الحفر مع حدود المبنى طبقا لتقرير الجسات.",
        "مراجعة منسوب التأسيس مع اللوحات وعمق الحفر من تقرير التربة.",
        "مراجعة تطهير قاع وجوانب الحفر واستواء الحفر.",
        "التأكد من أن نوع تربة قاع الحفر مطابقة لتقرير التربة ولا توجد اختلافات."
    ],
    "أعمال الإحلال": [
        "التأكد من منسوب طبقة الإحلال (منسوب التأسيس).",
        "التأكد من سمك طبقة الإحلال طبقاً لتوصيات الاستشاري.",
        "التأكد من طريقة الدمك المستخدمة (ميكانيكي/ستاتيكي/هز) طبقا للمواصفات.",
        "التأكد من الوزن المستخدم في الدمك.",
        "التأكد من نتيجة اختبار الدمك قبل البدء في الطبقة التالية.",
        "مطابقة طبقة الإحلال للعينة المعتمدة."
    ],
    "فحص واستلام الخنزيرة": [
        "أبعاد الخنزيرة أكبر من أبعاد قاع الحفر بمسافة تمنع تأثرها.",
        "استقامة أضلاع الخنزيرة بخيط الشد.",
        "تقوية الأضلاع بالخوابير أو الشكالات في حالة الارتفاع عن الأرض.",
        "أفقية الأضلاع بميزان المياه.",
        "مراجعة أفقية اضلاع الخنزيرة عند أماكن الالتقاء.",
        "مراجعة الزوايا المحصورة بين الأضلاع (نظرية فيثاغورث).",
        "التأكد من عدم حدوث أي حركة في زوايا الالتقاء بتقويتها جيدا.",
        "مراجعة المحاور على الخنزيرة."
    ],
    "القواعد الخرسانية العادية": [
        "التأكد من نظافة منسوب التأسيس مع تمام دمك السطح.",
        "رفع أركان القواعد بجهاز Total Station ومطابقتها مع الرسومات.",
        "مطابقة المحاور الإنشائية وصحة توقيع زواياها.",
        "مراجعة أبعاد القواعد وارتفاعاتها وحدود جوانبها.",
        "التقفيل الجيد لجوانب القواعد مع بعضها وتسديد الفتحات.",
        "مراجعة التقويات والتأكد من إتمامها بطريقة صحيحة.",
        "التأكد من وضع شرب صب القواعد (منسوب الظهر) بميزان القامة."
    ],
    "القواعد الخرسانية المسلحة والسملات": [
        "مراجعة أبعاد القواعد المسلحة والسملات.",
        "التأكد من نظافة القوالب الخشبية من الداخل.",
        "التأكد من دهان القوالب الخشبية بمادة مانعة للالتصاق.",
        "مراجعة أماكن الفتحات ومسارات الصحي والكهرباء.",
        "مراجعة التقويات (الزراجين والشكالات) لمنع الحركة أثناء الصب."
    ],
    "حديد تسليح الأساسات": [
        "التأكد من أقطار وأعداد حديد التسليح ومطابقتها للمخططات.",
        "التأكد من أطوال التراكب (الوصلات) وأماكنها.",
        "مراجعة تربيط الحديد بالسلك المجلفن بشكل جيد.",
        "مراجعة تركيب البسكويت (الغطاء الخرساني) بالسمك المطلوب.",
        "التأكد من نظافة حديد التسليح من الصدأ أو الزيوت."
    ],
    "نجارة الأعمدة الخرسانية": [
        "مراجعة قطاعات الأعمدة وأبعادها.",
        "التأكد من رأسية الأعمدة باستخدام ميزان الخيط (البلبل).",
        "مراجعة أماكن الأعمدة ومحاورها.",
        "التأكد من متانة التقويات (الأحزمة والزراجين).",
        "التأكد من منسوب نهاية صب العمود."
    ],
    "إستلام أعمال المباني": [
        "التأكد من عمل المدماك الأول بكامل الدور واسترباع الغرف.",
        "التأكد من وضع قوالب الطوب للمدماك الأول على فرشة كاملة من المونة.",
        "تحديد أماكن الفتحات (الأبواب والشبابيك) ووزن المباني أسفل الكمرات.",
        "التأكد من ملء العراميس الطولية والعرضية بالمونة.",
        "وزن المباني رأسياً بميزان الخيط وأفقياً بميزان المياه."
    ],
    "أعمال العزل المائي للأساسات": [
        "التأكد من جفاف ونظافة الأسطح المراد عزلها وتنعيمها.",
        "التأكد من عمل وزرة (رقبة زجاجة) من المونة في الزوايا.",
        "مراجعة دهان طبقة الأساس (الأساس البارد/البرايمر) وتغطيتها السطح بالكامل.",
        "التأكد من ركوب لفائف العزل (الرولات) على بعضها بالمسافة المطلوبة.",
        "التأكد من لحام الوصلات جيدا وعدم وجود فراغات."
    ]
}

# --- دالة إنشاء ملف Word شامل الصور ---
def create_docx(data):
    doc = Document()
    
    # تنسيق اتجاه النص من اليمين لليسار (اختياري حسب دعم المكتبة)
    style = doc.styles['Normal']
    font = style.font
    font.name = 'Arial'
    
    heading = doc.add_heading(f"نموذج فحص واستلام - {data['inspection_type']}", level=1)
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

    # جدول الفحص الفني
    doc.add_heading('بنود المراجعة الفنية:', level=2)
    t2 = doc.add_table(rows=1, cols=2)
    t2.style = 'Table Grid'
    t2.rows[0].cells[0].text = 'بند المراجعة'
    t2.rows[0].cells[1].text = 'النتيجة'
    for key, value in data['tech_checks'].items():
        row = t2.add_row().cells
        row[0].text = key
        row[1].text = value

    # التقييم النهائي
    doc.add_heading('النتيجة والملاحظات:', level=2)
    doc.add_paragraph(f"التقييم النهائي: {data['status']}")
    doc.add_paragraph(f"ملاحظات المهندس: {data['comments']}")

    # إضافة الصور إن وجدت
    if data['photos']:
        doc.add_page_break()
        doc.add_heading('توثيق الموقع (الصور والملاحظات):', level=1)
        for idx, photo_data in enumerate(data['photos']):
            try:
                # قراءة الصورة من الذاكرة
                image_stream = BytesIO(photo_data['image'].getvalue())
                # إضافة الصورة للتقرير بعرض مناسب
                doc.add_picture(image_stream, width=Inches(4.5))
                # إضافة وصف الصورة
                p_desc = doc.add_paragraph()
                p_desc.add_run(f"صورة رقم ({idx+1}): ").bold = True
                p_desc.add_run(photo_data['desc'])
                doc.add_paragraph("\n") # فاصل بين الصور
            except Exception as e:
                doc.add_paragraph(f"تعذر إدراج الصورة رقم {idx+1}")

    target = BytesIO()
    doc.save(target)
    return target.getvalue()

# --- واجهة المستخدم (UI) ---
st.title("🏗️ نظام الاستلامات الهندسية (المتكامل)")

# القائمة الجانبية
with st.sidebar:
    st.header("📋 بيانات الطلب")
    project_name = st.text_input("اسم المشروع")
    insp_no = st.text_input("رقم طلب الاستلام")
    contractor = st.text_input("اسم المقاول / المندوب")
    date_str = st.date_input("تاريخ الفحص", datetime.now()).strftime('%Y-%m-%d')

# اختيار نوع الاستلام
st.markdown("### 🔍 اختر مرحلة الاستلام المراد فحصها:")
selected_inspection = st.selectbox(
    "قائمة الاستلامات المتاحة:", 
    list(INSPECTIONS_DATA.keys())
)

st.markdown("---")

tab1, tab2, tab3, tab4 = st.tabs(["📋 المراجعة الإدارية", "🛠️ الفحص الفني", "📸 التوثيق والصور", "✅ الاعتماد والتقرير"])

with tab1:
    st.subheader("تحقق ما قبل الاستلام")
    g_checks = {
        "التأكد من اعتماد المواد المستخدمة في البند": st.checkbox("اعتماد المواد", value=True),
        "التأكد من اعتماد رسومات الورشة (Shop Drawings)": st.checkbox("اعتماد المخططات التنفيذية", value=True),
        "التأكد من اعتماد طريقة التنفيذ (Method Statement)": st.checkbox("اعتماد طريقة التنفيذ"),
        "توافر شروط السلامة المهنية بالموقع": st.checkbox("شروط السلامة والأمان"),
        "التأكد من إزالة المخلفات الناتجة عن التنفيذ": st.checkbox("إزالة المخلفات")
    }

with tab2:
    st.subheader(f"بنود فحص: {selected_inspection}")
    current_tech_items = INSPECTIONS_DATA[selected_inspection]
    
    t_results = {}
    for item in current_tech_items:
        # استخدام st.radio كبديل أفضل للموبايل للمطابقة
        t_results[item] = st.radio(item, ["مطابق", "غير مطابق", "غير مطبق"], horizontal=True, key=item)

with tab3:
    st.subheader("إرفاق الصور والملاحظات (اختياري - حتى 8 صور)")
    st.info("يمكنك التصوير المباشر من كاميرا الجوال أو رفع الصور من الاستديو، مع كتابة تعليق أسفل كل صورة ليظهر في التقرير.")
    
    captured_photos = []
    
    # عرض خانات الصور في عمودين لتوفير مساحة الشاشة
    col1, col2 = st.columns(2)
    
    for i in range(1, 9):
        current_col = col1 if i % 2 != 0 else col2
        with current_col:
            with st.expander(f"📷 إرفاق صورة رقم {i}"):
                img_file = st.file_uploader(f"اختر أو التقط صورة", type=['png', 'jpg', 'jpeg'], key=f"file_{i}")
                img_desc = st.text_input(f"ملاحظة للصورة:", key=f"desc_{i}")
                
                if img_file is not None:
                    st.image(img_file, use_column_width=True)
                    captured_photos.append({"image": img_file, "desc": img_desc if img_desc else "لا يوجد وصف"})

with tab4:
    st.subheader("التقييم والاعتماد")
    final_status = st.radio("القرار النهائي للمهندس:", [
        "1 - لا توجد ملاحظات (يعتمد البند)",
        "2 - يعتمد مع تنفيذ الملاحظات",
        "3 - يعاد التقديم بعد إتمام الملاحظات",
        "4 - مرفوض"
    ])
    notes = st.text_area("ملاحظات عامة / أسباب الرفض (إن وجدت)")
    
    st.markdown("---")
    
    if st.button("✅ إصدار التقرير (Word)", use_container_width=True):
        if not insp_no or not project_name:
            st.error("⚠️ يرجى إكمال 'اسم المشروع' و 'رقم الطلب' في القائمة الجانبية أولاً.")
        else:
            form_data = {
                "inspection_type": selected_inspection,
                "project": project_name,
                "inspection_no": insp_no,
                "contractor": contractor,
                "date": date_str,
                "general_checks": g_checks,
                "tech_checks": t_results,
                "photos": captured_photos,
                "status": final_status,
                "comments": notes
            }
            
            with st.spinner("جاري إنشاء التقرير ودمج الصور..."):
                doc_bytes = create_docx(form_data)
                
            st.success("🎉 تم تجهيز التقرير بنجاح!")
            st.download_button(
                label="📥 اضغط هنا لتحميل تقرير Word الشامل",
                data=doc_bytes,
                file_name=f"Inspection_{selected_inspection}_{insp_no}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True
            )
