import streamlit as st
import pandas as pd
import io
import base64
import requests
from datetime import datetime

# إعدادات الصفحة ودعم اللغة العربية (RTL)
st.set_page_config(page_title="مراجع قواعد بيانات المستشفيات - HIS Audit", layout="wide", page_icon="🏥")

st.markdown("""
<style>
    .main { direction: rtl; text-align: right; }
    div[data-testid="stSidebarNav"] { direction: rtl; text-align: right; }
    .stTabs [data-baseweb="tab-list"] { direction: rtl; }
    .stDataFrame { direction: ltr; }
    h1, h2, h3, h4, p, label { text-align: right; direction: rtl; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
</style>
""", unsafe_allow_html=True)

DB_FILE = "his_knowledge_base.xlsx"

# ----------------- البيانات التأسيسية (البيانات التي استخرجناها معاً) -----------------
def get_seed_data():
    tables_data = [
        {"اسم_الجدول": "CSHCASHUNITTRANS", "الموديول": "الخزينة والبنوك", "النوع": "Master", "الوصف": "إيصالات الدفع النقدية والشيكات وسندات الصرف والقبض للموردين والمرضى"},
        {"اسم_الجدول": "MMSTORETRANSACTIONS", "الموديول": "المخازن", "النوع": "Master", "الوصف": "الرأس الأساسي لحركات المخازن (أذون الإضافة، الفواتير، المرتجعات)"},
        {"اسم_الجدول": "MMSTORETRANSITEMS", "الموديول": "المخازن", "النوع": "Details", "الوصف": "تفاصيل أصناف حركات المخازن، الكميات والأسعار الفعلية لكل صنف"},
        {"اسم_الجدول": "GLTEMPGENERALJOURNAL", "الموديول": "الحسابات العامة", "النوع": "Master", "الوصف": "رأس قيود اليومية العامة، تاريخ القيد الفعلي والسنة المالية"},
        {"اسم_الجدول": "GLTEMPGENJOURNACCOUNTS", "الموديول": "الحسابات العامة", "النوع": "Details", "الوصف": "تفاصيل أطراف القيد المحاسبي (الحسابات المدينة والدائنة وقيمها)"},
        {"اسم_الجدول": "MMPURCHASINGORDER", "الموديول": "المشتريات", "النوع": "Master", "الوصف": "رأس أوامر الشراء الصادرة للموردين"},
        {"اسم_الجدول": "TAX", "الموديول": "الضرائب والاستقطاعات", "النوع": "Master", "الوصف": "جدول تعريف وربط الضرائب ونسب الخصم والاستقطاع"}
    ]

    columns_data = [
        {"اسم_الجدول": "CSHCASHUNITTRANS", "اسم_العمود": "TRANSSERIAL", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "PK", "الوصف": "المسلسل الداخلي لحركة الخزينة في القاعدة"},
        {"اسم_الجدول": "CSHCASHUNITTRANS", "اسم_العمود": "RECEIPTNO", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "Normal", "الوصف": "رقم الإيصال الفعلي المطبوع والظاهر للمستخدم"},
        {"اسم_الجدول": "CSHCASHUNITTRANS", "اسم_العمود": "TRANSDATE", "نوع_البيانات": "DATE", "نوع_المفتاح": "Normal", "الوصف": "تاريخ تحرير إيصال الدفع أو القبض"},
        {"اسم_الجدول": "CSHCASHUNITTRANS", "اسم_العمود": "SUPPLIERCD", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "FK", "الوصف": "كود المورد في الموديول (مثل 504 للفاتح، 583 لكريم فارما)"},
        {"اسم_الجدول": "CSHCASHUNITTRANS", "اسم_العمود": "TOTALAMOUNT", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "Normal", "الوصف": "إجمالي قيمة الإيصال في الموديول"},
        {"اسم_الجدول": "CSHCASHUNITTRANS", "اسم_العمود": "TAX_DEDUCTION", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "Normal", "الوصف": "قيمة الاستقطاع الضريبي (ضرائب المنبع) على الإيصال"},
        {"اسم_الجدول": "CSHCASHUNITTRANS", "اسم_العمود": "DISC_VALUE", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "Normal", "الوصف": "قيمة الخصم المكتسب على الإيصال"},
        {"اسم_الجدول": "CSHCASHUNITTRANS", "اسم_العمود": "ENTRYSERIAL", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "FK", "الوصف": "رقم القيد المحاسبي المرتبط في الحسابات العامة"},
        {"اسم_الجدول": "MMSTORETRANSACTIONS", "اسم_العمود": "GEN_SERIAL", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "PK", "الوصف": "المسلسل العام لحركة المخزن ومفتاح الربط مع الأصناف"},
        {"اسم_الجدول": "MMSTORETRANSACTIONS", "اسم_العمود": "STORETRANSCODE", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "Normal", "الوصف": "كود نوع الحركة المخزنية (2 = إذن إضافة مشتريات)"},
        {"اسم_الجدول": "MMSTORETRANSACTIONS", "اسم_العمود": "ENTRYSERIAL", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "FK", "الوصف": "رقم قيد اليومية المرتبط بالفاتورة في الحسابات العامة"},
        {"اسم_الجدول": "MMSTORETRANSITEMS", "اسم_العمود": "ACTQUANTITY", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "Normal", "الوصف": "الكمية الفعلية المستلمة من الصنف"},
        {"اسم_الجدول": "MMSTORETRANSITEMS", "اسم_العمود": "UNITCOSTPRICE", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "Normal", "الوصف": "سعر تكلفة الوحدة الواحدة"},
        {"اسم_الجدول": "GLTEMPGENERALJOURNAL", "اسم_العمود": "ENTRYSERIAL", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "PK", "الوصف": "المسلسل الرئيسي للقيد في الحسابات العامة"},
        {"اسم_الجدول": "GLTEMPGENERALJOURNAL", "اسم_العمود": "ENTRYACTDATE", "نوع_البيانات": "DATE", "نوع_المفتاح": "Normal", "الوصف": "التاريخ الفعلي لإنشاء القيد وتأثيره في اليومية"},
        {"اسم_الجدول": "GLTEMPGENERALJOURNAL", "اسم_العمود": "YEARCODE", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "Normal", "الوصف": "السنة المالية المقيد عليها القيد (2024، 2025)"},
        {"اسم_الجدول": "GLTEMPGENJOURNACCOUNTS", "اسم_العمود": "ACCOUNTCODE", "نوع_البيانات": "VARCHAR2", "نوع_المفتاح": "FK", "الوصف": "كود الحساب المالي (202006 متبوعاً بكود المورد)"},
        {"اسم_الجدول": "GLTEMPGENJOURNACCOUNTS", "اسم_العمود": "DEBIT", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "Normal", "الوصف": "القيمة المدينة (تخفض رصيد المورد)"},
        {"اسم_الجدول": "GLTEMPGENJOURNACCOUNTS", "اسم_العمود": "CREDIT", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "Normal", "الوصف": "القيمة الدائنة (تزيد رصيد المورد)"}
    ]

    sql_data = [
        {
            "رقم_الاستعلام": 1,
            "عنوان_المشكلة": "كشف فروق إيصالات الدفع والاستقطاعات الساقطة",
            "الجداول_المرتبطة": "CSHCASHUNITTRANS, GLTEMPGENJOURNACCOUNTS",
            "الاصدار": "v1.2",
            "شرح_الحل": "يربط الإيصالات بالقيود عبر ENTRYSERIAL ويقارن إجمالي الإيصال بمدين المورد لكشف الضرائب الساقطة. تم تحديث العمود إلى RECEIPTNO.",
            "كود_SQL": """SELECT 
    C.RECEIPTNO AS "رقم الإيصال",
    C.TRANSSERIAL AS "رقم الحركة",
    TO_CHAR(C.TRANSDATE, 'YYYY-MM-DD') AS "تاريخ الإيصال",
    C.ENTRYSERIAL AS "رقم القيد",
    C.TOTALAMOUNT AS "المبلغ في الموديول",
    A.DEBIT AS "مدين المورد في القيود",
    (A.DEBIT - NVL(C.TOTALAMOUNT, 0)) AS "الفرق (الاستقطاع الساقط)"
FROM CSHCASHUNITTRANS C
JOIN GLTEMPGENJOURNACCOUNTS A ON C.ENTRYSERIAL = A.ENTRYSERIAL
WHERE C.TRANSDATE BETWEEN TO_DATE('01/01/2025', 'DD/MM/YYYY') AND TO_DATE('31/12/2025', 'DD/MM/YYYY')
  AND C.SUPPLIERCD = 504
  AND A.ACCOUNTCODE = '2020060504'
  AND A.DEBIT > 0
  AND A.DEBIT <> NVL(C.TOTALAMOUNT, 0)
ORDER BY C.TRANSDATE;""",
            "تاريخ_التحديث": "2026-10-30"
        },
        {
            "رقم_الاستعلام": 2,
            "عنوان_المشكلة": "كشف إيصالات الدفع غير المرحلة (Unposted Receipts)",
            "الجداول_المرتبطة": "CSHCASHUNITTRANS",
            "الاصدار": "v1.0",
            "شرح_الحل": "يستخرج الإيصالات المحفوظة في الموديول وتخصم من كشف حساب الموديول ولكن لم يتولد لها قيد محاسبي.",
            "كود_SQL": """SELECT 
    RECEIPTNO AS "رقم الإيصال", 
    TRANSSERIAL AS "رقم الحركة",
    TO_CHAR(TRANSDATE, 'YYYY-MM-DD') AS "التاريخ", 
    TOTALAMOUNT AS "المبلغ",
    NVL(TAX_DEDUCTION, 0) AS "الضريبة"
FROM CSHCASHUNITTRANS 
WHERE TRANSDATE BETWEEN TO_DATE('01/01/2025', 'DD/MM/YYYY') AND TO_DATE('31/12/2025', 'DD/MM/YYYY')
  AND SUPPLIERCD = 504
  AND (ENTRYSERIAL IS NULL OR ENTRYSERIAL = 0);""",
            "تاريخ_التحديث": "2026-10-30"
        },
        {
            "رقم_الاستعلام": 3,
            "عنوان_المشكلة": "مطابقة أذون الإضافة المخزنية مع القيود الدائنة (تجاوز كسور التقريب)",
            "الجداول_المرتبطة": "MMSTORETRANSACTIONS, MMSTORETRANSITEMS, GLTEMPGENJOURNACCOUNTS",
            "الاصدار": "v1.1",
            "شرح_الحل": "يجمع (الكمية × السعر) لكل إذن ويقارنه بدائن المورد في القيد، مع استخدام ROUND لاستبعاد فروق الكسور العشرية الوهمية.",
            "كود_SQL": """SELECT 
    M.GEN_SERIAL AS "رقم الإذن/الفاتورة",
    TO_CHAR(M.TRANSDATE, 'YYYY-MM-DD') AS "التاريخ",
    M.ENTRYSERIAL AS "رقم القيد",
    ROUND(SUM(I.ACTQUANTITY * I.UNITCOSTPRICE), 2) AS "إجمالي الموديول",
    ROUND(A.CREDIT, 2) AS "دائن المورد بالقيود",
    ROUND(A.CREDIT - NVL(SUM(I.ACTQUANTITY * I.UNITCOSTPRICE), 0), 2) AS "الفرق الفعلي"
FROM MMSTORETRANSACTIONS M
JOIN MMSTORETRANSITEMS I ON M.GEN_SERIAL = I.GEN_SERIAL
JOIN GLTEMPGENJOURNACCOUNTS A ON M.ENTRYSERIAL = A.ENTRYSERIAL
WHERE M.TRANSDATE BETWEEN TO_DATE('01/01/2025', 'DD/MM/YYYY') AND TO_DATE('31/12/2025', 'DD/MM/YYYY')
  AND M.SUPPLIERCD = 504
  AND A.ACCOUNTCODE = '2020060504'
  AND A.CREDIT > 0
GROUP BY M.GEN_SERIAL, M.TRANSDATE, M.ENTRYSERIAL, A.CREDIT
HAVING ABS(ROUND(A.CREDIT - NVL(SUM(I.ACTQUANTITY * I.UNITCOSTPRICE), 0), 2)) > 0.01;""",
            "تاريخ_التحديث": "2026-10-30"
        },
        {
            "رقم_الاستعلام": 4,
            "عنوان_المشكلة": "القناص: كشف القيود اليتيمة وفروق الفترات المالية (Cross-Year Cutoff)",
            "الجداول_المرتبطة": "GLTEMPGENERALJOURNAL, GLTEMPGENJOURNACCOUNTS, CSHCASHUNITTRANS, MMSTORETRANSACTIONS",
            "الاصدار": "v1.0",
            "شرح_الحل": "يكشف القيود المسجلة في الحسابات العامة وليس لها مستند في إيصالات أو مخازن نفس السنة (مثل تعديل فاتورة تخص 2024 في سنة 2025 المفتوحة).",
            "كود_SQL": """SELECT 
    TO_CHAR(G.ENTRYACTDATE, 'YYYY-MM-DD HH24:MI:SS') AS "تاريخ ووقت القيد",
    G.ENTRYSERIAL AS "رقم القيد",
    A.DEBIT AS "مدين",
    A.CREDIT AS "دائن"
FROM GLTEMPGENJOURNACCOUNTS A
JOIN GLTEMPGENERALJOURNAL G ON A.ENTRYSERIAL = G.ENTRYSERIAL
WHERE G.YEARCODE = 2025 
  AND A.ACCOUNTCODE = '2020060504'
  AND A.ENTRYSERIAL NOT IN (SELECT ENTRYSERIAL FROM CSHCASHUNITTRANS WHERE ENTRYSERIAL IS NOT NULL)
  AND A.ENTRYSERIAL NOT IN (SELECT ENTRYSERIAL FROM MMSTORETRANSACTIONS WHERE ENTRYSERIAL IS NOT NULL)
ORDER BY G.ENTRYACTDATE;""",
            "تاريخ_التحديث": "2026-10-30"
        }
    ]

    logs_data = [
        {"الجهة_أو_المورد": "الفاتح فارما (504) / كريم فارما (583)", "قيمة_الفرق": 2463.72, "الحالة": "تم اكتشاف السبب الجذري", "السبب_الجذري": "تعديل فواتير مخزن تخص ديسمبر 2024 في يوم 8 يناير 2025 (قيود 13389، 13390، 13393) مما رمى القيد في 2025 بينما الحركة في موديول 2024."}
    ]

    return pd.DataFrame(tables_data), pd.DataFrame(columns_data), pd.DataFrame(sql_data), pd.DataFrame(logs_data)

# تحميل البيانات في الذاكرة
if 'initialized' not in st.session_state:
    try:
        xls = pd.ExcelFile(DB_FILE)
        st.session_state.df_tables = pd.read_excel(DB_FILE, sheet_name="Tables")
        st.session_state.df_columns = pd.read_excel(DB_FILE, sheet_name="Columns")
        st.session_state.df_sql = pd.read_excel(DB_FILE, sheet_name="SQL_Repo")
        st.session_state.df_logs = pd.read_excel(DB_FILE, sheet_name="Audit_Logs")
    except Exception:
        st.session_state.df_tables, st.session_state.df_columns, st.session_state.df_sql, st.session_state.df_logs = get_seed_data()
    st.session_state.initialized = True

# دالة تحويل قاعدة البيانات إلى ملف إكسيل في الذاكرة
def export_db_to_bytes():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        st.session_state.df_tables.to_excel(writer, sheet_name="Tables", index=False)
        st.session_state.df_columns.to_excel(writer, sheet_name="Columns", index=False)
        st.session_state.df_sql.to_excel(writer, sheet_name="SQL_Repo", index=False)
        st.session_state.df_logs.to_excel(writer, sheet_name="Audit_Logs", index=False)
    # حفظ محلي أيضاً
    with open(DB_FILE, "wb") as f:
        f.write(output.getvalue())
    return output.getvalue()

# دالة المزامنة السحابية مع GitHub لحفظ البيانات للأبد مجاناً
def sync_to_github(file_bytes):
    try:
        token = st.secrets["GITHUB_TOKEN"]
        repo = st.secrets["GITHUB_REPO"] # مثال: username/his-audit-app
        path = DB_FILE
        url = f"https://api.github.com/repos/{repo}/contents/{path}"
        headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
        
        r = requests.get(url, headers=headers)
        sha = r.json().get("sha") if r.status_code == 200 else None
        
        content_b64 = base64.b64encode(file_bytes).decode("utf-8")
        data = {
            "message": f"Auto-sync DB from Hospital at {datetime.now().strftime('%Y-%m-%d %H:%M')}",
            "content": content_b64
        }
        if sha:
            data["sha"] = sha
            
        put_r = requests.put(url, headers=headers, json=data)
        return put_r.status_code in [200, 201]
    except Exception as e:
        return False

# ----------------- القائمة الجانبية -----------------
with st.sidebar:
    st.title("🏥 نظام هندسة ومراجعة HIS")
    st.caption("إصدار سحابي متنقل للمراجعة المحاسبية وقواعد البيانات")
    menu = st.radio(
        "اختر الوحدة:",
        ["📊 لوحة القياس (Dashboard)",
         "🗂️ 1. قاموس البيانات (Data Dictionary)",
         "💻 2. مستودع الاستعلامات (SQL Repo)",
         "🔄 3. الاستيراد والتصدير والسحابة",
         "⚖️ 4. محرك المطابقة الآلي (Reconciliation)"]
    )
    st.divider()
    if st.button("☁️ حفظ ومزامنة التعديلات سحابياً", use_container_width=True):
        excel_bytes = export_db_to_bytes()
        if sync_to_github(excel_bytes):
            st.success("تم الحفظ في GitHub بنجاح! ستجد تعديلاتك من أي جهاز آخر.")
        else:
            st.warning("تم الحفظ محلياً. لتفعيل الحفظ السحابي الدائم، أضف GITHUB_TOKEN في إعدادات Streamlit، أو حمل ملف الإكسيل من وحدة التصدير.")

# ----------------- لوحة القياس -----------------
if menu == "📊 لوحة القياس (Dashboard)":
    st.header("📊 لوحة القياس ومتابعة الفروقات")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("الجداول الموثقة", len(st.session_state.df_tables))
    c2.metric("الأعمدة الموثقة", len(st.session_state.df_columns))
    c3.metric("استعلامات الفحص الجاهزة", len(st.session_state.df_sql))
    c4.metric("سجل الفروقات المحلولة", len(st.session_state.df_logs))
    
    st.subheader("📌 سجل القضايا والفروقات المكتشفة (Audit Logs)")
    st.dataframe(st.session_state.df_logs, use_container_width=True)
    
    with st.expander("➕ تسجيل قضية فرق جديدة أو حالة مورد"):
        with st.form("log_form"):
            ent = st.text_input("اسم المورد / الحساب وكوده:")
            val = st.number_input("قيمة الفرق المالي:", format="%.2f")
            stat = st.selectbox("الحالة:", ["مفتوحة - جاري الفحص", "تم اكتشاف السبب الجذري", "تم التعديل من الحسابات"])
            root = st.text_area("شرح السبب الجذري أو أرقام القيود:")
            if st.form_submit_button("إضافة للسجل"):
                new_log = pd.DataFrame([{"الجهة_أو_المورد": ent, "قيمة_الفرق": val, "الحالة": stat, "السبب_الجذري": root}])
                st.session_state.df_logs = pd.concat([st.session_state.df_logs, new_log], ignore_index=True)
                export_db_to_bytes()
                st.rerun()

# ----------------- الوحدة الأولى: قاموس البيانات -----------------
elif menu == "🗂️ 1. قاموس البيانات (Data Dictionary)":
    st.header("🗂️ إدارة وتوثيق الجداول والأعمدة")
    t1, t2 = st.tabs(["📁 الجداول (Tables)", "🔑 الأعمدة (Columns)"])
    
    with t1:
        st.dataframe(st.session_state.df_tables, use_container_width=True)
        with st.expander("➕ إضافة جدول جديد للقاموس"):
            with st.form("add_table"):
                t_name = st.text_input("اسم الجدول في أوراكل (مثال: MMINVOICES):").upper().strip()
                t_mod = st.selectbox("الموديول:", ["المخازن", "المشتريات", "الخزينة والبنوك", "الحسابات العامة", "الضرائب والاستقطاعات", "المرضى والفواتير"])
                t_type = st.selectbox("نوع الجدول:", ["Master", "Details", "Lookup"])
                t_desc = st.text_area("وصف وظيفة الجدول:")
                if st.form_submit_button("حفظ الجدول"):
                    new_t = pd.DataFrame([{"اسم_الجدول": t_name, "الموديول": t_mod, "النوع": t_type, "الوصف": t_desc}])
                    st.session_state.df_tables = pd.concat([st.session_state.df_tables, new_t], ignore_index=True)
                    export_db_to_bytes()
                    st.rerun()

    with t2:
        selected_tbl = st.selectbox("فلترة الأعمدة حسب الجدول:", ["الكل"] + list(st.session_state.df_tables["اسم_الجدول"].unique()))
        df_c_view = st.session_state.df_columns if selected_tbl == "الكل" else st.session_state.df_columns[st.session_state.df_columns["اسم_الجدول"] == selected_tbl]
        st.dataframe(df_c_view, use_container_width=True)
        
        with st.expander("➕ توثيق عمود جديد"):
            with st.form("add_col"):
                c_tbl = st.selectbox("الجدول التابع له:", list(st.session_state.df_tables["اسم_الجدول"].unique()))
                c_name = st.text_input("اسم العمود:").upper().strip()
                c_dtype = st.selectbox("نوع البيانات:", ["NUMBER", "VARCHAR2", "DATE", "CHAR"])
                c_key = st.selectbox("نوع المفتاح:", ["Normal", "PK", "FK"])
                c_desc = st.text_input("وصف العمود وما يخزنه:")
                if st.form_submit_button("حفظ العمود"):
                    new_c = pd.DataFrame([{"اسم_الجدول": c_tbl, "اسم_العمود": c_name, "نوع_البيانات": c_dtype, "نوع_المفتاح": c_key, "الوصف": c_desc}])
                    st.session_state.df_columns = pd.concat([st.session_state.df_columns, new_c], ignore_index=True)
                    export_db_to_bytes()
                    st.rerun()

# ----------------- الوحدة الثانية: مستودع الاستعلامات -----------------
elif menu == "💻 2. مستودع الاستعلامات (SQL Repo)":
    st.header("💻 مستودع جمل SQL والحلول الهندسية")
    
    for idx, row in st.session_state.df_sql.iterrows():
        with st.expander(f"📌 [{row['الاصدار']}] {row['عنوان_المشكلة']} ({row['الجداول_المرتبطة']})"):
            st.markdown(f"**شرح المشكلة وطريقة العمل:** {row['شرح_الحل']}")
            st.caption(f"تاريخ آخر تحديث: {row['تاريخ_التحديث']}")
            st.code(row['كود_SQL'], language="sql")
            
    st.divider()
    st.subheader("➕ إضافة استعلام جديد أو تحديث إصدار")
    with st.form("sql_form"):
        s_title = st.text_input("عنوان المشكلة / وظيفة الاستعلام:")
        s_tables = st.text_input("الجداول المرتبطة (افصل بينها بفاصلة):")
        s_ver = st.text_input("رقم الإصدار:", value="v1.0")
        s_desc = st.text_area("شرح المشكلة وكيفية عمل الكود:")
        s_code = st.text_area("كود SQL:", height=180)
        if st.form_submit_button("حفظ في المستودع"):
            new_id = len(st.session_state.df_sql) + 1
            new_s = pd.DataFrame([{
                "رقم_الاستعلام": new_id, "عنوان_المشكلة": s_title, "الجداول_المرتبطة": s_tables,
                "الاصدار": s_ver, "شرح_الحل": s_desc, "كود_SQL": s_code,
                "تاريخ_التحديث": datetime.now().strftime("%Y-%m-%d")
            }])
            st.session_state.df_sql = pd.concat([st.session_state.df_sql, new_s], ignore_index=True)
            export_db_to_bytes()
            st.rerun()

# ----------------- الوحدة الثالثة: الاستيراد والتصدير -----------------
elif menu == "🔄 3. الاستيراد والتصدير والسحابة":
    st.header("🔄 محرك الاستيراد والتصدير الموحد (Excel Sync)")
    st.write("يمكنك تحميل قاعدة المعرفة بالكامل في شيت إكسيل واحد، التعديل عليها أو إرسالها للذكاء الاصطناعي لتعبئتها، ثم إعادة رفعها هنا لتحديث النظام فوراً.")
    
    excel_data = export_db_to_bytes()
    st.download_button(
        label="📥 تحميل قاعدة بيانات التطبيق الكاملة (Excel Template)",
        data=excel_data,
        file_name=f"HIS_Audit_KnowledgeBase_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    
    st.divider()
    uploaded_db = st.file_uploader("📤 ارفع ملف الإكسيل المحدث لاستيراد الجداول والأعمدة والاستعلامات دفعة واحدة:", type=['xlsx'])
    if uploaded_db:
        if st.button("تأكيد استيراد وتحديث القاعدة"):
            st.session_state.df_tables = pd.read_excel(uploaded_db, sheet_name="Tables")
            st.session_state.df_columns = pd.read_excel(uploaded_db, sheet_name="Columns")
            st.session_state.df_sql = pd.read_excel(uploaded_db, sheet_name="SQL_Repo")
            st.session_state.df_logs = pd.read_excel(uploaded_db, sheet_name="Audit_Logs")
            export_db_to_bytes()
            st.success("تم استيراد وتحديث كافة الجداول والاستعلامات بنجاح!")
            st.rerun()

# ----------------- الوحدة الرابعة: محرك المطابقة الآلي -----------------
elif menu == "⚖️ 4. محرك المطابقة الآلي (Reconciliation)":
    st.header("⚖️ محرك المطابقة والتسويات الآلي عبر الإكسيل")
    
    col1, col2 = st.columns(2)
    with col1:
        mod_file = st.file_uploader("1️⃣ ارفع شيت الموديول (إيصالات / مخازن / فواتير)", type=['xlsx', 'csv'], key="mod")
    with col2:
        gl_file = st.file_uploader("2️⃣ ارفع شيت القيود اليومية (GL)", type=['xlsx', 'csv'], key="gl")

    if mod_file and gl_file:
        df_mod = pd.read_excel(mod_file) if mod_file.name.endswith('xlsx') else pd.read_csv(mod_file)
        df_gl = pd.read_excel(gl_file) if gl_file.name.endswith('xlsx') else pd.read_csv(gl_file)
        
        st.subheader("إعداد مفاتيح الربط والتجميع")
        c1, c2, c3 = st.columns(3)
        with c1:
            mod_key = st.selectbox("عمود رقم القيد في الموديول:", df_mod.columns)
            mod_val = st.selectbox("عمود المبلغ في الموديول:", df_mod.columns)
        with c2:
            gl_key = st.selectbox("عمود رقم القيد في الحسابات (GL):", df_gl.columns)
            gl_val = st.selectbox("عمود المبلغ في الحسابات (DEBIT أو CREDIT):", df_gl.columns)
        with c3:
            tolerance = st.number_input("هامش كسور التقريب المسموح:", value=0.01, step=0.01)

        if st.button("🚀 تشغيل الفحص واستخراج الفروقات", use_container_width=True):
            df_mod[mod_key] = pd.to_numeric(df_mod[mod_key], errors='coerce').fillna(0).astype(int)
            df_gl[gl_key] = pd.to_numeric(df_gl[gl_key], errors='coerce').fillna(0).astype(int)

            unposted_mod = df_mod[df_mod[mod_key] == 0]
            posted_mod = df_mod[df_mod[mod_key] != 0]

            mod_agg = posted_mod.groupby(mod_key)[mod_val].sum().reset_index().rename(columns={mod_val: 'إجمالي_الموديول'})
            gl_agg = df_gl.groupby(gl_key)[gl_val].sum().reset_index().rename(columns={gl_val: 'إجمالي_القيود'})

            merged = pd.merge(mod_agg, gl_agg, left_on=mod_key, right_on=gl_key, how='outer').fillna(0)
            merged['رقم_القيد'] = merged[mod_key].where(merged[mod_key] != 0, merged[gl_key]).astype(int)
            merged['الفرق'] = (merged['إجمالي_القيود'] - merged['إجمالي_الموديول']).round(2)

            mismatches = merged[(merged['الفرق'].abs() > tolerance) & (merged['إجمالي_الموديول'] != 0) & (merged['إجمالي_القيود'] != 0)]
            orphan_gl = merged[(merged['إجمالي_الموديول'] == 0) & (merged['إجمالي_القيود'] != 0)]

            t_mis, t_orph, t_unp = st.tabs([
                f"⚠️ فروق مبالغ وضرائب ساقطة ({len(mismatches)})",
                f"👻 قيود في الحسابات بدون موديول ({len(orphan_gl)})",
                f"🛑 حركات موديول غير مرحلة ({len(unposted_mod)})"
            ])
            with t_mis:
                st.dataframe(mismatches[['رقم_القيد', 'إجمالي_الموديول', 'إجمالي_القيود', 'الفرق']], use_container_width=True)
            with t_orph:
                st.dataframe(orphan_gl[['رقم_القيد', 'إجمالي_القيود']], use_container_width=True)
            with t_unp:
                st.dataframe(unposted_mod, use_container_width=True)