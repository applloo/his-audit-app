import streamlit as st
import pandas as pd
import io
import base64
import requests
from datetime import datetime

st.set_page_config(page_title="مراجع قواعد بيانات المستشفيات - HIS Audit", layout="wide", page_icon="🏥")

DB_FILE = "his_knowledge_base.xlsx"

# ----------------- القائمة الجانبية والدارك مود -----------------
with st.sidebar:
    st.title("🏥 نظام هندسة ومراجعة HIS")
    dark_mode = st.toggle("🌙 الوضع الليلي (Dark Mode)", value=True)
    st.caption("إصدار سحابي متنقل للمراجعة المحاسبية وقواعد البيانات")

# تطبيق الـ CSS للاتجاه العربي والدارك مود
bg_color = "#0e1117" if dark_mode else "#ffffff"
text_color = "#fafafa" if dark_mode else "#1f2937"
card_bg = "#1e2530" if dark_mode else "#f3f4f6"

st.markdown(f"""
<style>
    .stApp {{ background-color: {bg_color}; color: {text_color}; direction: rtl; text-align: right; }}
    .main {{ direction: rtl; text-align: right; }}
    div[data-testid="stSidebar"] {{ background-color: {card_bg}; direction: rtl; text-align: right; }}
    .stTabs [data-baseweb="tab-list"] {{ direction: rtl; }}
    .stDataFrame {{ direction: ltr; }}
    h1, h2, h3, h4, p, label, span, div {{ font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }}
    div[data-testid="stMetricValue"] {{ color: #38bdf8 !important; }}
</style>
""", unsafe_allow_html=True)

# ----------------- البيانات التأسيسية المؤكدة والمعتمدة -----------------
def get_seed_data():
    tables_data = [
        {"اسم_الجدول": "CSHCASHUNITTRANS", "الموديول": "الخزينة والبنوك", "النوع": "Master", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "إيصالات الدفع النقدية والشيكات وسندات الصرف والقبض للموردين والمرضى"},
        {"اسم_الجدول": "MMSTORETRANSACTIONS", "الموديول": "المخازن", "النوع": "Master", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "الرأس الأساسي لحركات المخازن (أذون الإضافة، الفواتير، المرتجعات)"},
        {"اسم_الجدول": "MMSTORETRANSITEMS", "الموديول": "المخازن", "النوع": "Details", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "تفاصيل أصناف حركات المخازن، الكميات والأسعار الفعلية لكل صنف"},
        {"اسم_الجدول": "GLTEMPGENERALJOURNAL", "الموديول": "الحسابات العامة", "النوع": "Master", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "رأس قيود اليومية العامة، تاريخ القيد الفعلي والسنة المالية ونوع اليومية"},
        {"اسم_الجدول": "GLTEMPGENJOURNACCOUNTS", "الموديول": "الحسابات العامة", "النوع": "Details", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "تفاصيل أطراف القيد المحاسبي (الحسابات المدينة والدائنة وقيمها)"},
        {"اسم_الجدول": "GLTEMPGENERALJOURNAL_TYPES", "الموديول": "الحسابات العامة", "النوع": "Lookup", "حالة_التأكد": "مؤكد من شجرة القاعدة", "الوصف": "أنواع القيود اليومية (مثل كود 112 = تعديل فاتورة مخزن)"},
        {"اسم_الجدول": "GLTEMPGENERALJOURNAL_DETAILS", "الموديول": "الحسابات العامة", "النوع": "Details", "حالة_التأكد": "مؤكد من شجرة القاعدة", "الوصف": "تفاصيل إضافية لقيود اليومية العامة"},
        {"اسم_الجدول": "MMPURCHASINGORDER", "الموديول": "المشتريات", "النوع": "Master", "حالة_التأكد": "مرجعي - يحتاج فحص الأعمدة", "الوصف": "رأس أوامر الشراء الصادرة للموردين"},
        {"اسم_الجدول": "TAX", "الموديول": "الضرائب والاستقطاعات", "النوع": "Master", "حالة_التأكد": "مرجعي - يحتاج فحص الأعمدة", "الوصف": "جدول تعريف وربط الضرائب ونسب الخصم والاستقطاع"}
    ]

    columns_data = [
        {"اسم_الجدول": "CSHCASHUNITTRANS", "اسم_العمود": "RECEIPTNO", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "Normal", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "رقم الإيصال الفعلي المطبوع والظاهر للمستخدم"},
        {"اسم_الجدول": "CSHCASHUNITTRANS", "اسم_العمود": "TRANSSERIAL", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "PK", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "المسلسل الداخلي لحركة الخزينة في القاعدة"},
        {"اسم_الجدول": "CSHCASHUNITTRANS", "اسم_العمود": "TRANSDATE", "نوع_البيانات": "DATE", "نوع_المفتاح": "Normal", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "تاريخ تحرير إيصال الدفع أو القبض"},
        {"اسم_الجدول": "CSHCASHUNITTRANS", "اسم_العمود": "SUPPLIERCD", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "FK", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "كود المورد في الموديول (مثل 504 للفاتح، 583 لكريم فارما)"},
        {"اسم_الجدول": "CSHCASHUNITTRANS", "اسم_العمود": "TOTALAMOUNT", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "Normal", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "إجمالي قيمة الإيصال في الموديول"},
        {"اسم_الجدول": "CSHCASHUNITTRANS", "اسم_العمود": "ENTRYSERIAL", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "FK", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "رقم القيد المحاسبي المرتبط في الحسابات العامة (0 أو فارغ = غير مرحل)"},
        {"اسم_الجدول": "CSHCASHUNITTRANS", "اسم_العمود": "TAX_DEDUCTION", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "Normal", "حالة_التأكد": "استرشادي", "الوصف": "قيمة الاستقطاع الضريبي (ضرائب المنبع) على الإيصال"},
        {"اسم_الجدول": "CSHCASHUNITTRANS", "اسم_العمود": "DISC_VALUE", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "Normal", "حالة_التأكد": "استرشادي", "الوصف": "قيمة الخصم المكتسب على الإيصال"},
        {"اسم_الجدول": "MMSTORETRANSACTIONS", "اسم_العمود": "GEN_SERIAL", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "PK", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "المسلسل العام لحركة المخزن ومفتاح الربط مع جدول الأصناف"},
        {"اسم_الجدول": "MMSTORETRANSACTIONS", "اسم_العمود": "TRANSDATE", "نوع_البيانات": "DATE", "نوع_المفتاح": "Normal", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "تاريخ حركة المخزن أو الفاتورة"},
        {"اسم_الجدول": "MMSTORETRANSACTIONS", "اسم_العمود": "SUPPLIERCD", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "FK", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "كود المورد صاحب الفاتورة المخزنية"},
        {"اسم_الجدول": "MMSTORETRANSACTIONS", "اسم_العمود": "ENTRYSERIAL", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "FK", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "رقم قيد اليومية المرتبط بالفاتورة في الحسابات العامة"},
        {"اسم_الجدول": "MMSTORETRANSITEMS", "اسم_العمود": "GEN_SERIAL", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "FK", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "مفتاح الربط مع رأس إذن المخزن MMSTORETRANSACTIONS"},
        {"اسم_الجدول": "MMSTORETRANSITEMS", "اسم_العمود": "ACTQUANTITY", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "Normal", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "الكمية الفعلية المستلمة من الصنف"},
        {"اسم_الجدول": "MMSTORETRANSITEMS", "اسم_العمود": "UNITCOSTPRICE", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "Normal", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "سعر تكلفة الوحدة الواحدة"},
        {"اسم_الجدول": "GLTEMPGENERALJOURNAL", "اسم_العمود": "ENTRYSERIAL", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "PK", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "المسلسل الرئيسي للقيد في الحسابات العامة (مثل 13389)"},
        {"اسم_الجدول": "GLTEMPGENERALJOURNAL", "اسم_العمود": "ENTRYACTDATE", "نوع_البيانات": "DATE", "نوع_المفتاح": "Normal", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "التاريخ والوقت الفعلي لإنشاء القيد وتأثيره في اليومية"},
        {"اسم_الجدول": "GLTEMPGENERALJOURNAL", "اسم_العمود": "YEARCODE", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "Normal", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "السنة المالية المقيد عليها القيد (2024، 2025)"},
        {"اسم_الجدول": "GLTEMPGENJOURNACCOUNTS", "اسم_العمود": "ENTRYSERIAL", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "FK", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "مفتاح الربط مع رأس القيد"},
        {"اسم_الجدول": "GLTEMPGENJOURNACCOUNTS", "اسم_العمود": "ACCOUNTCODE", "نوع_البيانات": "VARCHAR2", "نوع_المفتاح": "FK", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "كود الحساب المالي (202006 متبوعاً بكود المورد مثل 2020060504)"},
        {"اسم_الجدول": "GLTEMPGENJOURNACCOUNTS", "اسم_العمود": "DEBIT", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "Normal", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "القيمة المدينة (تخفض رصيد المورد)"},
        {"اسم_الجدول": "GLTEMPGENJOURNACCOUNTS", "اسم_العمود": "CREDIT", "نوع_البيانات": "NUMBER", "نوع_المفتاح": "Normal", "حالة_التأكد": "مؤكد ومجرب 100%", "الوصف": "القيمة الدائنة (تزيد رصيد المورد)"}
    ]

    sql_data = [
        {
            "رقم_الاستعلام": 1,
            "عنوان_المشكلة": "كشف فروق إيصالات الدفع والاستقطاعات الساقطة",
            "الجداول_المرتبطة": "CSHCASHUNITTRANS, GLTEMPGENJOURNACCOUNTS",
            "الاصدار": "v1.2 (مؤكد)",
            "شرح_الحل": "يربط الإيصالات بالقيود عبر ENTRYSERIAL ويقارن إجمالي الإيصال بمدين المورد لكشف الضرائب الساقطة باستخدام عمود RECEIPTNO الفعلي.",
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
WHERE C.TRANSDATE >= TO_DATE('01/01/2025', 'DD/MM/YYYY') 
  AND C.TRANSDATE <= TO_DATE('31/12/2025', 'DD/MM/YYYY')
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
            "الاصدار": "v1.1 (مؤكد)",
            "شرح_الحل": "يستخرج الإيصالات المحفوظة في الموديول وتخصم من كشف حساب الموديول ولكن لم يتولد لها قيد محاسبي.",
            "كود_SQL": """SELECT 
    RECEIPTNO AS "رقم الإيصال", 
    TRANSSERIAL AS "رقم الحركة",
    TO_CHAR(TRANSDATE, 'YYYY-MM-DD') AS "التاريخ", 
    TOTALAMOUNT AS "المبلغ"
FROM CSHCASHUNITTRANS 
WHERE TRANSDATE BETWEEN TO_DATE('01/01/2025', 'DD/MM/YYYY') AND TO_DATE('31/12/2025', 'DD/MM/YYYY')
  AND SUPPLIERCD = 504
  AND (ENTRYSERIAL IS NULL OR ENTRYSERIAL = 0);""",
            "تاريخ_التحديث": "2026-10-30"
        },
        {
            "رقم_الاستعلام": 3,
            "عنوان_المشكلة": "مطابقة أذون الإضافة المخزنية مع القيود الدائنة (مع فلترة كسور التقريب)",
            "الجداول_المرتبطة": "MMSTORETRANSACTIONS, MMSTORETRANSITEMS, GLTEMPGENJOURNACCOUNTS",
            "الاصدار": "v1.2 (مؤكد)",
            "شرح_الحل": "يجمع (الكمية × السعر) لكل إذن ويقارنه بدائن المورد في القيد، مع استخدام ROUND لاستبعاد فروق الكسور العشرية الوهمية مثل 0.0000000000005.",
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
            "الاصدار": "v1.2 (مؤكد ومجرب)",
            "شرح_الحل": "النسخة المعتمدة بعد استبعاد عمود REMARKS غير الصالح. يكشف القيود المسجلة في الحسابات العامة وليس لها مستند في إيصالات أو مخازن نفس السنة.",
            "كود_SQL": """SELECT 
    G.ENTRYACTDATE AS "التاريخ",
    G.ENTRYSERIAL AS "رقم القيد",
    A.DEBIT AS "مدين",
    A.CREDIT AS "دائن"
FROM GLTEMPGENJOURNACCOUNTS A
JOIN GLTEMPGENERALJOURNAL G ON A.ENTRYSERIAL = G.ENTRYSERIAL
WHERE G.YEARCODE = 2025 
  AND A.ACCOUNTCODE = '2020060504'
  AND A.ENTRYSERIAL NOT IN (SELECT ENTRYSERIAL FROM CSHCASHUNITTRANS WHERE ENTRYSERIAL IS NOT NULL)
  AND A.ENTRYSERIAL NOT IN (SELECT ENTRYSERIAL FROM MMSTORETRANSACTIONS WHERE ENTRYSERIAL IS NOT NULL);""",
            "تاريخ_التحديث": "2026-10-30"
        }
    ]

    logs_data = [
        {"الجهة_أو_المورد": "الفاتح فارما (504) / كريم فارما (583)", "قيمة_الفرق": 2463.72, "من_الأعلى": "الموديول أعلى من الحسابات العامة", "الحالة": "تم اكتشاف السبب الجذري", "السبب_الجذري": "تعديل فواتير مخزن (نوع القيد 112) تخص ديسمبر 2024 في يوم 8 يناير 2025 بواسطة GL.HHUSSEIN (قيود 13389، 13390، 13393) مما رمى القيد في 2025 بينما الحركة في موديول 2024."}
    ]

    return pd.DataFrame(tables_data), pd.DataFrame(columns_data), pd.DataFrame(sql_data), pd.DataFrame(logs_data)

# تحميل البيانات
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

def export_db_to_bytes():
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        st.session_state.df_tables.to_excel(writer, sheet_name="Tables", index=False)
        st.session_state.df_columns.to_excel(writer, sheet_name="Columns", index=False)
        st.session_state.df_sql.to_excel(writer, sheet_name="SQL_Repo", index=False)
        st.session_state.df_logs.to_excel(writer, sheet_name="Audit_Logs", index=False)
    with open(DB_FILE, "wb") as f:
        f.write(output.getvalue())
    return output.getvalue()

def sync_to_github(file_bytes):
    try:
        token = st.secrets["GITHUB_TOKEN"]
        repo = st.secrets["GITHUB_REPO"]
        url = f"https://api.github.com/repos/{repo}/contents/{DB_FILE}"
        headers = {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"}
        r = requests.get(url, headers=headers)
        sha = r.json().get("sha") if r.status_code == 200 else None
        content_b64 = base64.b64encode(file_bytes).decode("utf-8")
        data = {"message": f"Auto-sync DB at {datetime.now().strftime('%Y-%m-%d %H:%M')}", "content": content_b64}
        if sha:
            data["sha"] = sha
        put_r = requests.put(url, headers=headers, json=data)
        return put_r.status_code in [200, 201]
    except Exception:
        return False

# دالة تنظيف شيتات الإكسيل (إزالة الرابش)
def clean_dataframe(df, skip_rows=0, drop_empty=True, ffill_cols=None, calc_mode="بدون", col_a=None, col_b=None):
    df_c = df.copy()
    if skip_rows > 0 and skip_rows < len(df_c):
        df_c.columns = df_c.iloc[skip_rows - 1].astype(str)
        df_c = df_c.iloc[skip_rows:].reset_index(drop=True)
    if drop_empty:
        df_c.dropna(how='all', axis=0, inplace=True)
        df_c.dropna(how='all', axis=1, inplace=True)
        df_c.reset_index(drop=True, inplace=True)
    if ffill_cols:
        for col in ffill_cols:
            if col in df_c.columns:
                df_c[col] = df_c[col].ffill()
    if calc_mode == "ضرب عمودين (الكمية × السعر)" and col_a and col_b:
        a_num = pd.to_numeric(df_c[col_a], errors='coerce').fillna(0)
        b_num = pd.to_numeric(df_c[col_b], errors='coerce').fillna(0)
        df_c["الإجمالي_المحسوب"] = (a_num * b_num).round(2)
    elif calc_mode == "طرح عمودين (مدين - دائن)" and col_a and col_b:
        a_num = pd.to_numeric(df_c[col_a], errors='coerce').fillna(0)
        b_num = pd.to_numeric(df_c[col_b], errors='coerce').fillna(0)
        df_c["الصافي_المحسوب"] = (a_num - b_num).round(2)
    elif calc_mode == "دمج نصي لعمودين" and col_a and col_b:
        df_c["العمود_المدمج"] = df_c[col_a].astype(str).fillna('') + " - " + df_c[col_b].astype(str).fillna('')
    return df_c

# ----------------- عناصر القائمة الجانبية -----------------
with st.sidebar:
    menu = st.radio(
        "اختر الوحدة:",
        ["📊 لوحة القياس (Dashboard)",
         "🗂️ 1. قاموس البيانات (Data Dictionary)",
         "💻 2. مستودع الاستعلامات (SQL Repo)",
         "🧹 3. مغسلة ومنظف الإكسيل (إزالة الرابش)",
         "⚖️ 4. محرك المطابقة الآلي (Reconciliation)",
         "🔄 5. الاستيراد والتصدير والسحابة"]
    )
    st.divider()
    if st.button("☁️ حفظ ومزامنة التعديلات سحابياً", use_container_width=True):
        excel_bytes = export_db_to_bytes()
        if sync_to_github(excel_bytes):
            st.success("تم الحفظ في GitHub بنجاح!")
        else:
            st.warning("تم الحفظ محلياً.")
    if st.button("🔄 تحديث القاموس بالبيانات المعتمدة", use_container_width=True):
        st.session_state.df_tables, st.session_state.df_columns, st.session_state.df_sql, st.session_state.df_logs = get_seed_data()
        export_db_to_bytes()
        st.success("تم تحديث الجداول والأعمدة المعتمدة!")
        st.rerun()

# ----------------- لوحة القياس -----------------
if menu == "📊 لوحة القياس (Dashboard)":
    st.header("📊 لوحة القياس ومتابعة الفروقات")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("الجداول الموثقة", len(st.session_state.df_tables))
    c2.metric("الأعمدة الموثقة", len(st.session_state.df_columns))
    c3.metric("استعلامات الفحص المؤكدة", len(st.session_state.df_sql))
    c4.metric("سجل الفروقات المحلولة", len(st.session_state.df_logs))
    
    st.subheader("📌 سجل القضايا والفروقات المكتشفة (Audit Logs)")
    st.dataframe(st.session_state.df_logs, use_container_width=True)
    
    with st.expander("➕ تسجيل قضية فرق جديدة أو حالة مورد"):
        with st.form("log_form"):
            ent = st.text_input("اسم المورد / الحساب وكوده:")
            val = st.number_input("قيمة الفرق المالي:", format="%.2f")
            higher = st.selectbox("من الأعلى بالفرق؟", ["الموديول أعلى من الحسابات العامة", "الحسابات العامة أعلى من الموديول", "غير محدد"])
            stat = st.selectbox("الحالة:", ["مفتوحة - جاري الفحص", "تم اكتشاف السبب الجذري", "تم التعديل من الحسابات"])
            root = st.text_area("شرح السبب الجذري أو أرقام القيود:")
            if st.form_submit_button("إضافة للسجل"):
                new_log = pd.DataFrame([{"الجهة_أو_المورد": ent, "قيمة_الفرق": val, "من_الأعلى": higher, "الحالة": stat, "السبب_الجذري": root}])
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
                t_name = st.text_input("اسم الجدول في أوراكل:").upper().strip()
                t_mod = st.selectbox("الموديول:", ["المخازن", "المشتريات", "الخزينة والبنوك", "الحسابات العامة", "الضرائب والاستقطاعات", "المرضى والفواتير"])
                t_type = st.selectbox("نوع الجدول:", ["Master", "Details", "Lookup"])
                t_ver = st.selectbox("حالة التأكد:", ["مؤكد ومجرب 100%", "مؤكد من شجرة القاعدة", "مرجعي - يحتاج فحص الأعمدة"])
                t_desc = st.text_area("وصف وظيفة الجدول:")
                if st.form_submit_button("حفظ الجدول"):
                    new_t = pd.DataFrame([{"اسم_الجدول": t_name, "الموديول": t_mod, "النوع": t_type, "حالة_التأكد": t_ver, "الوصف": t_desc}])
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
                c_ver = st.selectbox("حالة التأكد:", ["مؤكد ومجرب 100%", "استرشادي"])
                c_desc = st.text_input("وصف العمود وما يخزنه:")
                if st.form_submit_button("حفظ العمود"):
                    new_c = pd.DataFrame([{"اسم_الجدول": c_tbl, "اسم_العمود": c_name, "نوع_البيانات": c_dtype, "نوع_المفتاح": c_key, "حالة_التأكد": c_ver, "الوصف": c_desc}])
                    st.session_state.df_columns = pd.concat([st.session_state.df_columns, new_c], ignore_index=True)
                    export_db_to_bytes()
                    st.rerun()

# ----------------- الوحدة الثانية: مستودع الاستعلامات -----------------
elif menu == "💻 2. مستودع الاستعلامات (SQL Repo)":
    st.header("💻 مستودع جمل SQL والحلول الهندسية المعتمدة")
    for idx, row in st.session_state.df_sql.iterrows():
        with st.expander(f"📌 [{row['الاصدار']}] {row['عنوان_المشكلة']} ({row['الجداول_المرتبطة']})"):
            st.markdown(f"**شرح المشكلة وطريقة العمل:** {row['شرح_الحل']}")
            st.caption(f"تاريخ آخر تحديث: {row['تاريخ_التحديث']}")
            st.code(row['كود_SQL'], language="sql")
            
    st.divider()
    st.subheader("➕ إضافة استعلام جديد أو تحديث إصدار")
    with st.form("sql_form"):
        s_title = st.text_input("عنوان المشكلة / وظيفة الاستعلام:")
        s_tables = st.text_input("الجداول المرتبطة:")
        s_ver = st.text_input("رقم الإصدار:", value="v1.0 (مؤكد)")
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

# ----------------- الوحدة الثالثة: مغسلة ومنظف الإكسيل -----------------
elif menu == "🧹 3. مغسلة ومنظف الإكسيل (إزالة الرابش)":
    st.header("🧹 معمل تنظيف وتجهيز شيتات الإكسيل المستخرجة من التقارير")
    st.info("استخدم هذه الأداة إذا كان ملف الإكسيل خارجاً من تقرير HISPro وبه صفوف لوجو المستشفى، خلايا مدمجة، أعمدة فارغة، أو يحتاج ضرب الكمية × السعر.")
    
    raw_file = st.file_uploader("ارفع شيت الإكسيل الخام لتنظيفه:", type=['xlsx', 'csv'])
    if raw_file:
        df_raw = pd.read_excel(raw_file) if raw_file.name.endswith('xlsx') else pd.read_csv(raw_file)
        st.subheader("1️⃣ شكل الملف الخام قبل التنظيف:")
        st.dataframe(df_raw.head(10), use_container_width=True)
        
        c1, c2, c3 = st.columns(3)
        with c1:
            skip_r = st.number_input("عدد الصفوف العلوية المراد حذفها (رابش اللوجو والعناوين):", min_value=0, max_value=50, value=0)
            drop_emp = st.checkbox("حذف الصفوف والأعمدة الفارغة تماماً", value=True)
        with c2:
            temp_df = clean_dataframe(df_raw, skip_rows=skip_r, drop_empty=drop_emp)
            ffill_c = st.multiselect("أعمدة مدمجة تحتاج تعبئة تلقائية للأسفل (Forward Fill):", temp_df.columns)
        with c3:
            calc_m = st.selectbox("عمليات حسابية أو دمج أعمدة:", ["بدون", "ضرب عمودين (الكمية × السعر)", "طرح عمودين (مدين - دائن)", "دمج نصي لعمودين"])
            col_1, col_2 = None, None
            if calc_m != "بدون":
                col_1 = st.selectbox("العمود الأول:", temp_df.columns)
                col_2 = st.selectbox("العمود الثاني:", temp_df.columns)

        df_cleaned = clean_dataframe(df_raw, skip_rows=skip_r, drop_empty=drop_emp, ffill_cols=ffill_c, calc_mode=calc_m, col_a=col_1, col_b=col_2)
        
        filter_col = st.selectbox("فلترة وحذف صفوف الإجماليات والنصوص المتكررة (اختر عمود رقمي كشرط):", ["بدون فلترة"] + list(df_cleaned.columns))
        if filter_col != "بدون فلترة":
            df_cleaned = df_cleaned[pd.to_numeric(df_cleaned[filter_col], errors='coerce').notna()].reset_index(drop=True)

        st.subheader(f"2️⃣ النتيجة بعد التنظيف ({len(df_cleaned)} صف نظيف):")
        st.dataframe(df_cleaned, use_container_width=True)
        
        out_clean = io.BytesIO()
        with pd.ExcelWriter(out_clean, engine='openpyxl') as writer:
            df_cleaned.to_excel(writer, index=False, sheet_name="Cleaned_Data")
        
        st.download_button(
            label="📥 تحميل شيت الإكسيل النظيف والجاهز للمطابقة",
            data=out_clean.getvalue(),
            file_name=f"Cleaned_{raw_file.name.split('.')[0]}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

# ----------------- الوحدة الرابعة: محرك المطابقة الآلي -----------------
elif menu == "⚖️ 4. محرك المطابقة الآلي (Reconciliation)":
    st.header("⚖️ محرك المطابقة والتسويات الآلي عبر الإكسيل")
    
    col1, col2 = st.columns(2)
    with col1:
        mod_file = st.file_uploader("1️⃣ ارفع شيت الموديول (إيصالات / مخازن / كشف حساب)", type=['xlsx', 'csv'], key="mod")
    with col2:
        gl_file = st.file_uploader("2️⃣ ارفع شيت القيود اليومية (GL)", type=['xlsx', 'csv'], key="gl")

    if mod_file and gl_file:
        df_mod_raw = pd.read_excel(mod_file) if mod_file.name.endswith('xlsx') else pd.read_csv(mod_file)
        df_gl_raw = pd.read_excel(gl_file) if gl_file.name.endswith('xlsx') else pd.read_csv(gl_file)
        
        # خيارات التنظيف السريع داخل المطابقة
        with st.expander("🧹 خيارات تنظيف سريعة للشيتات المرفوعة (لو فيها رابش أو عناوين علوية أو تحتاج ضرب كمية × سعر)"):
            cl1, cl2 = st.columns(2)
            with cl1:
                st.markdown("**تنظيف شيت الموديول:**")
                skip_mod = st.number_input("تخطي صفوف علوية بالموديول:", 0, 30, 0, key="s_mod")
                df_mod = clean_dataframe(df_mod_raw, skip_rows=skip_mod, drop_empty=True)
                calc_mod = st.checkbox("حساب المبلغ بضرب (الكمية × السعر) في الموديول")
                if calc_mod:
                    q_col = st.selectbox("عمود الكمية:", df_mod.columns, key="q_m")
                    p_col = st.selectbox("عمود السعر:", df_mod.columns, key="p_m")
                    df_mod = clean_dataframe(df_mod, skip_rows=0, drop_empty=False, calc_mode="ضرب عمودين (الكمية × السعر)", col_a=q_col, col_b=p_col)
            with cl2:
                st.markdown("**تنظيف شيت القيود (GL):**")
                skip_gl = st.number_input("تخطي صفوف علوية بالقيود:", 0, 30, 0, key="s_gl")
                df_gl = clean_dataframe(df_gl_raw, skip_rows=skip_gl, drop_empty=True)

        st.subheader("⚙️ إعداد مفاتيح الربط والتجميع")
        c1, c2, c3 = st.columns(3)
        with c1:
            mod_key = st.selectbox("عمود رقم القيد في الموديول (مثل ENTRYSERIAL):", df_mod.columns)
            mod_val = st.selectbox("عمود المبلغ في الموديول:", df_mod.columns, index=len(df_mod.columns)-1 if 'الإجمالي_المحسوب' in df_mod.columns else 0)
        with c2:
            gl_key = st.selectbox("عمود رقم القيد في الحسابات (GL):", df_gl.columns)
            gl_val = st.selectbox("عمود المبلغ في الحسابات (DEBIT أو CREDIT):", df_gl.columns)
        with c3:
            tolerance = st.number_input("هامش كسور التقريب المسموح:", value=0.01, step=0.01)

        if st.button("🚀 تشغيل الفحص واستخراج الفروقات والتقرير الشامل", use_container_width=True):
            # تنظيف القيم الرقمية
            df_mod[mod_key] = pd.to_numeric(df_mod[mod_key], errors='coerce').fillna(0).astype(int)
            df_mod[mod_val] = pd.to_numeric(df_mod[mod_val], errors='coerce').fillna(0)
            df_gl[gl_key] = pd.to_numeric(df_gl[gl_key], errors='coerce').fillna(0).astype(int)
            df_gl[gl_val] = pd.to_numeric(df_gl[gl_val], errors='coerce').fillna(0)

            unposted_mod = df_mod[(df_mod[mod_key] == 0) & (df_mod[mod_val] != 0)]
            posted_mod = df_mod[df_mod[mod_key] != 0]

            # تجميع التفاصيل مع الاحتفاظ بعدد الأسطر المجمعة
            mod_agg = posted_mod.groupby(mod_key).agg(
                إجمالي_الموديول=(mod_val, 'sum'),
                عدد_أسطر_الموديول=(mod_val, 'count')
            ).reset_index()

            gl_agg = df_gl[df_gl[gl_key] != 0].groupby(gl_key).agg(
                إجمالي_القيود=(gl_val, 'sum'),
                عدد_أسطر_القيود=(gl_val, 'count')
            ).reset_index()

            merged = pd.merge(mod_agg, gl_agg, left_on=mod_key, right_on=gl_key, how='outer').fillna(0)
            merged['رقم_القيد'] = merged[mod_key].where(merged[mod_key] != 0, merged[gl_key]).astype(int)
            merged['إجمالي_الموديول'] = merged['إجمالي_الموديول'].round(2)
            merged['إجمالي_القيود'] = merged['إجمالي_القيود'].round(2)
            merged['الفرق (القيود - الموديول)'] = (merged['إجمالي_القيود'] - merged['إجمالي_الموديول']).round(2)

            # تحديد حالة كل قيد
            def classify_row(r):
                if abs(r['الفرق (القيود - الموديول)']) <= tolerance:
                    return "✅ متطابق تماماً"
                elif r['إجمالي_الموديول'] == 0:
                    return "👻 قيد يتيم بالحسابات بدون موديول"
                elif r['إجمالي_القيود'] == 0:
                    return "🛑 حركة موديول برقم قيد غير موجود بالحسابات"
                else:
                    return "⚠️ فرق في القيمة (استقطاع أو تعديل)"

            merged['تشخيص_الحركة'] = merged.apply(classify_row, axis=1)
            cols_order = ['رقم_القيد', 'إجمالي_الموديول', 'عدد_أسطر_الموديول', 'إجمالي_القيود', 'عدد_أسطر_القيود', 'الفرق (القيود - الموديول)', 'تشخيص_الحركة']
            merged = merged[cols_order].sort_values(by='رقم_القيد')

            mismatches = merged[(merged['الفرق (القيود - الموديول)'].abs() > tolerance) & (merged['إجمالي_الموديول'] != 0) & (merged['إجمالي_القيود'] != 0)]
            orphan_gl = merged[(merged['إجمالي_الموديول'] == 0) & (merged['إجمالي_القيود'] != 0)]

            # إجماليات وتشخيص كلي
            total_mod_all = (posted_mod[mod_val].sum() + unposted_mod[mod_val].sum()).round(2)
            total_gl_all = df_gl[gl_val].sum().round(2)
            net_diff = (total_gl_all - total_mod_all).round(2)
            who_higher = "متطابقان تماماً" if abs(net_diff) <= tolerance else ("الحسابات العامة (GL) أعلى من الموديول" if net_diff > 0 else "الموديول أعلى من الحسابات العامة (GL)")

            st.divider()
            st.subheader("📌 الملخص التنفيذي وتشخيص المطابقة")
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("إجمالي شيت الموديول بالكامل", f"{total_mod_all:,.2f}")
            m2.metric("إجمالي شيت القيود (GL)", f"{total_gl_all:,.2f}")
            m3.metric("صافي الفرق الكلي", f"{abs(net_diff):,.2f}")
            m4.metric("من الأعلى؟", who_higher)

            t_mis, t_orph, t_unp, t_all = st.tabs([
                f"⚠️ فروق المبالغ ({len(mismatches)})",
                f"👻 قيود بالحسابات بدون موديول ({len(orphan_gl)})",
                f"🛑 حركات موديول غير مرحلة ({len(unposted_mod)})",
                f"📋 إجماليات التفاصيل المجمعة ({len(merged)})"
            ])
            with t_mis:
                st.dataframe(mismatches, use_container_width=True)
            with t_orph:
                st.dataframe(orphan_gl, use_container_width=True)
            with t_unp:
                st.dataframe(unposted_mod, use_container_width=True)
            with t_all:
                st.dataframe(merged, use_container_width=True)

            # تجهيز ملف الإكسيل الشامل للتحميل
            summary_df = pd.DataFrame([
                {"البيان": "إجمالي الموديول (مرحل + غير مرحل)", "القيمة": total_mod_all},
                {"البيان": "إجمالي الحسابات العامة (GL)", "القيمة": total_gl_all},
                {"البيان": "صافي الفرق المالي", "القيمة": abs(net_diff)},
                {"البيان": "الطرف الأعلى بالفرق", "القيمة": who_higher},
                {"البيان": "عدد القيود التي بها اختلاف مبالغ", "القيمة": len(mismatches)},
                {"البيان": "عدد القيود اليتيمة بالحسابات (بدون موديول)", "القيمة": len(orphan_gl)},
                {"البيان": "عدد حركات الموديول غير المرحلة (بدون قيد)", "القيمة": len(unposted_mod)}
            ])

            report_bytes = io.BytesIO()
            with pd.ExcelWriter(report_bytes, engine='openpyxl') as writer:
                summary_df.to_excel(writer, sheet_name="1_ملخص_المطابقة_والتشخيص", index=False)
                mismatches.to_excel(writer, sheet_name="2_فروق_المبالغ", index=False)
                orphan_gl.to_excel(writer, sheet_name="3_قيود_يتيمة_بالحسابات", index=False)
                unposted_mod.to_excel(writer, sheet_name="4_حركات_موديول_غير_مرحلة", index=False)
                merged.to_excel(writer, sheet_name="5_إجماليات_التفاصيل_الكاملة", index=False)

            st.download_button(
                label="📥 تحميل تقرير المطابقة الشامل (Excel مكون من 5 شيتات بالملخص والفروقات وإجماليات التفاصيل)",
                data=report_bytes.getvalue(),
                file_name=f"HIS_Reconciliation_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

# ----------------- الوحدة الخامسة: الاستيراد والتصدير -----------------
elif menu == "🔄 5. الاستيراد والتصدير والسحابة":
    st.header("🔄 محرك الاستيراد والتصدير الموحد لقاعدة المعرفة (Excel Sync)")
    excel_data = export_db_to_bytes()
    st.download_button(
        label="📥 تحميل قاعدة بيانات التطبيق الكاملة (Excel Template)",
        data=excel_data,
        file_name=f"HIS_Audit_KnowledgeBase_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    st.divider()
    uploaded_db = st.file_uploader("📤 ارفع ملف الإكسيل المحدث لاستيراد الجداول والأعمدة والاستعلامات:", type=['xlsx'])
    if uploaded_db:
        if st.button("تأكيد استيراد وتحديث القاعدة"):
            st.session_state.df_tables = pd.read_excel(uploaded_db, sheet_name="Tables")
            st.session_state.df_columns = pd.read_excel(uploaded_db, sheet_name="Columns")
            st.session_state.df_sql = pd.read_excel(uploaded_db, sheet_name="SQL_Repo")
            st.session_state.df_logs = pd.read_excel(uploaded_db, sheet_name="Audit_Logs")
            export_db_to_bytes()
            st.success("تم التحديث بنجاح!")
            st.rerun()
