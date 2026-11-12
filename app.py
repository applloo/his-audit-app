# -*- coding: utf-8 -*-
# =====================================================================
#  مراجع قواعد بيانات المستشفيات - HIS Audit  (الإصدار 2.0)
#  صاحب التطبيق: Ehab Ibrahim
#  الفكرة: أي جملة SQL أو جدول جديد تضيفه، التطبيق يقراه لوحده
#  ويوزع الجداول والأعمدة والعلاقات على القاموس ويربطهم ببعض.
# =====================================================================
import re
import io
import base64
from datetime import datetime

import pandas as pd
import requests
import streamlit as st

st.set_page_config(page_title="مراجع قواعد بيانات المستشفيات - HIS Audit",
                   layout="wide", page_icon="🏥")

DB_FILE = "his_knowledge_base.xlsx"

# ---------------------------------------------------------------------
#  شكل الجداول (أسماء الأعمدة) - نفس أسماء النسخة القديمة عشان ملفك القديم يشتغل
# ---------------------------------------------------------------------
TABLE_COLS = ["اسم_الجدول", "الموديول", "النوع", "حالة_التأكد", "الوصف"]
COL_COLS = ["اسم_الجدول", "اسم_العمود", "نوع_البيانات", "نوع_المفتاح", "حالة_التأكد", "الوصف"]
SQL_COLS = ["رقم_الاستعلام", "عنوان_المشكلة", "الجداول_المرتبطة", "الاصدار",
            "شرح_الحل", "كود_SQL", "تاريخ_التحديث"]
REL_COLS = ["الجدول_الأول", "العمود_الأول", "الجدول_الثاني", "العمود_الثاني", "المصدر"]
LOG_COLS = ["الجهة_أو_المورد", "قيمة_الفرق", "من_الأعلى", "الحالة", "السبب_الجذري", "التاريخ"]
HIST_COLS = ["رقم_الاستعلام", "الاصدار", "عنوان_المشكلة", "كود_SQL", "تاريخ_الحفظ"]

SHEETS = {
    "Tables": ("df_tables", TABLE_COLS),
    "Columns": ("df_columns", COL_COLS),
    "SQL_Repo": ("df_sql", SQL_COLS),
    "Relations": ("df_rel", REL_COLS),
    "Audit_Logs": ("df_logs", LOG_COLS),
    "SQL_History": ("df_hist", HIST_COLS),
}

AUTO_STATUS = "مكتشف تلقائياً من استعلام"
MODULES = ["غير محدد", "المخازن", "المشتريات", "الخزينة والبنوك", "الحسابات العامة",
           "الضرائب والاستقطاعات", "المرضى والفواتير", "شئون العاملين", "أخرى"]


# =====================================================================
#  1) قارئ جمل SQL  (القلب بتاع التطبيق)
# =====================================================================
SQL_KEYWORDS = set("""
SELECT FROM WHERE AND OR NOT IN IS NULL AS ON JOIN LEFT RIGHT FULL OUTER INNER CROSS NATURAL
GROUP BY ORDER HAVING UNION ALL DISTINCT CASE WHEN THEN ELSE END WITH BETWEEN LIKE EXISTS
ASC DESC INSERT INTO VALUES UPDATE SET DELETE CREATE TABLE ALTER DROP PRIMARY KEY FOREIGN
REFERENCES DEFAULT CONSTRAINT UNIQUE CHECK INDEX VIEW NULLS FIRST LAST OVER PARTITION ROWNUM
ROWID SYSDATE SYSTIMESTAMP DUAL LEVEL CONNECT PRIOR START MINUS INTERSECT FETCH ROWS ONLY
OFFSET NEXT ROW USING LATERAL YEAR MONTH DAY HOUR MINUTE SECOND GENERATED ALWAYS IDENTITY
COMMIT ROLLBACK TRUE FALSE ANY SOME ESCAPE MERGE MATCHED REPLACE OR FORCE NOCOPY
""".split())

_ID = r"[A-Z_][A-Z0-9_$#]*"


def _clean_sql(sql: str) -> str:
    s = re.sub(r"/\*.*?\*/", " ", str(sql), flags=re.S)
    s = re.sub(r"--[^\n]*", " ", s)
    s = re.sub(r"'(?:[^']|'')*'", "''", s)        # النصوص بين علامات '
    s = re.sub(r'"[^"]*"', " ", s)                # الأسماء العربي بين " "
    s = re.sub(r":\w+", " ", s)                   # متغيرات :X
    s = s.upper()
    s = re.sub(r"\bEXTRACT\s*\(\s*\w+\s+FROM\b", "EXTRACT(", s)
    s = re.sub(r"\bTRIM\s*\(\s*(?:LEADING|TRAILING|BOTH)?[^()]*?\bFROM\b", "TRIM(", s)
    return s


def parse_sql(sql: str) -> dict:
    """يرجع: الجداول (مع اختصاراتها) + الأعمدة + العلاقات اللي في الجملة."""
    s = _clean_sql(sql)

    cte_names = set(re.findall(rf"\b({_ID})\s+AS\s*\(", s))
    derived = {a for a in re.findall(rf"\)\s*(?:AS\s+)?({_ID})", s) if a not in SQL_KEYWORDS}
    as_aliases = {a for a in re.findall(rf"\bAS\s+({_ID})", s) if a not in SQL_KEYWORDS}

    tables, alias_map = {}, {}

    def add_table(raw_name, alias):
        name = raw_name.split(".")[-1]
        if not name or name in SQL_KEYWORDS or name in cte_names or name == "DUAL":
            return
        tables.setdefault(name, set())
        alias_map[name] = name
        if alias and alias not in SQL_KEYWORDS:
            tables[name].add(alias)
            alias_map[alias] = name

    stop = (r"(?=\bWHERE\b|\bGROUP\b|\bORDER\b|\bHAVING\b|\bUNION\b|\bMINUS\b|\bINTERSECT\b"
            r"|\bCONNECT\b|\bSTART\b|\bFETCH\b|\bSELECT\b|\b(?:LEFT|RIGHT|FULL|INNER|CROSS|NATURAL)\b"
            r"|\bJOIN\b|\bON\b|\bSET\b|\bVALUES\b|\(|\)|;|$)")
    for m in re.finditer(r"\b(?:FROM|INTO|UPDATE)\b(.*?)" + stop, s, flags=re.S):
        for piece in m.group(1).split(","):
            tok = piece.split()
            if not tok:
                continue
            alias = tok[-1] if len(tok) > 1 else None
            add_table(tok[0], alias)
    for m in re.finditer(rf"\bJOIN\s+({_ID}(?:\.{_ID})?)(?:\s+(?:AS\s+)?({_ID}))?", s):
        add_table(m.group(1), m.group(2))

    def resolve(q, col):
        if q in cte_names or q in derived:
            if col in as_aliases or q not in alias_map:
                return None
        return alias_map.get(q)

    columns = set()
    for q, col in re.findall(rf"\b({_ID})\.({_ID})\b", s):
        t = resolve(q, col)
        if t and col not in SQL_KEYWORDS:
            columns.add((t, col, "مؤكد من الجملة"))

    # جملة فيها جدول واحد بس: الأعمدة اللي من غير اختصار تبقى بتاعته
    if len(tables) == 1 and not cte_names and not derived - as_aliases:
        only = next(iter(tables))
        body = re.sub(rf"\b{_ID}\.{_ID}\b", " ", s)
        funcs = set(re.findall(rf"\b({_ID})\s*\(", body))
        skip = SQL_KEYWORDS | funcs | as_aliases | set(alias_map) | {"TO_DATE", "TO_CHAR"}
        for ident in set(re.findall(rf"\b{_ID}\b", body)):
            if ident not in skip and not ident.isdigit():
                columns.add((only, ident, "مكتشف - يحتاج تأكيد"))

    relations = set()
    for q1, c1, q2, c2 in re.findall(rf"\b({_ID})\.({_ID})\s*=\s*({_ID})\.({_ID})\b", s):
        t1, t2 = resolve(q1, c1), resolve(q2, c2)
        if t1 and t2 and t1 != t2:
            a, b = sorted([(t1, c1), (t2, c2)])
            relations.add((a[0], a[1], b[0], b[1]))

    return {"tables": {k: sorted(v) for k, v in tables.items()},
            "columns": sorted(columns), "relations": sorted(relations)}


def _split_top_level(body: str):
    parts, depth, cur = [], 0, ""
    for ch in body:
        if ch == "(":
            depth += 1
        elif ch == ")":
            depth -= 1
        if ch == "," and depth == 0:
            parts.append(cur)
            cur = ""
        else:
            cur += ch
    if cur.strip():
        parts.append(cur)
    return parts


def parse_table_definition(text: str, table_name: str = "") -> dict:
    """يقرا CREATE TABLE أو ناتج أمر DESC اللي بتنسخه من Toad/SQL Developer."""
    out = {"tables": {}, "relations": []}
    up = re.sub(r"--[^\n]*", " ", str(text)).upper().replace('"', "")

    for m in re.finditer(rf"CREATE\s+(?:GLOBAL\s+TEMPORARY\s+)?TABLE\s+({_ID}(?:\.{_ID})?)\s*\(", up):
        tname = m.group(1).split(".")[-1]
        i, depth = m.end(), 1
        while i < len(up) and depth:
            depth += {"(": 1, ")": -1}.get(up[i], 0)
            i += 1
        cols = {}
        for part in _split_top_level(up[m.end():i - 1]):
            p = part.strip()
            if not p:
                continue
            pk = re.match(r"(?:CONSTRAINT\s+\S+\s+)?PRIMARY\s+KEY\s*\(([^)]*)\)", p)
            fk = re.match(rf"(?:CONSTRAINT\s+\S+\s+)?FOREIGN\s+KEY\s*\(([^)]*)\)\s*REFERENCES\s+({_ID}(?:\.{_ID})?)\s*(?:\(([^)]*)\))?", p)
            if pk:
                for c in pk.group(1).split(","):
                    if c.strip() in cols:
                        cols[c.strip()][1] = "PK"
                continue
            if fk:
                c, rt = fk.group(1).split(",")[0].strip(), fk.group(2).split(".")[-1]
                rc = (fk.group(3) or c).split(",")[0].strip()
                if c in cols:
                    cols[c][1] = "FK"
                out["relations"].append((tname, c, rt, rc))
                continue
            if re.match(r"(CONSTRAINT|UNIQUE|CHECK|INDEX)\b", p):
                continue
            tok = p.split()
            cname = tok[0]
            dtype = re.split(r"[\s(]", tok[1])[0] if len(tok) > 1 else "غير معروف"
            key = "PK" if "PRIMARY KEY" in p else ("FK" if "REFERENCES" in p else "Normal")
            cols[cname] = [dtype, key]
            r = re.search(rf"REFERENCES\s+({_ID}(?:\.{_ID})?)\s*(?:\(\s*({_ID})\s*\))?", p)
            if r:
                out["relations"].append((tname, cname, r.group(1).split(".")[-1], r.group(2) or cname))
        out["tables"][tname] = [(c, v[0], v[1]) for c, v in cols.items()]

    if out["tables"]:
        return out

    # ناتج DESC: كل سطر = اسم العمود ... النوع
    tname = table_name.strip().upper()
    if not tname:
        return out
    cols = []
    for line in up.splitlines():
        tok = line.split()
        if not tok or tok[0] in ("NAME", "COLUMN_NAME") or set(tok[0]) <= set("-"):
            continue
        if not re.fullmatch(_ID, tok[0]):
            continue
        dtype = re.split(r"[\s(]", tok[-1])[0] if len(tok) > 1 else "غير معروف"
        cols.append((tok[0], dtype, "Normal"))
    out["tables"][tname] = cols
    return out


# =====================================================================
#  2) البيانات الأساسية المعتمدة
# =====================================================================
def get_seed_data():
    tables = [
        ["CSHCASHUNITTRANS", "الخزينة والبنوك", "Master", "مؤكد ومجرب 100%", "إيصالات الدفع النقدية والشيكات وسندات الصرف والقبض للموردين والمرضى"],
        ["MMSTORETRANSACTIONS", "المخازن", "Master", "مؤكد ومجرب 100%", "الرأس الأساسي لحركات المخازن (أذون الإضافة، الفواتير، المرتجعات)"],
        ["MMSTORETRANSITEMS", "المخازن", "Details", "مؤكد ومجرب 100%", "تفاصيل أصناف حركات المخازن، الكميات والأسعار الفعلية لكل صنف"],
        ["GLTEMPGENERALJOURNAL", "الحسابات العامة", "Master", "مؤكد ومجرب 100%", "رأس قيود اليومية العامة، تاريخ القيد الفعلي والسنة المالية ونوع اليومية"],
        ["GLTEMPGENJOURNACCOUNTS", "الحسابات العامة", "Details", "مؤكد ومجرب 100%", "تفاصيل أطراف القيد المحاسبي (الحسابات المدينة والدائنة وقيمها)"],
        ["GLTEMPGENERALJOURNAL_TYPES", "الحسابات العامة", "Lookup", "مؤكد من شجرة القاعدة", "أنواع القيود اليومية (مثل كود 112 = تعديل فاتورة مخزن)"],
        ["GLTEMPGENERALJOURNAL_DETAILS", "الحسابات العامة", "Details", "مؤكد من شجرة القاعدة", "تفاصيل إضافية لقيود اليومية العامة"],
        ["MMPURCHASINGORDER", "المشتريات", "Master", "مرجعي - يحتاج فحص الأعمدة", "رأس أوامر الشراء الصادرة للموردين"],
        ["TAX", "الضرائب والاستقطاعات", "Master", "مرجعي - يحتاج فحص الأعمدة", "جدول تعريف وربط الضرائب ونسب الخصم والاستقطاع"],
    ]
    cols = [
        ["CSHCASHUNITTRANS", "RECEIPTNO", "NUMBER", "Normal", "مؤكد ومجرب 100%", "رقم الإيصال الفعلي المطبوع والظاهر للمستخدم"],
        ["CSHCASHUNITTRANS", "TRANSSERIAL", "NUMBER", "PK", "مؤكد ومجرب 100%", "المسلسل الداخلي لحركة الخزينة في القاعدة"],
        ["CSHCASHUNITTRANS", "TRANSDATE", "DATE", "Normal", "مؤكد ومجرب 100%", "تاريخ تحرير إيصال الدفع أو القبض"],
        ["CSHCASHUNITTRANS", "SUPPLIERCD", "NUMBER", "FK", "مؤكد ومجرب 100%", "كود المورد في الموديول (مثل 504 للفاتح، 583 لكريم فارما)"],
        ["CSHCASHUNITTRANS", "TOTALAMOUNT", "NUMBER", "Normal", "مؤكد ومجرب 100%", "إجمالي قيمة الإيصال في الموديول"],
        ["CSHCASHUNITTRANS", "ENTRYSERIAL", "NUMBER", "FK", "مؤكد ومجرب 100%", "رقم القيد المحاسبي المرتبط في الحسابات العامة (0 أو فارغ = غير مرحل)"],
        ["CSHCASHUNITTRANS", "TAX_DEDUCTION", "NUMBER", "Normal", "استرشادي", "قيمة الاستقطاع الضريبي (ضرائب المنبع) على الإيصال"],
        ["CSHCASHUNITTRANS", "DISC_VALUE", "NUMBER", "Normal", "استرشادي", "قيمة الخصم المكتسب على الإيصال"],
        ["MMSTORETRANSACTIONS", "GEN_SERIAL", "NUMBER", "PK", "مؤكد ومجرب 100%", "المسلسل العام لحركة المخزن ومفتاح الربط مع جدول الأصناف"],
        ["MMSTORETRANSACTIONS", "TRANSDATE", "DATE", "Normal", "مؤكد ومجرب 100%", "تاريخ حركة المخزن أو الفاتورة"],
        ["MMSTORETRANSACTIONS", "SUPPLIERCD", "NUMBER", "FK", "مؤكد ومجرب 100%", "كود المورد صاحب الفاتورة المخزنية"],
        ["MMSTORETRANSACTIONS", "ENTRYSERIAL", "NUMBER", "FK", "مؤكد ومجرب 100%", "رقم قيد اليومية المرتبط بالفاتورة في الحسابات العامة"],
        ["MMSTORETRANSITEMS", "GEN_SERIAL", "NUMBER", "FK", "مؤكد ومجرب 100%", "مفتاح الربط مع رأس إذن المخزن MMSTORETRANSACTIONS"],
        ["MMSTORETRANSITEMS", "ACTQUANTITY", "NUMBER", "Normal", "مؤكد ومجرب 100%", "الكمية الفعلية المستلمة من الصنف"],
        ["MMSTORETRANSITEMS", "UNITCOSTPRICE", "NUMBER", "Normal", "مؤكد ومجرب 100%", "سعر تكلفة الوحدة الواحدة"],
        ["GLTEMPGENERALJOURNAL", "ENTRYSERIAL", "NUMBER", "PK", "مؤكد ومجرب 100%", "المسلسل الرئيسي للقيد في الحسابات العامة (مثل 13389)"],
        ["GLTEMPGENERALJOURNAL", "ENTRYACTDATE", "DATE", "Normal", "مؤكد ومجرب 100%", "التاريخ والوقت الفعلي لإنشاء القيد وتأثيره في اليومية"],
        ["GLTEMPGENERALJOURNAL", "YEARCODE", "NUMBER", "Normal", "مؤكد ومجرب 100%", "السنة المالية المقيد عليها القيد (2024، 2025)"],
        ["GLTEMPGENJOURNACCOUNTS", "ENTRYSERIAL", "NUMBER", "FK", "مؤكد ومجرب 100%", "مفتاح الربط مع رأس القيد"],
        ["GLTEMPGENJOURNACCOUNTS", "ACCOUNTCODE", "VARCHAR2", "FK", "مؤكد ومجرب 100%", "كود الحساب المالي (202006 متبوعاً بكود المورد مثل 2020060504)"],
        ["GLTEMPGENJOURNACCOUNTS", "DEBIT", "NUMBER", "Normal", "مؤكد ومجرب 100%", "القيمة المدينة (تخفض رصيد المورد)"],
        ["GLTEMPGENJOURNACCOUNTS", "CREDIT", "NUMBER", "Normal", "مؤكد ومجرب 100%", "القيمة الدائنة (تزيد رصيد المورد)"],
    ]
    q = [
        [1, "كشف فروق إيصالات الدفع والاستقطاعات الساقطة", "", "v1.2 (مؤكد)",
         "يربط الإيصالات بالقيود عبر ENTRYSERIAL ويقارن إجمالي الإيصال بمدين المورد لكشف الضرائب الساقطة.",
         """SELECT
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
ORDER BY C.TRANSDATE;""", "2026-10-30"],
        [2, "كشف إيصالات الدفع غير المرحلة (Unposted Receipts)", "", "v1.1 (مؤكد)",
         "يستخرج الإيصالات المحفوظة في الموديول ولم يتولد لها قيد محاسبي.",
         """SELECT
    RECEIPTNO AS "رقم الإيصال",
    TRANSSERIAL AS "رقم الحركة",
    TO_CHAR(TRANSDATE, 'YYYY-MM-DD') AS "التاريخ",
    TOTALAMOUNT AS "المبلغ"
FROM CSHCASHUNITTRANS
WHERE TRANSDATE BETWEEN TO_DATE('01/01/2025', 'DD/MM/YYYY') AND TO_DATE('31/12/2025', 'DD/MM/YYYY')
  AND SUPPLIERCD = 504
  AND (ENTRYSERIAL IS NULL OR ENTRYSERIAL = 0);""", "2026-10-30"],
        [3, "مطابقة أذون الإضافة المخزنية مع القيود الدائنة (مع فلترة كسور التقريب)", "", "v1.2 (مؤكد)",
         "يجمع (الكمية × السعر) لكل إذن ويقارنه بدائن المورد في القيد، مع ROUND لاستبعاد فروق الكسور.",
         """SELECT
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
HAVING ABS(ROUND(A.CREDIT - NVL(SUM(I.ACTQUANTITY * I.UNITCOSTPRICE), 0), 2)) > 0.01;""", "2026-10-30"],
        [4, "القناص: كشف القيود اليتيمة وفروق الفترات المالية (Cross-Year Cutoff)", "", "v1.2 (مؤكد ومجرب)",
         "يكشف القيود المسجلة في الحسابات العامة وليس لها مستند في إيصالات أو مخازن.",
         """SELECT
    G.ENTRYACTDATE AS "التاريخ",
    G.ENTRYSERIAL AS "رقم القيد",
    A.DEBIT AS "مدين",
    A.CREDIT AS "دائن"
FROM GLTEMPGENJOURNACCOUNTS A
JOIN GLTEMPGENERALJOURNAL G ON A.ENTRYSERIAL = G.ENTRYSERIAL
WHERE G.YEARCODE = 2025
  AND A.ACCOUNTCODE = '2020060504'
  AND A.ENTRYSERIAL NOT IN (SELECT ENTRYSERIAL FROM CSHCASHUNITTRANS WHERE ENTRYSERIAL IS NOT NULL)
  AND A.ENTRYSERIAL NOT IN (SELECT ENTRYSERIAL FROM MMSTORETRANSACTIONS WHERE ENTRYSERIAL IS NOT NULL);""", "2026-10-30"],
        [5, "مطابقة شاملة لكل الموردين: القيود × المخازن × الخزينة (سنة 2025)", "", "v2.0 (مصحح)",
         "بيعمل قايمة بكل (مورد + رقم قيد) من الثلاث جهات، وبعدين يجيب قيمة كل جهة جنبها، عشان كل حركة تظهر في صف واحد ويبان الفرق صح.",
         """WITH G AS (
    SELECT TO_NUMBER(SUBSTR(A.ACCOUNTCODE, 7)) AS SUPPLIERCD, A.ENTRYSERIAL,
           TRUNC(J.ENTRYACTDATE) AS ENTRYACTDATE,
           SUM(A.CREDIT) AS GL_CREDIT, SUM(A.DEBIT) AS GL_DEBIT
    FROM GLTEMPGENJOURNACCOUNTS A
    JOIN GLTEMPGENERALJOURNAL J ON A.ENTRYSERIAL = J.ENTRYSERIAL
    WHERE J.YEARCODE = 2025
      AND A.ACCOUNTCODE LIKE '202006%' AND LENGTH(TRIM(A.ACCOUNTCODE)) > 6
    GROUP BY TO_NUMBER(SUBSTR(A.ACCOUNTCODE, 7)), A.ENTRYSERIAL, TRUNC(J.ENTRYACTDATE)
),
M AS (
    SELECT M.SUPPLIERCD, M.ENTRYSERIAL, TRUNC(M.TRANSDATE) AS TRANSDATE,
           SUM(I.ACTQUANTITY * I.UNITCOSTPRICE) AS MM_TOTAL
    FROM MMSTORETRANSACTIONS M
    JOIN MMSTORETRANSITEMS I ON M.GEN_SERIAL = I.GEN_SERIAL
    WHERE M.ENTRYSERIAL IS NOT NULL AND EXTRACT(YEAR FROM M.TRANSDATE) = 2025
    GROUP BY M.SUPPLIERCD, M.ENTRYSERIAL, TRUNC(M.TRANSDATE)
),
C AS (
    SELECT SUPPLIERCD, ENTRYSERIAL, TRUNC(TRANSDATE) AS TRANSDATE, SUM(TOTALAMOUNT) AS CSH_TOTAL
    FROM CSHCASHUNITTRANS
    WHERE ENTRYSERIAL IS NOT NULL AND EXTRACT(YEAR FROM TRANSDATE) = 2025
    GROUP BY SUPPLIERCD, ENTRYSERIAL, TRUNC(TRANSDATE)
),
K AS (
    SELECT SUPPLIERCD, ENTRYSERIAL FROM G
    UNION SELECT SUPPLIERCD, ENTRYSERIAL FROM M
    UNION SELECT SUPPLIERCD, ENTRYSERIAL FROM C
)
SELECT K.SUPPLIERCD AS "رقم المورد",
       TO_CHAR(COALESCE(G.ENTRYACTDATE, M.TRANSDATE, C.TRANSDATE), 'YYYY-MM-DD') AS "تاريخ الحركة",
       K.ENTRYSERIAL AS "رقم القيد",
       NVL(M.MM_TOTAL, 0) AS "إجمالي فواتير المخازن",
       NVL(G.GL_CREDIT, 0) AS "دائن بالقيود (GL)",
       NVL(G.GL_CREDIT, 0) - NVL(M.MM_TOTAL, 0) AS "فرق الدائن",
       NVL(C.CSH_TOTAL, 0) AS "إجمالي إيصالات الخزينة",
       NVL(G.GL_DEBIT, 0) AS "مدين بالقيود (GL)",
       NVL(G.GL_DEBIT, 0) - NVL(C.CSH_TOTAL, 0) AS "فرق المدين"
FROM K
LEFT JOIN G ON G.SUPPLIERCD = K.SUPPLIERCD AND G.ENTRYSERIAL = K.ENTRYSERIAL
LEFT JOIN M ON M.SUPPLIERCD = K.SUPPLIERCD AND M.ENTRYSERIAL = K.ENTRYSERIAL
LEFT JOIN C ON C.SUPPLIERCD = K.SUPPLIERCD AND C.ENTRYSERIAL = K.ENTRYSERIAL
ORDER BY 1, 2, 3;""", "2026-11-09"],
    ]
    logs = [["الفاتح فارما (504) / كريم فارما (583)", 2463.72, "الموديول أعلى من الحسابات العامة",
             "تم اكتشاف السبب الجذري",
             "تعديل فواتير مخزن (نوع القيد 112) تخص ديسمبر 2024 في يوم 8 يناير 2025 بواسطة GL.HHUSSEIN (قيود 13389، 13390، 13393) مما رمى القيد في 2025 بينما الحركة في موديول 2024.",
             "2026-10-30"]]
    return {
        "df_tables": pd.DataFrame(tables, columns=TABLE_COLS),
        "df_columns": pd.DataFrame(cols, columns=COL_COLS),
        "df_sql": pd.DataFrame(q, columns=SQL_COLS),
        "df_rel": pd.DataFrame(columns=REL_COLS),
        "df_logs": pd.DataFrame(logs, columns=LOG_COLS),
        "df_hist": pd.DataFrame(columns=HIST_COLS),
    }


# =====================================================================
#  3) الحفظ والتحميل (ملف إكسيل + مزامنة GitHub أوتوماتيك)
# =====================================================================
def _secret(name):
    try:
        return st.secrets[name]
    except Exception:
        return None


def _gh():
    token, repo = _secret("GITHUB_TOKEN"), _secret("GITHUB_REPO")
    if not token or not repo:
        return None
    return (f"https://api.github.com/repos/{repo}/contents/{DB_FILE}",
            {"Authorization": f"token {token}", "Accept": "application/vnd.github.v3+json"})


def _fix(df, cols):
    df = df.copy() if df is not None else pd.DataFrame(columns=cols)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    df = df[cols]
    for c in cols:
        if df[c].dtype == object or str(df[c].dtype).startswith("str"):
            df[c] = df[c].fillna("").astype(str).replace("nan", "")
    return df.reset_index(drop=True)


def _from_excel_bytes(data):
    book = pd.read_excel(io.BytesIO(data), sheet_name=None)
    out = {}
    for sheet, (key, cols) in SHEETS.items():
        out[key] = _fix(book.get(sheet), cols)
    if "قيمة_الفرق" in out["df_logs"]:
        out["df_logs"]["قيمة_الفرق"] = pd.to_numeric(out["df_logs"]["قيمة_الفرق"], errors="coerce").fillna(0)
    out["df_sql"]["رقم_الاستعلام"] = pd.to_numeric(out["df_sql"]["رقم_الاستعلام"], errors="coerce").fillna(0).astype(int)
    return out


def load_db():
    try:
        with open(DB_FILE, "rb") as f:
            return _from_excel_bytes(f.read()), "ملف محلي"
    except Exception:
        pass
    gh = _gh()
    if gh:
        try:
            url, headers = gh
            r = requests.get(url, headers={**headers, "Accept": "application/vnd.github.raw"}, timeout=20)
            if r.status_code == 200:
                return _from_excel_bytes(r.content), "GitHub"
        except Exception:
            pass
    return get_seed_data(), "البيانات الأساسية"


def db_bytes():
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        for sheet, (key, cols) in SHEETS.items():
            _fix(st.session_state[key], cols).to_excel(w, sheet_name=sheet, index=False)
    return out.getvalue()


def push_github(data):
    gh = _gh()
    if not gh:
        return None
    url, headers = gh
    try:
        r = requests.get(url, headers=headers, timeout=20)
        sha = r.json().get("sha") if r.status_code == 200 else None
        body = {"message": f"Auto-sync {datetime.now():%Y-%m-%d %H:%M}",
                "content": base64.b64encode(data).decode()}
        if sha:
            body["sha"] = sha
        return requests.put(url, headers=headers, json=body, timeout=30).status_code in (200, 201)
    except Exception:
        return False


def save_all(msg="تم الحفظ"):
    """بتتنده بعد أي تعديل: تحفظ محلي وتبعت لـ GitHub لوحدها."""
    data = db_bytes()
    try:
        with open(DB_FILE, "wb") as f:
            f.write(data)
    except Exception:
        pass
    ok = push_github(data)
    if ok is True:
        st.session_state.sync_status = f"☁️ آخر مزامنة: {datetime.now():%H:%M}"
        st.toast(f"✅ {msg} - واتبعت للسحابة")
    elif ok is False:
        st.session_state.sync_status = "⚠️ المزامنة فشلت - اتحفظ على الجهاز بس"
        st.toast(f"⚠️ {msg} محلياً بس (المزامنة مع GitHub فشلت)")
    else:
        st.session_state.sync_status = "💾 محفوظ محلياً (GitHub مش متوصل)"
        st.toast(f"✅ {msg}")


# =====================================================================
#  4) التسكين الأوتوماتيك: توزيع اللي اتقرا على القاموس
# =====================================================================
S = st.session_state


def upsert_table(name, status=AUTO_STATUS, module="غير محدد", ttype="غير محدد", desc=""):
    name = str(name).strip().upper()
    if not name or name in set(S.df_tables["اسم_الجدول"]):
        return False
    S.df_tables = pd.concat([S.df_tables, pd.DataFrame([[name, module, ttype, status, desc]], columns=TABLE_COLS)],
                            ignore_index=True)
    return True


def upsert_column(table, col, dtype="", key="", status="", desc=""):
    table, col = str(table).upper().strip(), str(col).upper().strip()
    df = S.df_columns
    mask = (df["اسم_الجدول"] == table) & (df["اسم_العمود"] == col)
    if mask.any():
        idx = df[mask].index[0]
        if dtype and str(df.at[idx, "نوع_البيانات"]).strip() in ("", "غير معروف"):
            df.at[idx, "نوع_البيانات"] = dtype
        if key in ("PK", "FK") and str(df.at[idx, "نوع_المفتاح"]).strip() in ("", "Normal"):
            df.at[idx, "نوع_المفتاح"] = key
        if desc and not str(df.at[idx, "الوصف"]).strip():
            df.at[idx, "الوصف"] = desc
        return False
    S.df_columns = pd.concat([df, pd.DataFrame([[table, col, dtype or "غير معروف", key or "Normal",
                                                  status or AUTO_STATUS, desc]], columns=COL_COLS)],
                             ignore_index=True)
    return True


def upsert_relation(t1, c1, t2, c2, source):
    a, b = sorted([(t1, c1), (t2, c2)])
    df = S.df_rel
    dup = ((df["الجدول_الأول"] == a[0]) & (df["العمود_الأول"] == a[1]) &
           (df["الجدول_الثاني"] == b[0]) & (df["العمود_الثاني"] == b[1])).any()
    if dup:
        return False
    S.df_rel = pd.concat([df, pd.DataFrame([[a[0], a[1], b[0], b[1], source]], columns=REL_COLS)],
                         ignore_index=True)
    return True


def absorb_sql(sql_text, source):
    """يقرا الجملة ويضيف أي جدول/عمود/علاقة جديدة. يرجع ملخص باللي اتضاف."""
    p = parse_sql(sql_text)
    rep = {"tables": [], "columns": [], "relations": [], "all_tables": sorted(p["tables"])}
    for t in p["tables"]:
        if upsert_table(t, status=f"{AUTO_STATUS} ({source})"):
            rep["tables"].append(t)
    for t, c, how in p["columns"]:
        st_txt = f"{AUTO_STATUS} ({source})" if how == "مؤكد من الجملة" else f"مكتشف - يحتاج تأكيد ({source})"
        if upsert_column(t, c, status=st_txt):
            rep["columns"].append(f"{t}.{c}")
    for t1, c1, t2, c2 in p["relations"]:
        if upsert_relation(t1, c1, t2, c2, source):
            rep["relations"].append(f"{t1}.{c1} = {t2}.{c2}")
            for t, c in ((t1, c1), (t2, c2)):
                upsert_column(t, c, key="FK")
    return rep


def absorb_definition(parsed, module, ttype, status, desc, source):
    rep = {"tables": [], "columns": [], "relations": []}
    for t, cols in parsed["tables"].items():
        if upsert_table(t, status=status, module=module, ttype=ttype, desc=desc):
            rep["tables"].append(t)
        for c, dtype, key in cols:
            if upsert_column(t, c, dtype=dtype, key=key, status=status):
                rep["columns"].append(f"{t}.{c}")
    for t1, c1, t2, c2 in parsed["relations"]:
        upsert_table(t2, status=f"{AUTO_STATUS} ({source})")
        if upsert_relation(t1, c1, t2, c2, source):
            rep["relations"].append(f"{t1}.{c1} → {t2}.{c2}")
    return rep


def resync_all_queries():
    total = {"tables": 0, "columns": 0, "relations": 0}
    for i, row in S.df_sql.iterrows():
        rep = absorb_sql(row["كود_SQL"], f"#{int(row['رقم_الاستعلام'])}")
        S.df_sql.at[i, "الجداول_المرتبطة"] = ", ".join(rep["all_tables"])
        for k in total:
            total[k] += len(rep[k])
    return total


def show_report(rep):
    if not (rep["tables"] or rep["columns"] or rep["relations"]):
        st.info("مفيش حاجة جديدة، كل الجداول والأعمدة دي موجودة في القاموس قبل كده 👍")
        return
    st.success(f"اتسكن أوتوماتيك: {len(rep['tables'])} جدول جديد، "
               f"{len(rep['columns'])} عمود جديد، {len(rep['relations'])} علاقة ربط جديدة")
    if rep["tables"]:
        st.write("📁 **جداول جديدة:** " + "، ".join(rep["tables"]))
    if rep["columns"]:
        st.write("🔑 **أعمدة جديدة:** " + "، ".join(rep["columns"]))
    if rep["relations"]:
        st.write("🔗 **علاقات جديدة:** " + "، ".join(rep["relations"]))


def queries_using(table):
    t = table.upper()
    df = S.df_sql
    mask = df["الجداول_المرتبطة"].str.upper().str.split(r"\s*,\s*").apply(lambda l: t in l if isinstance(l, list) else False)
    mask |= df["كود_SQL"].str.upper().str.contains(rf"\b{re.escape(t)}\b", regex=True)
    return df[mask]


def next_version(v):
    m = re.search(r"v?(\d+)\.(\d+)", str(v))
    return f"v{m.group(1)}.{int(m.group(2)) + 1}" if m else "v1.1"


# =====================================================================
#  5) تحميل البيانات أول مرة
# =====================================================================
if "initialized" not in S:
    data, src = load_db()
    for k, v in data.items():
        S[k] = v
    S.load_source = src
    S.sync_status = f"📂 اتحمل من: {src}"
    tot = resync_all_queries()          # يربط الاستعلامات القديمة بالقاموس
    if sum(tot.values()):
        try:
            with open(DB_FILE, "wb") as f:
                f.write(db_bytes())
        except Exception:
            pass
    S.initialized = True


# =====================================================================
#  6) الشكل (عربي + خط كبير مريح للعين)
# =====================================================================
with st.sidebar:
    st.title("🏥 نظام مراجعة HIS")
    dark_mode = st.toggle("🌙 الوضع الليلي", value=True)
    big_font = st.toggle("🔠 خط كبير مريح للعين", value=True)

bg = "#0e1117" if dark_mode else "#ffffff"
fg = "#fafafa" if dark_mode else "#1f2937"
card = "#1e2530" if dark_mode else "#f3f4f6"
fsize = "19px" if big_font else "15px"

st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700&display=swap');
html, body, [class*="css"] {{ font-size: {fsize}; }}
.stApp {{ background-color: {bg}; color: {fg}; direction: rtl; text-align: right; }}
div[data-testid="stSidebar"] {{ background-color: {card}; direction: rtl; text-align: right; }}
h1, h2, h3, h4, p, label, li, .stMarkdown, button {{ font-family: 'Cairo', 'Segoe UI', Tahoma, sans-serif !important; }}
.stTabs [data-baseweb="tab-list"] {{ direction: rtl; }}
.stDataFrame, .stCode, pre, code {{ direction: ltr; text-align: left; }}
.stButton button, .stDownloadButton button, .stFormSubmitButton button {{ min-height: 3rem; font-size: 1rem; font-weight: 600; }}
div[data-testid="stMetricValue"] {{ color: #38bdf8 !important; }}
.owner {{ text-align:center; opacity:.7; font-size:.85rem; margin-top:1rem; }}
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    menu = st.radio("اختر الوحدة:", [
        "📊 لوحة القياس والبحث",
        "🗂️ قاموس البيانات وبطاقة الجدول",
        "➕ إضافة جدول جديد",
        "💻 مستودع الاستعلامات",
        "🧹 مغسلة الإكسيل",
        "⚖️ محرك المطابقة",
        "🔄 الاستيراد والتصدير",
    ])
    st.divider()
    st.caption(S.get("sync_status", ""))
    if st.button("☁️ مزامنة يدوية دلوقتي", use_container_width=True):
        save_all("تمت المزامنة")
    if st.button("🔁 إعادة فحص كل الاستعلامات", use_container_width=True):
        tot = resync_all_queries()
        save_all("تم إعادة الفحص")
        st.success(f"اتضاف {tot['tables']} جدول، {tot['columns']} عمود، {tot['relations']} علاقة")
    st.markdown('<div class="owner">© Ehab Ibrahim</div>', unsafe_allow_html=True)


# =====================================================================
#  📊 لوحة القياس والبحث
# =====================================================================
if menu == "📊 لوحة القياس والبحث":
    st.header("📊 لوحة القياس")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("الجداول", len(S.df_tables))
    c2.metric("الأعمدة", len(S.df_columns))
    c3.metric("العلاقات", len(S.df_rel))
    c4.metric("الاستعلامات", len(S.df_sql))
    c5.metric("قضايا الفروقات", len(S.df_logs))

    need = S.df_columns[S.df_columns["حالة_التأكد"].str.contains("يحتاج تأكيد|مكتشف", regex=True)]
    if len(need):
        st.warning(f"فيه {len(need)} عمود اتكتشف أوتوماتيك ومحتاج وصف أو تأكيد - هتلاقيهم في قاموس البيانات.")

    st.subheader("🔎 بحث في كل حاجة")
    q = st.text_input("اكتب اسم جدول أو عمود أو كلمة من شرح استعلام:", placeholder="مثلاً: ENTRYSERIAL أو مخازن")
    if q:
        ql = q.strip().upper()

        def hit(df, cols):
            m = pd.Series(False, index=df.index)
            for c in cols:
                m |= df[c].astype(str).str.upper().str.contains(re.escape(ql), regex=True)
            return df[m]

        ht = hit(S.df_tables, ["اسم_الجدول", "الوصف", "الموديول"])
        hc = hit(S.df_columns, ["اسم_الجدول", "اسم_العمود", "الوصف"])
        hq = hit(S.df_sql, ["عنوان_المشكلة", "شرح_الحل", "كود_SQL", "الجداول_المرتبطة"])
        st.write(f"لقيت: **{len(ht)}** جدول، **{len(hc)}** عمود، **{len(hq)}** استعلام")
        if len(ht):
            st.dataframe(ht, use_container_width=True, hide_index=True)
        if len(hc):
            st.dataframe(hc, use_container_width=True, hide_index=True)
        for _, r in hq.iterrows():
            with st.expander(f"💻 #{r['رقم_الاستعلام']} - {r['عنوان_المشكلة']}"):
                st.write(r["شرح_الحل"])
                st.code(r["كود_SQL"], language="sql")

    st.subheader("📌 سجل قضايا الفروقات")
    st.caption("تقدر تعدل في الجدول ده مباشرة أو تمسح صف، وبعدين دوس حفظ.")
    edited = st.data_editor(S.df_logs, use_container_width=True, num_rows="dynamic", key="logs_ed", hide_index=True)
    if st.button("💾 حفظ تعديلات السجل"):
        S.df_logs = _fix(edited, LOG_COLS)
        save_all("اتحفظ السجل")
        st.rerun()

    with st.expander("➕ تسجيل قضية فرق جديدة"):
        with st.form("log_form", clear_on_submit=True):
            ent = st.text_input("اسم المورد / الحساب وكوده:")
            val = st.number_input("قيمة الفرق:", format="%.2f")
            higher = st.selectbox("مين الأعلى؟", ["الموديول أعلى من الحسابات العامة", "الحسابات العامة أعلى من الموديول", "غير محدد"])
            stat = st.selectbox("الحالة:", ["مفتوحة - جاري الفحص", "تم اكتشاف السبب الجذري", "تم التعديل من الحسابات"])
            root = st.text_area("السبب الجذري أو أرقام القيود:")
            if st.form_submit_button("إضافة للسجل", use_container_width=True):
                S.df_logs = pd.concat([S.df_logs, pd.DataFrame([[ent, val, higher, stat, root, f"{datetime.now():%Y-%m-%d}"]],
                                                               columns=LOG_COLS)], ignore_index=True)
                save_all("اتضافت القضية")
                st.rerun()


# =====================================================================
#  🗂️ قاموس البيانات وبطاقة الجدول
# =====================================================================
elif menu == "🗂️ قاموس البيانات وبطاقة الجدول":
    st.header("🗂️ قاموس البيانات")
    t_card, t_tables, t_cols, t_rel = st.tabs(["🔎 بطاقة جدول", "📁 كل الجداول", "🔑 كل الأعمدة", "🔗 العلاقات"])

    with t_card:
        names = sorted(S.df_tables["اسم_الجدول"].unique())
        if not names:
            st.info("القاموس فاضي. ضيف جدول أو استعلام وهيتملي لوحده.")
        else:
            tbl = st.selectbox("اختار الجدول:", names)
            info = S.df_tables[S.df_tables["اسم_الجدول"] == tbl].iloc[0]
            a, b, c = st.columns(3)
            a.markdown(f"**الموديول:** {info['الموديول']}")
            b.markdown(f"**النوع:** {info['النوع']}")
            c.markdown(f"**الحالة:** {info['حالة_التأكد']}")
            st.markdown(f"**الوصف:** {info['الوصف'] or '— لسه مفيش وصف —'}")

            cols_df = S.df_columns[S.df_columns["اسم_الجدول"] == tbl]
            st.subheader(f"🔑 الأعمدة ({len(cols_df)})")
            st.dataframe(cols_df.drop(columns=["اسم_الجدول"]), use_container_width=True, hide_index=True)

            rel = S.df_rel[(S.df_rel["الجدول_الأول"] == tbl) | (S.df_rel["الجدول_الثاني"] == tbl)]
            st.subheader(f"🔗 بيترتبط بإيه ({len(rel)})")
            for _, r in rel.iterrows():
                st.code(f"{r['الجدول_الأول']}.{r['العمود_الأول']} = {r['الجدول_الثاني']}.{r['العمود_الثاني']}", language="sql")

            used = queries_using(tbl)
            st.subheader(f"💻 الاستعلامات اللي بتستخدمه ({len(used)})")
            for _, r in used.iterrows():
                with st.expander(f"#{r['رقم_الاستعلام']} - {r['عنوان_المشكلة']}"):
                    st.code(r["كود_SQL"], language="sql")

            st.subheader("⚡ جمل جاهزة")
            col_list = ",\n    ".join(cols_df["اسم_العمود"].tolist()) or "*"
            st.markdown("**SELECT بكل الأعمدة المعروفة:**")
            st.code(f"SELECT\n    {col_list}\nFROM {tbl}\nWHERE ROWNUM <= 100;", language="sql")
            st.markdown("**هات كل أعمدة الجدول الحقيقية من أوراكل** (شغلها وانسخ النتيجة في صفحة إضافة جدول):")
            st.code(f"SELECT COLUMN_NAME, DATA_TYPE\nFROM ALL_TAB_COLUMNS\nWHERE TABLE_NAME = '{tbl}'\nORDER BY COLUMN_ID;", language="sql")

            with st.expander("✏️ تعديل بيانات الجدول ده"):
                with st.form("edit_tbl"):
                    mod = st.selectbox("الموديول:", MODULES, index=MODULES.index(info["الموديول"]) if info["الموديول"] in MODULES else 0)
                    typ = st.selectbox("النوع:", ["Master", "Details", "Lookup", "غير محدد"],
                                       index=["Master", "Details", "Lookup", "غير محدد"].index(info["النوع"]) if info["النوع"] in ["Master", "Details", "Lookup", "غير محدد"] else 3)
                    stt = st.text_input("حالة التأكد:", value=info["حالة_التأكد"])
                    dsc = st.text_area("الوصف:", value=info["الوصف"])
                    if st.form_submit_button("💾 حفظ", use_container_width=True):
                        i = S.df_tables[S.df_tables["اسم_الجدول"] == tbl].index[0]
                        S.df_tables.loc[i, ["الموديول", "النوع", "حالة_التأكد", "الوصف"]] = [mod, typ, stt, dsc]
                        save_all("اتعدل الجدول")
                        st.rerun()

    with t_tables:
        st.caption("عدل في أي خانة مباشرة، أو امسح صف، وبعدين دوس حفظ.")
        ed = st.data_editor(S.df_tables, use_container_width=True, num_rows="dynamic", key="tbl_ed", hide_index=True)
        if st.button("💾 حفظ الجداول"):
            S.df_tables = _fix(ed, TABLE_COLS)
            save_all("اتحفظت الجداول")
            st.rerun()

    with t_cols:
        only_new = st.checkbox("اعرض الأعمدة اللي محتاجة تأكيد بس")
        view = S.df_columns
        if only_new:
            view = view[view["حالة_التأكد"].str.contains("مكتشف|يحتاج", regex=True)]
        ed = st.data_editor(view, use_container_width=True, num_rows="dynamic", key="col_ed", hide_index=False)
        if st.button("💾 حفظ الأعمدة"):
            if only_new:
                rest = S.df_columns.drop(index=view.index)
                S.df_columns = _fix(pd.concat([rest, ed]), COL_COLS)
            else:
                S.df_columns = _fix(ed, COL_COLS)
            save_all("اتحفظت الأعمدة")
            st.rerun()

    with t_rel:
        st.caption("العلاقات دي بتتجمع لوحدها من شروط الربط (ON A.X = B.Y) في الاستعلامات.")
        ed = st.data_editor(S.df_rel, use_container_width=True, num_rows="dynamic", key="rel_ed", hide_index=True)
        if st.button("💾 حفظ العلاقات"):
            S.df_rel = _fix(ed, REL_COLS)
            save_all("اتحفظت العلاقات")
            st.rerun()


# =====================================================================
#  ➕ إضافة جدول جديد
# =====================================================================
elif menu == "➕ إضافة جدول جديد":
    st.header("➕ إضافة جدول جديد للقاموس")
    m1, m2, m3 = st.tabs(["📋 الصق تعريف الجدول (الأسهل)", "📤 ارفع إكسيل أعمدة من أوراكل", "✍️ إدخال يدوي"])

    with m1:
        st.info("انسخ نتيجة أمر **DESC اسم_الجدول** أو جملة **CREATE TABLE** من Toad أو SQL Developer والصقها هنا. "
                "التطبيق هيطلع الأعمدة وأنواعها لوحده.")
        with st.form("ddl_form"):
            tname = st.text_input("اسم الجدول (مطلوب لو لصقت ناتج DESC بس):").upper().strip()
            text = st.text_area("الصق هنا:", height=220,
                                placeholder="Name          Null?    Type\n------------- -------- ------------\nRECEIPTNO              NUMBER\nTRANSDATE              DATE")
            a, b = st.columns(2)
            mod = a.selectbox("الموديول:", MODULES)
            typ = b.selectbox("النوع:", ["Master", "Details", "Lookup", "غير محدد"])
            dsc = st.text_input("وصف الجدول (اختياري):")
            if st.form_submit_button("🚀 اقرا وسكّن", use_container_width=True):
                parsed = parse_table_definition(text, tname)
                if not parsed["tables"] or not any(parsed["tables"].values()):
                    st.error("مقدرتش أقرا أعمدة. اتأكد إنك كتبت اسم الجدول ولصقت ناتج DESC أو CREATE TABLE.")
                else:
                    rep = absorb_definition(parsed, mod, typ, "مؤكد من تعريف الجدول", dsc, "تعريف جدول")
                    save_all("اتضاف الجدول")
                    show_report(rep)

    with m2:
        st.info("شغّل الجملة دي في أوراكل (غير اسم الجداول)، وصدّر النتيجة إكسيل وارفعها. ينفع تجيب جداول كتير مرة واحدة.")
        st.code("""SELECT C.TABLE_NAME, C.COLUMN_NAME, C.DATA_TYPE, M.COMMENTS
FROM ALL_TAB_COLUMNS C
LEFT JOIN ALL_COL_COMMENTS M
       ON M.OWNER = C.OWNER AND M.TABLE_NAME = C.TABLE_NAME AND M.COLUMN_NAME = C.COLUMN_NAME
WHERE C.TABLE_NAME IN ('MMPURCHASINGORDER', 'TAX')
ORDER BY C.TABLE_NAME, C.COLUMN_ID;""", language="sql")
        up = st.file_uploader("ارفع الملف:", type=["xlsx", "csv"], key="bulk_cols")
        if up:
            df = pd.read_excel(up) if up.name.endswith("xlsx") else pd.read_csv(up)
            df.columns = [str(c).upper().strip() for c in df.columns]
            st.dataframe(df.head(15), use_container_width=True)
            if not {"TABLE_NAME", "COLUMN_NAME"} <= set(df.columns):
                st.error("الملف لازم يكون فيه عمود TABLE_NAME وعمود COLUMN_NAME.")
            elif st.button("🚀 سكّن كل اللي في الملف", use_container_width=True):
                rep = {"tables": [], "columns": [], "relations": []}
                for _, r in df.iterrows():
                    t, c = str(r["TABLE_NAME"]).upper().strip(), str(r["COLUMN_NAME"]).upper().strip()
                    if not t or t == "NAN" or not c or c == "NAN":
                        continue
                    if upsert_table(t, status="مؤكد من أوراكل"):
                        rep["tables"].append(t)
                    desc = "" if pd.isna(r.get("COMMENTS")) else str(r.get("COMMENTS", ""))
                    dtype = "" if pd.isna(r.get("DATA_TYPE")) else str(r.get("DATA_TYPE", ""))
                    if upsert_column(t, c, dtype=dtype, status="مؤكد من أوراكل", desc=desc):
                        rep["columns"].append(f"{t}.{c}")
                save_all("اتسكن الملف")
                show_report(rep)

    with m3:
        with st.form("add_table", clear_on_submit=True):
            t_name = st.text_input("اسم الجدول في أوراكل:").upper().strip()
            a, b = st.columns(2)
            t_mod = a.selectbox("الموديول:", MODULES)
            t_type = b.selectbox("نوع الجدول:", ["Master", "Details", "Lookup"])
            t_ver = st.selectbox("حالة التأكد:", ["مؤكد ومجرب 100%", "مؤكد من شجرة القاعدة", "مرجعي - يحتاج فحص الأعمدة"])
            t_desc = st.text_area("وصف الجدول:")
            t_cols = st.text_area("أسماء الأعمدة (اختياري - كل عمود في سطر أو بينهم فاصلة):")
            if st.form_submit_button("حفظ الجدول", use_container_width=True):
                if not t_name:
                    st.error("اكتب اسم الجدول.")
                else:
                    added = upsert_table(t_name, status=t_ver, module=t_mod, ttype=t_type, desc=t_desc)
                    n = 0
                    for c in re.split(r"[,\n]+", t_cols):
                        if c.strip():
                            n += upsert_column(t_name, c.strip(), status=t_ver)
                    save_all("اتضاف الجدول")
                    st.success(("اتضاف الجدول" if added else "الجدول موجود قبل كده") + f" + {n} عمود جديد")


# =====================================================================
#  💻 مستودع الاستعلامات
# =====================================================================
elif menu == "💻 مستودع الاستعلامات":
    st.header("💻 مستودع جمل SQL")
    t_list, t_add = st.tabs(["📚 الاستعلامات المحفوظة", "➕ إضافة / تعديل استعلام"])

    with t_list:
        a, b = st.columns([2, 1])
        sq = a.text_input("🔎 دور بكلمة:", placeholder="مورد، خزينة، ENTRYSERIAL ...")
        tf = b.selectbox("فلتر بجدول:", ["الكل"] + sorted(S.df_tables["اسم_الجدول"].unique()))
        view = S.df_sql.sort_values("رقم_الاستعلام")
        if sq:
            key = re.escape(sq.strip().upper())
            m = pd.Series(False, index=view.index)
            for c in ["عنوان_المشكلة", "شرح_الحل", "كود_SQL", "الجداول_المرتبطة"]:
                m |= view[c].astype(str).str.upper().str.contains(key, regex=True)
            view = view[m]
        if tf != "الكل":
            view = view[view.index.isin(queries_using(tf).index)]
        st.caption(f"{len(view)} استعلام")

        for idx, row in view.iterrows():
            qid = int(row["رقم_الاستعلام"])
            with st.expander(f"📌 #{qid} [{row['الاصدار']}] {row['عنوان_المشكلة']}"):
                st.markdown(f"**الشرح:** {row['شرح_الحل']}")
                st.markdown(f"**الجداول (اتعرفت لوحدها):** `{row['الجداول_المرتبطة']}`")
                st.caption(f"آخر تحديث: {row['تاريخ_التحديث']}")
                st.code(row["كود_SQL"], language="sql")
                hist = S.df_hist[pd.to_numeric(S.df_hist["رقم_الاستعلام"], errors="coerce") == qid]
                if len(hist):
                    with st.popover(f"🕘 الإصدارات القديمة ({len(hist)})"):
                        for _, h in hist.iterrows():
                            st.markdown(f"**{h['الاصدار']}** - اتحفظ {h['تاريخ_الحفظ']}")
                            st.code(h["كود_SQL"], language="sql")
                c1, c2, c3 = st.columns(3)
                if c1.button("✏️ تعديل", key=f"ed{qid}", use_container_width=True):
                    S.edit_id = qid
                    st.info("روح لتبويب «إضافة / تعديل استعلام» هتلاقيه جاهز.")
                sure = c3.checkbox("متأكد من المسح", key=f"sure{qid}")
                if c2.button("🗑️ مسح", key=f"del{qid}", use_container_width=True, disabled=not sure):
                    S.df_sql = S.df_sql.drop(index=idx).reset_index(drop=True)
                    save_all("اتمسح الاستعلام")
                    st.rerun()

    with t_add:
        edit_id = S.get("edit_id")
        cur = None
        if edit_id is not None:
            rows = S.df_sql[S.df_sql["رقم_الاستعلام"] == edit_id]
            cur = rows.iloc[0] if len(rows) else None
            if cur is not None:
                st.warning(f"✏️ بتعدل استعلام #{edit_id} - الإصدار القديم هيتحفظ في السجل لوحده.")
                if st.button("إلغاء التعديل وبدء استعلام جديد"):
                    S.edit_id = None
                    st.rerun()

        with st.form(f"sql_form_{edit_id}", clear_on_submit=cur is None):
            s_title = st.text_input("عنوان / وظيفة الاستعلام:", value=cur["عنوان_المشكلة"] if cur is not None else "")
            s_desc = st.text_area("شرح بسيط بيعمل إيه:", value=cur["شرح_الحل"] if cur is not None else "")
            s_code = st.text_area("كود SQL:", height=260, value=cur["كود_SQL"] if cur is not None else "")
            s_ver = st.text_input("رقم الإصدار:", value=next_version(cur["الاصدار"]) if cur is not None else "v1.0")
            st.caption("💡 الجداول والأعمدة والعلاقات هتتقرا من الكود لوحدها وتتسكن في القاموس.")
            if st.form_submit_button("💾 حفظ وتسكين أوتوماتيك", use_container_width=True):
                if not s_code.strip() or not s_title.strip():
                    st.error("لازم عنوان وكود.")
                else:
                    today = f"{datetime.now():%Y-%m-%d}"
                    if cur is not None:
                        qid = int(edit_id)
                        S.df_hist = pd.concat([S.df_hist, pd.DataFrame([[qid, cur["الاصدار"], cur["عنوان_المشكلة"], cur["كود_SQL"], today]],
                                                                       columns=HIST_COLS)], ignore_index=True)
                        S.df_sql = S.df_sql[S.df_sql["رقم_الاستعلام"] != qid]
                    else:
                        qid = int(S.df_sql["رقم_الاستعلام"].max()) + 1 if len(S.df_sql) else 1
                    rep = absorb_sql(s_code, f"#{qid}")
                    new = pd.DataFrame([[qid, s_title, ", ".join(rep["all_tables"]), s_ver, s_desc, s_code, today]], columns=SQL_COLS)
                    S.df_sql = pd.concat([S.df_sql, new], ignore_index=True).sort_values("رقم_الاستعلام").reset_index(drop=True)
                    S.edit_id = None
                    save_all(f"اتحفظ الاستعلام #{qid}")
                    st.success(f"اتحفظ استعلام #{qid} ✅")
                    show_report(rep)

        with st.expander("👀 جرّب القارئ من غير ما تحفظ"):
            test = st.text_area("الصق أي جملة:", height=150, key="try_sql")
            if test.strip():
                p = parse_sql(test)
                st.write("**الجداول:**", ", ".join(f"{t} ({'/'.join(a) or '-'})" for t, a in p["tables"].items()) or "مفيش")
                st.write("**الأعمدة:**", ", ".join(f"{t}.{c}" for t, c, _ in p["columns"]) or "مفيش")
                st.write("**العلاقات:**", ", ".join(f"{a}.{b} = {c}.{d}" for a, b, c, d in p["relations"]) or "مفيش")


# =====================================================================
#  🧹 مغسلة الإكسيل
# =====================================================================
def clean_dataframe(df, skip_rows=0, drop_empty=True, ffill_cols=None, calc_mode="بدون", col_a=None, col_b=None):
    d = df.copy()
    if 0 < skip_rows < len(d):
        d.columns = [str(c) for c in d.iloc[skip_rows - 1]]
        d = d.iloc[skip_rows:].reset_index(drop=True)
    if drop_empty:
        d = d.dropna(how="all", axis=0).dropna(how="all", axis=1).reset_index(drop=True)
    for c in ffill_cols or []:
        if c in d.columns:
            d[c] = d[c].ffill()
    if calc_mode != "بدون" and col_a and col_b:
        a = pd.to_numeric(d[col_a], errors="coerce").fillna(0)
        b = pd.to_numeric(d[col_b], errors="coerce").fillna(0)
        if calc_mode.startswith("ضرب"):
            d["الإجمالي_المحسوب"] = (a * b).round(2)
        elif calc_mode.startswith("طرح"):
            d["الصافي_المحسوب"] = (a - b).round(2)
        else:
            d["العمود_المدمج"] = d[col_a].astype(str) + " - " + d[col_b].astype(str)
    return d


def read_any(f):
    return pd.read_excel(f) if f.name.lower().endswith("xlsx") else pd.read_csv(f)


def to_xlsx(sheets: dict):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        for name, df in sheets.items():
            df.to_excel(w, sheet_name=name[:31], index=False)
    return out.getvalue()


if menu == "🧹 مغسلة الإكسيل":
    st.header("🧹 تنظيف شيتات الإكسيل الخارجة من التقارير")
    st.info("لو الملف فيه لوجو المستشفى فوق، أو خلايا مدمجة، أو أعمدة فاضية، أو محتاج تضرب الكمية × السعر.")
    raw = st.file_uploader("ارفع الشيت الخام:", type=["xlsx", "csv"])
    if raw:
        df_raw = read_any(raw)
        st.subheader("1️⃣ قبل التنظيف")
        st.dataframe(df_raw.head(10), use_container_width=True)
        c1, c2, c3 = st.columns(3)
        with c1:
            skip_r = st.number_input("صفوف فوق تتمسح:", 0, 50, 0)
            drop_emp = st.checkbox("امسح الصفوف والأعمدة الفاضية", value=True)
        temp = clean_dataframe(df_raw, skip_r, drop_emp)
        with c2:
            ffill_c = st.multiselect("أعمدة مدمجة تتملي لتحت:", list(temp.columns))
        with c3:
            calc_m = st.selectbox("عملية حسابية:", ["بدون", "ضرب عمودين (الكمية × السعر)", "طرح عمودين (مدين - دائن)", "دمج نصي لعمودين"])
            ca = cb = None
            if calc_m != "بدون":
                ca = st.selectbox("العمود الأول:", list(temp.columns))
                cb = st.selectbox("العمود التاني:", list(temp.columns))
        clean = clean_dataframe(df_raw, skip_r, drop_emp, ffill_c, calc_m, ca, cb)
        fcol = st.selectbox("امسح صفوف الإجماليات والنصوص (اختار عمود لازم يكون رقم):", ["بدون"] + list(clean.columns))
        if fcol != "بدون":
            clean = clean[pd.to_numeric(clean[fcol], errors="coerce").notna()].reset_index(drop=True)
        st.subheader(f"2️⃣ بعد التنظيف ({len(clean)} صف)")
        st.dataframe(clean, use_container_width=True)
        st.download_button("📥 نزّل الشيت النضيف", to_xlsx({"Cleaned_Data": clean}),
                           file_name=f"Cleaned_{raw.name.rsplit('.', 1)[0]}.xlsx", use_container_width=True)


# =====================================================================
#  ⚖️ محرك المطابقة  (المفتاح بقى: رقم القيد + المورد)
# =====================================================================
def norm_key(series):
    num = pd.to_numeric(series, errors="coerce")
    out = num.map(lambda v: "" if pd.isna(v) else (str(int(v)) if float(v).is_integer() else str(v)))
    return out.where(out != "0", "")


if menu == "⚖️ محرك المطابقة":
    st.header("⚖️ مطابقة الموديول مع القيود")
    st.caption("المطابقة بقت على (رقم القيد + المورد) عشان لو القيد فيه أكتر من مورد ميتلخبطوش.")
    c1, c2 = st.columns(2)
    mod_file = c1.file_uploader("1️⃣ شيت الموديول (إيصالات / مخازن)", type=["xlsx", "csv"], key="mod")
    gl_file = c2.file_uploader("2️⃣ شيت القيود (GL)", type=["xlsx", "csv"], key="gl")

    if mod_file and gl_file:
        mod_raw, gl_raw = read_any(mod_file), read_any(gl_file)
        with st.expander("🧹 تنظيف سريع (لو فيه صفوف لوجو فوق أو محتاج كمية × سعر)"):
            a, b = st.columns(2)
            df_mod = clean_dataframe(mod_raw, a.number_input("صفوف فوق في الموديول:", 0, 30, 0, key="s_mod"))
            if a.checkbox("احسب المبلغ = الكمية × السعر"):
                qc = a.selectbox("عمود الكمية:", list(df_mod.columns), key="q_m")
                pc = a.selectbox("عمود السعر:", list(df_mod.columns), key="p_m")
                df_mod = clean_dataframe(df_mod, 0, False, None, "ضرب", qc, pc)
            df_gl = clean_dataframe(gl_raw, b.number_input("صفوف فوق في القيود:", 0, 30, 0, key="s_gl"))

        mc, gc = list(df_mod.columns), list(df_gl.columns)
        NONE = "— بدون —"

        def guess(cols, *names):
            for n in names:
                for i, c in enumerate(cols):
                    if n in str(c).upper():
                        return i
            return 0

        st.subheader("⚙️ اختار الأعمدة")
        a, b, c = st.columns(3)
        with a:
            st.markdown("**الموديول**")
            mod_key = st.selectbox("رقم القيد:", mc, index=guess(mc, "ENTRYSERIAL", "قيد"))
            mod_val = st.selectbox("المبلغ:", mc, index=mc.index("الإجمالي_المحسوب") if "الإجمالي_المحسوب" in mc else guess(mc, "TOTAL", "مبلغ", "AMOUNT"))
            mod_sup = st.selectbox("المورد (اختياري):", [NONE] + mc, index=1 + guess(mc, "SUPPLIER") if any("SUPPLIER" in str(x).upper() for x in mc) else 0)
        with b:
            st.markdown("**القيود (GL)**")
            gl_key = st.selectbox("رقم القيد:", gc, index=guess(gc, "ENTRYSERIAL", "قيد"))
            gl_val = st.selectbox("المبلغ (DEBIT أو CREDIT):", gc, index=guess(gc, "CREDIT", "DEBIT"))
            gl_sup = st.selectbox("المورد / رقم الحساب (اختياري):", [NONE] + gc,
                                  index=1 + guess(gc, "ACCOUNTCODE", "SUPPLIER") if any(k in str(x).upper() for x in gc for k in ("ACCOUNTCODE", "SUPPLIER")) else 0)
            from_acc = st.checkbox("ده رقم حساب: خد كود المورد من بعد أول 6 أرقام (2020060504 ← 504)", value=True)
        with c:
            st.markdown("**إعدادات**")
            tolerance = st.number_input("هامش كسور التقريب:", value=0.01, step=0.01)
            use_year = st.checkbox("فلتر بسنة معينة")
            year = st.number_input("السنة:", 2000, 2100, 2025) if use_year else None
            mod_date = st.selectbox("عمود التاريخ في الموديول:", mc, index=guess(mc, "DATE", "تاريخ")) if use_year else None
            gl_date = st.selectbox("عمود التاريخ/السنة في القيود:", gc, index=guess(gc, "YEARCODE", "DATE", "تاريخ")) if use_year else None

        if st.button("🚀 شغّل المطابقة", use_container_width=True):
            M, G = df_mod.copy(), df_gl.copy()

            def year_of(s):
                n = pd.to_numeric(s, errors="coerce")
                if n.between(1900, 2100).mean() > 0.8:
                    return n
                return pd.to_datetime(s, errors="coerce", dayfirst=True).dt.year

            if use_year:
                M = M[year_of(M[mod_date]) == year]
                G = G[year_of(G[gl_date]) == year]

            M["_E"], G["_E"] = norm_key(M[mod_key]), norm_key(G[gl_key])
            M["_V"] = pd.to_numeric(M[mod_val], errors="coerce").fillna(0)
            G["_V"] = pd.to_numeric(G[gl_val], errors="coerce").fillna(0)
            M["_S"] = norm_key(M[mod_sup]) if mod_sup != NONE else ""
            if gl_sup != NONE:
                raw_s = G[gl_sup].astype(str).str.strip().str.replace(r"\.0$", "", regex=True)
                G["_S"] = norm_key(raw_s.str[6:]) if from_acc else norm_key(raw_s)
            else:
                G["_S"] = ""
            if mod_sup == NONE or gl_sup == NONE:
                M["_S"], G["_S"] = "", ""

            unposted = M[(M["_E"] == "") & (M["_V"] != 0)].drop(columns=["_E", "_V", "_S"])
            posted = M[M["_E"] != ""]
            ma = posted.groupby(["_E", "_S"]).agg(إجمالي_الموديول=("_V", "sum"), عدد_أسطر_الموديول=("_V", "count")).reset_index()
            ga = G[G["_E"] != ""].groupby(["_E", "_S"]).agg(إجمالي_القيود=("_V", "sum"), عدد_أسطر_القيود=("_V", "count")).reset_index()
            mg = ma.merge(ga, on=["_E", "_S"], how="outer")
            for col in ["إجمالي_الموديول", "عدد_أسطر_الموديول", "إجمالي_القيود", "عدد_أسطر_القيود"]:
                mg[col] = mg[col].fillna(0)
            mg = mg.rename(columns={"_E": "رقم_القيد", "_S": "المورد"})
            mg["إجمالي_الموديول"] = mg["إجمالي_الموديول"].round(2)
            mg["إجمالي_القيود"] = mg["إجمالي_القيود"].round(2)
            mg["الفرق (القيود - الموديول)"] = (mg["إجمالي_القيود"] - mg["إجمالي_الموديول"]).round(2)

            def diag(r):
                if abs(r["الفرق (القيود - الموديول)"]) <= tolerance:
                    return "✅ متطابق"
                if r["عدد_أسطر_الموديول"] == 0:
                    return "👻 قيد بالحسابات من غير موديول"
                if r["عدد_أسطر_القيود"] == 0:
                    return "🛑 حركة موديول رقم قيدها مش موجود في القيود"
                return "⚠️ فرق في القيمة (استقطاع أو تعديل)"

            mg["التشخيص"] = mg.apply(diag, axis=1)
            mg["_n"] = pd.to_numeric(mg["رقم_القيد"], errors="coerce")
            mg = mg.sort_values(["_n", "المورد"]).drop(columns="_n")
            if mod_sup == NONE or gl_sup == NONE:
                mg = mg.drop(columns="المورد")

            tm = round(float(M["_V"].sum()), 2)
            tg = round(float(G["_V"].sum()), 2)
            net = round(tg - tm, 2)
            S.recon = {
                "all": mg, "unposted": unposted,
                "mis": mg[mg["التشخيص"].str.startswith("⚠️")],
                "orph": mg[mg["التشخيص"].str.startswith("👻")],
                "miss": mg[mg["التشخيص"].str.startswith("🛑")],
                "tm": tm, "tg": tg, "net": net,
                "who": "متطابقين" if abs(net) <= tolerance else ("الحسابات العامة أعلى" if net > 0 else "الموديول أعلى"),
                "name": f"{mod_file.name} × {gl_file.name}",
            }

    R = S.get("recon")
    if R:
        st.divider()
        st.subheader("📌 الخلاصة")
        a, b, c, d = st.columns(4)
        a.metric("إجمالي الموديول", f"{R['tm']:,.2f}")
        b.metric("إجمالي القيود", f"{R['tg']:,.2f}")
        c.metric("صافي الفرق", f"{abs(R['net']):,.2f}")
        d.metric("مين الأعلى؟", R["who"])
        tabs = st.tabs([f"⚠️ فروق مبالغ ({len(R['mis'])})", f"👻 قيود من غير موديول ({len(R['orph'])})",
                        f"🛑 قيد مش موجود ({len(R['miss'])})", f"⏳ موديول مش مترحل ({len(R['unposted'])})",
                        f"📋 الكل ({len(R['all'])})"])
        for t, k in zip(tabs, ["mis", "orph", "miss", "unposted", "all"]):
            t.dataframe(R[k], use_container_width=True, hide_index=True)

        summary = pd.DataFrame([
            ["إجمالي الموديول (مترحل + مش مترحل)", R["tm"]], ["إجمالي القيود", R["tg"]],
            ["صافي الفرق", abs(R["net"])], ["مين الأعلى", R["who"]],
            ["قيود فيها فرق مبالغ", len(R["mis"])], ["قيود من غير موديول", len(R["orph"])],
            ["حركات رقم قيدها مش موجود", len(R["miss"])], ["حركات مش مترحلة", len(R["unposted"])],
        ], columns=["البيان", "القيمة"])
        x, y = st.columns(2)
        x.download_button("📥 نزّل تقرير المطابقة (Excel)", to_xlsx({
            "1_الملخص": summary, "2_فروق_المبالغ": R["mis"], "3_قيود_بدون_موديول": R["orph"],
            "4_قيد_غير_موجود": R["miss"], "5_غير_مرحل": R["unposted"], "6_الكل": R["all"]}),
            file_name=f"HIS_Reconciliation_{datetime.now():%Y%m%d_%H%M}.xlsx", use_container_width=True)
        if y.button("📝 سجّل النتيجة دي في سجل القضايا", use_container_width=True):
            higher = ("الحسابات العامة أعلى من الموديول" if R["net"] > 0 else "الموديول أعلى من الحسابات العامة") if R["who"] != "متطابقين" else "غير محدد"
            note = f"مطابقة {R['name']}: {len(R['mis'])} فرق مبالغ، {len(R['orph'])} قيد بدون موديول، {len(R['unposted'])} غير مرحل"
            S.df_logs = pd.concat([S.df_logs, pd.DataFrame([[R["name"], abs(R["net"]), higher, "مفتوحة - جاري الفحص", note,
                                                             f"{datetime.now():%Y-%m-%d}"]], columns=LOG_COLS)], ignore_index=True)
            save_all("اتسجلت القضية")


# =====================================================================
#  🔄 الاستيراد والتصدير
# =====================================================================
if menu == "🔄 الاستيراد والتصدير":
    st.header("🔄 نسخة احتياطية من كل البيانات")
    st.download_button("📥 نزّل قاعدة المعرفة كاملة (Excel)", db_bytes(),
                       file_name=f"HIS_Audit_KnowledgeBase_{datetime.now():%Y%m%d}.xlsx", use_container_width=True)
    st.divider()
    up = st.file_uploader("📤 ارفع ملف قاعدة معرفة (ينفع ملفك القديم كمان):", type=["xlsx"])
    if up:
        mode = st.radio("تعمل إيه بيه؟", ["دمج مع الموجود (مفيش حاجة تتمسح)", "استبدال كامل"])
        if st.button("تأكيد", use_container_width=True):
            new = _from_excel_bytes(up.getvalue())
            if mode.startswith("استبدال"):
                for k, v in new.items():
                    S[k] = v
            else:
                for _, r in new["df_tables"].iterrows():
                    upsert_table(r["اسم_الجدول"], r["حالة_التأكد"], r["الموديول"], r["النوع"], r["الوصف"])
                for _, r in new["df_columns"].iterrows():
                    upsert_column(r["اسم_الجدول"], r["اسم_العمود"], r["نوع_البيانات"], r["نوع_المفتاح"], r["حالة_التأكد"], r["الوصف"])
                for _, r in new["df_rel"].iterrows():
                    upsert_relation(r["الجدول_الأول"], r["العمود_الأول"], r["الجدول_الثاني"], r["العمود_الثاني"], r["المصدر"])
                have = set(S.df_sql["كود_SQL"].str.strip())
                nid = int(S.df_sql["رقم_الاستعلام"].max()) + 1 if len(S.df_sql) else 1
                for _, r in new["df_sql"].iterrows():
                    if r["كود_SQL"].strip() and r["كود_SQL"].strip() not in have:
                        r = r.copy()
                        r["رقم_الاستعلام"] = nid
                        nid += 1
                        S.df_sql = pd.concat([S.df_sql, r.to_frame().T], ignore_index=True)
                S.df_logs = pd.concat([S.df_logs, new["df_logs"]], ignore_index=True).drop_duplicates()
            resync_all_queries()
            save_all("اتستورد الملف")
            st.success("تم ✅")
            st.rerun()
