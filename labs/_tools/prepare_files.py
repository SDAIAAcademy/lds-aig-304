# /// script
# requires-python = ">=3.11"
# dependencies = ["openpyxl>=3.1"]
# ///
"""Writes the synthetic file pack of LDS-AIG-304 into labs/files/ (fixed seed, same bytes on every run).

Run from the course root:  uv run labs/_tools/prepare_files.py
Everything is fictional: Mataya Logistics, Ayyuq, Ghafwa, the suppliers and the people do not exist.
The files are synthetic and classified «عام» (public) for the course, so they may be given to any
assistant the lab allows.
"""
import csv
import datetime
import io
import random
import re
import sys
import zipfile
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

sys.path.insert(0, str(Path(__file__).resolve().parent))
import case_data as D  # noqa: E402

OUT = Path(__file__).resolve().parents[1] / "files"
EVID = OUT / "evidence"
FIXED = (2026, 11, 1, 0, 0, 0)
NOTE = "ملف اصطناعي لمقرر LDS-AIG-304. شركة مطايا وكل من فيها افتراضيون. التصنيف لأغراض المقرر: عام."


def save_xlsx(wb, path):
    """Save with fixed timestamps so the bytes do not change between runs."""
    wb.properties.created = datetime.datetime(2026, 11, 1)
    wb.properties.modified = datetime.datetime(2026, 11, 1)
    wb.properties.creator = "LDS-AIG-304"
    buf = io.BytesIO()
    wb.save(buf)
    src = zipfile.ZipFile(io.BytesIO(buf.getvalue()))
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as z:
        for info in sorted(src.infolist(), key=lambda i: i.filename):
            zi = zipfile.ZipInfo(info.filename, date_time=FIXED)
            zi.compress_type = zipfile.ZIP_DEFLATED
            data = src.read(info.filename)
            if info.filename == "docProps/core.xml":
                data = re.sub(rb"<dcterms:(created|modified)([^>]*)>[^<]*<", rb"<dcterms:\1\2>2026-11-01T00:00:00Z<", data)
            z.writestr(zi, data)
    path.write_bytes(out.getvalue())


def add_sheet(wb, title, header, rows, widths=None, first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = title
    ws.sheet_view.rightToLeft = True
    ws.append(header)
    for c in ws[1]:
        c.font = Font(bold=True)
        c.fill = PatternFill("solid", fgColor="DDEBF7")
    for r in rows:
        ws.append(list(r))
    for i, w in enumerate(widths or [18] * len(header)):
        ws.column_dimensions[chr(65 + i)].width = w
    for row in ws.iter_rows(min_row=2):
        for c in row:
            c.alignment = Alignment(wrap_text=True, vertical="top")
    return ws


def book(path, sheets):
    wb = Workbook()
    for k, (title, header, rows, widths) in enumerate(sheets):
        add_sheet(wb, title, header, rows, widths, first=(k == 0))
    save_xlsx(wb, OUT / path)


def text(path, body):
    p = OUT / path
    p.parent.mkdir(parents=True, exist_ok=True)
    p.write_text(body.strip() + "\n", encoding="utf-8", newline="\n")


def table_csv(path, header, rows):
    p = OUT / path
    p.parent.mkdir(parents=True, exist_ok=True)
    buf = io.StringIO()
    w = csv.writer(buf, lineterminator="\n")
    w.writerow(header)
    w.writerows(rows)
    p.write_text(buf.getvalue(), encoding="utf-8", newline="\n")


def name(k):
    return D.PEOPLE[k][0]


def days_ar(n):
    """Arabic count of days: 3 to 10 take the plural, 11 and more the singular accusative."""
    return f"{n} أيام" if 3 <= n <= 10 else f"{n} يومًا"


# --- Shared ------------------------------------------------------------------------------
def case_brief():
    people = "\n".join(f"- {n}: {r}" for n, r in D.PEOPLE.values())
    systems = "\n".join(f"- {s[0]} {s[1]}: {s[2]} ({s[3]})" for s in D.SYSTEMS[:6])
    text("case_brief.txt", f"""
{NOTE}

{D.COMPANY}
====================

شركة خاصة للخدمات اللوجستية: نقل بري للبضائع، وتخزين بعقود، وسلسلة تبريد، في مدن المملكة.
نحو {D.STAFF:,} موظف، و{D.TRUCKS} شاحنة، و{D.WAREHOUSES} مستودعات، و{D.DEPOTS} مراكز للأسطول.
تنقل مطايا شحنات لعملاء من القطاع الخاص، وتنقل شحنات لجهات حكومية بعقود، فتحفظ بيانات تلك الشحنات بصفتها شريكًا للجهة.

لماذا نظام إدارة للذكاء الاصطناعي الآن
---------------------------------------
- تبني مطايا منتجين وتشتري ثلاث خدمات، وأضاف الموظفون أدواتهم الخاصة دون سجل.
- طلب عميلان كبيران دليلًا على حوكمة «عَيّوق» قبل تجديد العقد.
- قررت الرئيسة التنفيذية أن تبني مطايا نظام إدارة للذكاء الاصطناعي على المعيار الدولي، وأن تدققه داخليًا قبل أن تقرر هل تطلب شهادة مطابقة.

الأنظمة والأدوات
----------------
{systems}

الأشخاص
-------
{people}

ملاحظة عن الأدوار في المقرر
- في اليومين الأول والثاني تعمل محللًا في مكتب حوكمة الذكاء الاصطناعي مع درة الرويلي.
- في اليوم الثالث تعمل مدققًا داخليًا في إدارة التدقيق الداخلي مع أريام الشمري، وتدقق سجلات مطايا التي كتبتها وظائفها، لا ما كتبته أنت في المختبرات.
""")


# --- Day 1 --------------------------------------------------------------------------------
def lab_1_1():
    book("clause_map.xlsx", [(
        "خريطة البنود", ["البند", "ملخص البند (بكلمات المقرر)", "يستدعي: تقييم المخاطر / تقييم الأثر / لا شيء", "الإرشاد الذي يخدمه", "ملاحظتك"],
        [(c, s, "", "", "") for c, s, _ in D.CLAUSES], [9, 52, 26, 34, 30])])


def lab_1_2():
    book("systems.xlsx", [(
        "الأنظمة والأدوات", ["المعرف", "الاسم", "ما يفعله", "كيف حصلت عليه مطايا", "دور مطايا (مطوّر، مقدم خدمة، مستخدم، لا شيء)", "نظام ذكاء اصطناعي؟", "مرشح للنطاق؟ ولماذا"],
        [(s[0], s[1], s[2], s[3], "", "", "") for s in D.SYSTEMS], [8, 26, 44, 26, 24, 16, 36])])


def lab_1_3():
    sys_name = {s[0]: s[1] for s in D.SYSTEMS}
    excl = "\n".join(f"- {sid} {sys_name[sid]}: {why}" for sid, why, _ in D.DRAFT_SCOPE_EXCLUSIONS)
    text("draft_scope.txt", f"""
{NOTE}

مسودة نطاق نظام إدارة الذكاء الاصطناعي في {D.COMPANY}
الإصدار 0.3 · أعدها مستشار خارجي · لم تُعتمد

1. الغرض
يشمل نظام الإدارة أنظمة الذكاء الاصطناعي التي تطورها مطايا.

2. الحدود
- الوظائف: علم البيانات، وعمليات الأسطول، وخدمة العملاء.
- المواقع: المقر الرئيس ومراكز الأسطول التسعة.
- الأنظمة داخل النطاق: S1 عَيّوق، S3 مساعد خدمة العملاء، S4 خدمة قراءة المستندات.

3. الاستثناءات
{excl}

4. أصحاب المصلحة
عملاء الشحن.

5. المتطلبات
نظام حماية البيانات الشخصية.
""")
    obl = "\n".join(f"- {o[0]} {o[1]}: {o[2]}" for o in D.OBLIGATIONS)
    text("obligations.txt", f"""
{NOTE}

الالتزامات والجهات المعنية بأنظمة مطايا (من سجل الالتزامات لدى الشؤون القانونية)
{obl}

الجهات المتأثرة بالأنظمة، كما سجلها مكتب الحوكمة
- عملاء الشحن ومستلمو شحناتهم (S1، S3)
- السائقون ومستعملو الطريق (S2)
- المتقدمون للوظائف (S5)
- موظفو مطايا الذين يبنون الأدوات أو يستعملونها (S6)
- الجهات الحكومية التي تنقل مطايا شحناتها (S4)
""")


def lab_1_4():
    text("policy_draft.txt", f"""
{NOTE}

مسودة سياسة الذكاء الاصطناعي في {D.COMPANY} · أعدها مستشار · لم تُعتمد

P1. تستعمل مطايا الذكاء الاصطناعي استعمالًا أخلاقيًا.
P2. تلتزم مطايا بأحدث التقنيات لتكون رائدة في السوق.
P3. يجري فريق علم البيانات اختبار دقة أسبوعيًا لـ«عَيّوق»، ويُرفض الإصدار إذا قلت الدقة عن 92%.
P4. تتبنى مطايا مبادئ أخلاقيات الذكاء الاصطناعي الصادرة عن سدايا في كل نظام تبنيه أو تشتريه أو يبنيه موظفوها.
P5. يُبلَّغ مكتب الحوكمة بأي حادثة عبر النموذج الإلكتروني F-21 خلال خمسة أيام عمل.
P6. لا يُتخذ قرار يمس سائقًا أو متقدمًا لوظيفة بناءً على درجة آلية دون مراجعة بشرية.
P7. يُحظر استعمال أي أداة ذكاء اصطناعي دون موافقة مسبقة من الرئيسة التنفيذية.
P8. تحدد مطايا كل عام أهدافًا للذكاء الاصطناعي في ضوء هذه السياسة.
""")
    text("policy_checklist.txt", f"""
{NOTE}

قائمة مراجعة السياسة (قاعدة المقرر، مواءمة مع بند السياسة 5.2 في النسخة المرخّصة)
C1. تناسب غرض مطايا وأدوارها: مطوّر ومقدم خدمة ومستخدم.
C2. تعطي إطارًا لوضع الأهداف.
C3. تلتزم بالوفاء بالمتطلبات المنطبقة (الأنظمة والضوابط الوطنية وعقود العملاء).
C4. تلتزم بالتحسين المستمر لنظام الإدارة.
C5. تذكر المبادئ التي تلتزم بها مطايا.
C6. تحيل إلى سياسات مطايا الأخرى: الخصوصية، وأمن المعلومات، والمشتريات، والموارد البشرية.
C7. لها مالك، وتعتمدها الإدارة العليا، ولها مدة مراجعة.
C8. تُبلَّغ للموظفين وتتاح لأصحاب المصلحة المعنيين.

ما لا يدخل السياسة: الأرقام والعتبات والنماذج والمهل، فمكانها الإجراءات.
قاعدة مطايا الخاصة (ليست نصًا وطنيًا): لا يدخل أدوات الذكاء الاصطناعي التوليدي إلا ما صُنّف «عامًا»، على غرار إرشادات سدايا للجهات الحكومية التي تمنع إدخال ما صُنّف «مقيدًا فأعلى».
""")


# --- Day 2 --------------------------------------------------------------------------------
def lab_2_1():
    book("objectives_draft.xlsx", [(
        "الأهداف", ["المعرف", "الهدف كما كُتب", "القياس", "المستهدف", "الموعد", "المالك", "ملاحظتك"],
        [("O1", "ضمان استعمال أخلاقي للذكاء الاصطناعي", "", "", "", "", ""),
         ("O2", "أن يقع وصول الشحنة الفعلي خلال ساعتين من تنبؤ «عَيّوق»", "نسبة الشحنات خلال ساعتين", "90%", "نهاية الربع الثاني 2027", name("obaid"), ""),
         ("O3", "تدريب الموظفين على الذكاء الاصطناعي", "", "", "قريبًا", "", ""),
         ("O4", "خفض تجاوزات «غَفوة» دون مراجعة", "نسبة التجاوزات دون مراجعة من كل التجاوزات شهريًا", "أقل من 2%", "مارس 2027", name("farhan"), ""),
         ("O5", "تقييم الموردين الثلاثة", "عدد الموردين المقيمين", "3", "", "", ""),
         ("O6", "صفر حوادث ذكاء اصطناعي", "عدد الحوادث", "0", "2027", name("durra"), "")],
        [8, 40, 34, 14, 22, 18, 30])])
    roles = ["مديرة نظام الإدارة", "عالم بيانات", "مشرف نوبات (يراجع تجاوزات غفوة)", "موظف خدمة عملاء (يصعّد أخطاء المساعد)",
             "أخصائي مشتريات (يقيّم موردي الذكاء الاصطناعي)", "مسؤول توظيف (يراجع فرز السير الذاتية)", "مدقق داخلي", "موظف يبني أداة بخدمة توليدية"]
    comps = ["أساسيات مخاطر الذكاء الاصطناعي", "تقييم الأثر", "التحقق من النماذج", "مراجعة تجاوز الدرجة",
             "تقييم الموردين", "التحيز في الفرز الآلي", "تصنيف البيانات قبل الأدوات التوليدية", "أساليب التدقيق"]
    book("competence_matrix.xlsx", [(
        "مصفوفة الكفاءة", ["الدور"] + comps,
        [(r,) + ("",) * len(comps) for r in roles], [34] + [16] * len(comps)),
        ("التدريب المتاح", ["البرنامج", "المدة", "ما يغطيه"],
         [("مدخل إلى مخاطر الذكاء الاصطناعي", "3 ساعات", "أساسيات المخاطر، وتصنيف البيانات قبل الأدوات التوليدية"),
          ("التحقق من نماذج التعلم الآلي", "يومان", "التحقق من النماذج"),
          ("مراجعة تجاوز درجة غفوة", "ساعتان", "متى يُقبل التجاوز وكيف يُسجل"),
          ("تقييم موردي الخدمات الرقمية", "يوم", "تقييم الموردين، دون جزء خاص بالذكاء الاصطناعي"),
          ("المدقق الداخلي لأنظمة الإدارة", "ثلاثة أيام", "أساليب التدقيق")],
         [34, 12, 60])])


def soa_rows():
    rows = []
    for cid, ref, grp, title in D.CONTROLS:
        ev = D.STATUS_EVIDENCE.get(cid, "")
        if cid in D.SOA_BLANK:
            rows.append((cid, ref, grp, title, "", "", "", "", ev))
        elif cid in D.SOA_FLAGGED:
            inc, why, own, st = D.SOA_FLAGGED[cid]
            rows.append((cid, ref, grp, title, inc, why, own, st, ev + " · راجِع هذا السطر"))
        elif cid in D.SOA_FOR_LAB_2_3:
            rows.append((cid, ref, grp, title, "", "يُستكمل في المختبر 2.3", "", "", ev))
        else:
            rows.append((cid, ref, grp, title, "مُدرج", D.soa_reason(cid), D.owner(cid), D.status(cid), ev))
    return rows


def lab_2_2():
    book("risk_register.xlsx", [(
        "سجل المخاطر", ["المعرف", "النظام", "الخطر (السبب، الحدث، النتيجة، من يتأثر)", "الفئة الوطنية", "الاحتمالية", "الأثر", "المستوى", "النطاق", "الضوابط المرتبطة"],
        [(r[0], r[1], r[2], r[3], r[4], r[5], D.level(r[4], r[5]), D.band(D.level(r[4], r[5])), "، ".join(r[6])) for r in D.RISKS],
        [8, 8, 62, 22, 10, 8, 9, 10, 26])])
    book("soa_draft.xlsx", [(
        "بيان قابلية التطبيق", ["المعرف", "المرجع في النسخة المرخّصة", "المجموعة", "الضابط (بكلمات مطايا)", "مُدرج / مستبعد", "المبرر", "المالك", "الحالة", "دليل الحالة"],
        soa_rows(), [8, 12, 18, 34, 12, 34, 24, 12, 40]),
        ("قواعد البيان", ["القاعدة"],
         [("كل ضابط مرجعي له سطر: مُدرج أو مستبعد، مع مبرر.",),
          ("مبرر الإدراج رقم خطر من السجل، أو بند، أو التزام. مبرر الاستبعاد يقول لماذا لا يوجد خطر يحتاجه.",),
          ("المالك دور في مطايا، لا اسم مورّد.",),
          ("الحالة: منفذ، أو منفذ جزئيًا، أو مخطط؛ تؤخذ من دليل الحالة، لا من الأمنيات.",)],
         [90])])


def lab_2_3():
    lines = []
    for sid, label, train, loc, bias, notice, inc, logs in D.SUPPLIERS:
        lines.append(f"""
{sid} · {label}
- استعمال بيانات مطايا في تدريب نموذج المورّد: {train}
- مكان المعالجة: {loc}
- أدلة اختبار التحيز: {bias}
- مهلة الإشعار بتغيير النموذج: {days_ar(notice) if notice else "لا التزام"}
- مهلة الإشعار بالحادثة: {inc}
- مدة حفظ السجلات لدى المورّد: {days_ar(logs)}""")
    text("supplier_sheets.txt", NOTE + "\n\nملخص شروط الموردين الثلاثة كما وردت في مسودات تجديد العقود هذا العام (المورّدون بلا أسماء في المقرر)\n" + "\n".join(lines))
    text("agent_sheet.txt", f"""
{NOTE}

أداة إعادة الجدولة في مساعد خدمة العملاء (S3)
- يستطيع المساعد أن يغيّر موعد التسليم حين يطلب المستلم ذلك في المحادثة.
- لا حد اليوم لعدد التسليمات التي يغيّرها المساعد في طلب واحد، ولا موافقة بشرية.
- يستطيع المساعد إعادة جدولة كل شحنات عميل الشحن حين يطلب ذلك حساب العميل.
- يحفظ المورّد سجل استدعاءات الأداة سبعة أيام.
- مصدر الطلب في السجل: نص المحادثة فقط، دون رقم الشحنة أو هوية الطالب.
""")
    book("staff_tools.xlsx", [(
        "سجل أدوات الموظفين", ["المعرف", "الفريق", "الغرض", "ما ترسله الأداة إلى الخدمة التوليدية", "تصنيف تلك البيانات", "ملاحظتك"],
        [t + ("",) for t in D.STAFF_TOOLS], [8, 16, 34, 44, 14, 30])])


def lab_2_4():
    sub = D.nist()
    book("crosswalk.xlsx", [(
        "المقابلة", ["ضابط مطايا", "المرجع", "الضابط", "الفئة الفرعية في إطار NIST", "لماذا تقابله (جملة)"],
        [(c, D.CONTROL[c][1], D.CONTROL[c][3], "", "") for c in D.XWALK_CONTROLS], [10, 10, 34, 22, 50]),
        ("مرشحات الفجوة", ["الفئة الفرعية", "نصها في إطار NIST", "هل في كتالوج مطايا ضابط يقابلها؟ (نعم / لا)", "أقرب ضابط إن وُجد"],
         [(g, sub[g], "", "") for g in D.GAP_CANDIDATES], [14, 80, 22, 20]),
        ("الفئات الفرعية في NIST", ["الفئة الفرعية", "النص"],
         sorted(sub.items(), key=lambda kv: (["GOVERN", "MAP", "MEASURE", "MANAGE"].index(kv[0].split()[0]), [int(x) for x in kv[0].split()[1].split(".")])),
         [14, 110])])


# --- Day 3 --------------------------------------------------------------------------------
def lab_3_1():
    aud = {a[0]: a for a in D.AUDITORS}
    book("audit_programme_draft.xlsx", [(
        "برنامج التدقيق", ["التدقيق", "الفترة", "البنود", "الضوابط", "الوظيفة المدققة", "المدققون"],
        [(a, q, "، ".join(cl), "، ".join(ct), fn, "، ".join(f"{x} {aud[x][1]}" for x in au)) for a, q, cl, ct, fn, au in D.DRAFT_PROGRAMME],
        [8, 12, 26, 50, 30, 40]),
        ("المدققون", ["المعرف", "الاسم أو الدور", "الإدارة الأصلية", "مدرب على التدقيق"],
         [(a[0], a[1], a[2], "نعم" if a[3] else "لا") for a in D.AUDITORS], [8, 26, 30, 14]),
        ("قواعد البرنامج", ["القاعدة"],
         [("يغطي البرنامج في الدورة الواحدة كل مجموعات البنود من 4 إلى 10.",),
          ("يغطي كل ضابط مرتبط بخطر مرتفع أو كارثي في السجل (المستوى 8 فأكثر).",),
          ("لا يدقق أحد عمله ولا عمل الإدارة التي يتبعها.",),
          ("يُدقق الأعلى خطرًا أولًا، ويُذكر سبب الترتيب.",)], [90])])
    text("audit_plan_template.txt", f"""
{NOTE}

خطة تدقيق واحدة (قالب)
- رقم التدقيق والوظيفة المدققة:
- الهدف:
- النطاق (البنود والضوابط والأنظمة والفترة):
- المعايير (الإجراءات والبنود التي يُقاس عليها):
- المدققون واستقلالهم:
- من تُجرى معه المقابلات:
- السجلات التي تُعاين وحجم العينة وطريقة اختيارها:
- المدة (بالساعات، دون أوقات):
- كيف يُبلَّغ عن النتائج:
""")


def evidence():
    rnd = random.Random(304)
    text("evidence/README.txt", f"""
{NOTE}

مستودع أدلة نظام إدارة الذكاء الاصطناعي في مطايا (الإصدار الأول، فترة التدقيق {D.AUDIT_PERIOD[0]} إلى {D.AUDIT_PERIOD[1]})
كتبته وظائف مطايا، لا المشاركون. سجل الإصدارات في _history.csv.

procedures/  الإجراءات المعتمدة
records/     السجلات التي تُعاين
""")
    text("evidence/procedures/P-03_release.txt", """
P-03 إجراء إصدار النماذج · الإصدار 2 · المالك: قائد علم البيانات · معتمد 2025-12-10
1. يسري على كل إصدار من «عَيّوق» و«غَفوة».
2. لكل إصدار تقرير تحقق يقيس الدقة على بيانات اختبار محفوظة.
3. يعتمد تقرير التحقق شخص لم يدرّب النموذج.
4. لا يُنشر إصدار قبل اعتماد تقرير التحقق.
5. يُسجل كل إصدار في سجل الإصدارات مع تاريخ النشر ومرجع التقرير وتاريخ اعتماده.
""")
    text("evidence/procedures/P-05_impact_assessment.txt", """
P-05 إجراء تقييم أثر أنظمة الذكاء الاصطناعي · الإصدار 1 · المالك: مكتب حوكمة الذكاء الاصطناعي · معتمد 2025-11-20
1. يُقيَّم أثر كل نظام قبل استعماله لأول مرة.
2. يُحفظ التقييم في سجل تقييمات الأثر بمرجع وتاريخ.
3. يُعاد التقييم قبل التشغيل عند أي من المحفزات: مصدر بيانات جديد، أو فئة متأثرة جديدة، أو استعمال جديد، أو تغيير جوهري في النموذج.
4. يذكر طلب التغيير مرجع التقييم الجديد أو سبب عدم لزومه.
""")
    text("evidence/procedures/P-07_suppliers.txt", """
P-07 إجراء موردي الذكاء الاصطناعي · الإصدار 1 · المالك: مديرة المشتريات · معتمد 2026-01-05
1. يسري على كل خدمة أو وحدة ذكاء اصطناعي تشتريها مطايا، ولو جاءت ضمن منصة أكبر.
2. يُقيَّم المورّد قبل توقيع العقد أو تجديده، ويُحفظ التقييم في سجل تقييمات الموردين.
3. يغطي التقييم استعمال البيانات، ومكان المعالجة، وأدلة الاختبار، والإشعار بالتغيير والحوادث.
""")
    text("evidence/procedures/P-08_competence.txt", """
P-08 إجراء كفاءة أدوار الذكاء الاصطناعي · الإصدار 1 · المالك: إدارة الموارد البشرية · معتمد 2026-01-12
1. مصفوفة الكفاءة تحدد التدريب اللازم لكل دور.
2. لا يتولى أحد مراجعة تجاوزات «غَفوة» قبل إتمام برنامج «مراجعة تجاوز درجة غفوة» وتسجيله.
3. تُحفظ سجلات التدريب في سجل التدريب.
""")
    text("evidence/procedures/P-09_incidents.txt", """
P-09 إجراء حوادث الذكاء الاصطناعي · الإصدار 1 · المالك: مكتب حوكمة الذكاء الاصطناعي · معتمد 2026-07-01
1. حادثة الذكاء الاصطناعي: خطأ من نظام ذكاء اصطناعي مس عميلًا أو موظفًا أو متقدمًا أو طرفًا آخر، أو كاد.
2. من يعلم بالحادثة يسجلها في نظام وظيفته.
3. تُبلَّغ الحادثة إلى مكتب حوكمة الذكاء الاصطناعي خلال خمسة أيام عمل من العلم بها.
4. يسجل المكتب الحادثة في سجل حوادث الذكاء الاصطناعي بتاريخ الاستلام.
""")
    # Records
    table_csv("evidence/records/ayyuq_releases.csv", ["version", "deploy_date", "validation_report", "signoff_date", "signed_by_role"],
              [r + ("مراجع مستقل من علم البيانات",) for r in D.AYYUQ_RELEASES])
    table_csv("evidence/records/ayyuq_validation_summary.csv", ["validation_report", "within_2h_share", "threshold", "result"],
              [("VAL-301", "0.91", "0.90", "ناجح"), ("VAL-302", "0.92", "0.90", "ناجح"), ("VAL-311", "0.90", "0.90", "ناجح"),
               ("VAL-312", "0.91", "0.90", "ناجح"), ("VAL-321", "0.93", "0.90", "ناجح"), ("VAL-322", "0.92", "0.90", "ناجح")])
    table_csv("evidence/records/ghafwa_changes.csv", ["change_id", "date", "description", "impact_assessment_ref"], D.GHAFWA_CHANGES)
    table_csv("evidence/records/impact_assessments.csv", ["ref", "date", "system", "scope"],
              [(r, d, "S2 غَفوة", s) for r, d, s in D.GHAFWA_IMPACT_ASSESSMENTS] +
              [("IA-A-01", "2025-11-28", "S1 عَيّوق", "الإصدار 3: بيانات الشحنات والمسارات"),
               ("IA-C-01", "2026-01-15", "S3 مساعد خدمة العملاء", "قبل التشغيل مع أداة إعادة الجدولة")])
    text("evidence/records/ghafwa_fairness_check_2026Q3.txt", """
فحص الإنصاف الربعي لـ«غَفوة» · الربع الثالث 2026 · أعده علم البيانات وراجعه مكتب الحوكمة
متوسط الدرجة لكل فئة عمرية مقارنة بمعدل الحوادث الفعلي لكل فئة.
الفرق الأكبر بين الفئات بعد ضبط ساعات القيادة: 3 نقاط، والعتبة في الإجراء 5 نقاط.
النتيجة: ضمن العتبة. الفحص لا يشمل مدخل كاميرا المقصورة لأنه أضيف بعد تحديد بيانات الفحص.
""")
    table_csv("evidence/records/supplier_assessments.csv", ["system", "supplier", "contract_date", "assessment_ref", "assessment_date", "outcome"], D.SUPPLIER_ASSESSMENTS)
    # Training records
    tr = []
    for s in D.SHIFT_SUPERVISORS:
        if s not in D.UNTRAINED_SUPERVISORS:
            tr.append((s, "مشرف نوبات", "مراجعة تجاوز درجة غفوة", f"2026-0{rnd.randint(1, 3)}-{rnd.randint(10, 27)}"))
        tr.append((s, "مشرف نوبات", "مدخل إلى مخاطر الذكاء الاصطناعي", f"2026-0{rnd.randint(1, 3)}-{rnd.randint(10, 27)}"))
    for i in range(1, 5):
        tr.append((f"DS-{i:02d}", "عالم بيانات", "التحقق من نماذج التعلم الآلي", f"2025-1{rnd.randint(0, 2)}-{rnd.randint(10, 27)}"))
    for i in range(1, 4):
        tr.append((f"PR-{i:02d}", "أخصائي مشتريات", "تقييم موردي الخدمات الرقمية", f"2026-02-{rnd.randint(10, 27)}"))
    table_csv("evidence/records/training_records.csv", ["person_id", "role", "programme", "completed"], tr)
    # Override log
    rows = []
    trained = [s for s in D.SHIFT_SUPERVISORS if s not in D.UNTRAINED_SUPERVISORS]
    flags = [True] * D.OVERRIDES_BY_UNTRAINED + [False] * (D.OVERRIDES_TOTAL - D.OVERRIDES_BY_UNTRAINED)
    rnd.shuffle(flags)
    start = datetime.date(2026, 3, 1)
    for k, bad in enumerate(flags, 1):
        d = start + datetime.timedelta(days=(k * 233) // D.OVERRIDES_TOTAL)
        sup = rnd.choice(D.UNTRAINED_SUPERVISORS) if bad else rnd.choice(trained)
        rows.append((f"OV-{k:03d}", d.isoformat(), f"DRV-{rnd.randint(100, 899)}", rnd.randint(66, 92), "سُمح بالنوبة بعد المراجعة", sup))
    table_csv("evidence/records/ghafwa_overrides.csv", ["override_id", "date", "driver_id", "score", "decision", "reviewer"], rows)
    # Incidents
    cs = [("INC-0318", "2026-07-08", "المساعد ذكر رسمًا جمركيًا قديمًا لمستلم", "نعم"),
          ("INC-0344", "2026-07-21", "تأخر شاحنة بسبب عطل", "لا"),
          ("INC-0367", "2026-08-04", "خطأ في استخراج وزن من بوليصة شحن", "نعم"),
          ("INC-0389", "2026-08-18", "شكوى من سلوك موظف توصيل", "لا"),
          (D.INCIDENT_LATE[0], D.INCIDENT_LATE[1], "المساعد أعاد جدولة 37 تسليمًا إلى مواعيد خاطئة لعميل شحن واحد", "نعم"),
          ("INC-0420", "2026-09-15", "تنبؤ وصول خاطئ لشحنة مبردة", "نعم"),
          ("INC-0429", "2026-10-06", "المساعد أعاد جدولة تسليم دون طلب واضح", "نعم")]
    table_csv("evidence/records/cs_incidents.csv", ["incident", "recorded", "description", "ai_related"], cs)
    reg = [("INC-0318", "2026-07-12"), ("INC-0367", "2026-08-06"), (D.INCIDENT_LATE[0], D.INCIDENT_LATE[2]),
           ("INC-0420", "2026-09-17"), ("INC-0429", "2026-10-08")]
    table_csv("evidence/records/ai_incident_register.csv", ["incident", "received_by_office"], reg)
    weeks = []
    for w in range(1, 19):
        weeks.append((f"2026-W{26 + w:02d}", f"{0.90 + (rnd.randint(0, 3)) / 100:.2f}", "0.90"))
    table_csv("evidence/records/ayyuq_weekly_monitoring.csv", ["week", "within_2h_share", "threshold"], weeks)
    hist = [("P-03_release", "2", "2025-12-10", "قائد علم البيانات", "إضافة اعتماد مستقل لتقرير التحقق"),
            ("P-05_impact_assessment", "1", "2025-11-20", "مكتب حوكمة الذكاء الاصطناعي", "الإصدار الأول"),
            ("P-07_suppliers", "1", "2026-01-05", "مديرة المشتريات", "الإصدار الأول"),
            ("P-08_competence", "1", "2026-01-12", "إدارة الموارد البشرية", "الإصدار الأول"),
            ("P-09_incidents", "0.4", "2026-03-02", "مكتب حوكمة الذكاء الاصطناعي", "مسودة"),
            ("P-09_incidents", "1", "2026-07-01", "مكتب حوكمة الذكاء الاصطناعي", "اعتماد بعد مراجعة خدمة العملاء"),
            ("ghafwa_changes", "9", "2026-07-19", "عمليات الأسطول", "إضافة CH-G-09"),
            ("supplier_assessments", "2", "2026-03-02", "المشتريات", "إضافة SA-2026-02")]
    table_csv("evidence/_history.csv", ["document", "version", "date", "changed_by", "change"], hist)
    text("interview_notes.txt", f"""
{NOTE}

ورقة ملاحظات المقابلة (تُكتب باليد أو في الملف؛ دون أي مساعد ذكي)
- من قابلت، ودوره، ومدة المقابلة بالدقائق:
- السؤال الذي سألته:
- ما قاله (بكلماته قدر الإمكان):
- السجل الذي طلبت رؤيته بعد إجابته:
- ما رأيته في السجل:
- هل يطابق القول السجل؟
""")


def lab_3_3():
    book("findings.xlsx", [(
        "النتائج", ["المعرف", "ما وجدته", "المتطلب (الإجراء والبند والضابط)", "الدليل (السجل ومكانه، والمقابلة)", "الدرجة: جسيمة / طفيفة / ملاحظة", "لماذا هذه الدرجة"],
        [(f"F{i}", "", "", "", "", "") for i in range(1, 8)], [6, 40, 34, 44, 18, 40]),
        ("قاعدة الدرجات", ["الدرجة", "متى (قاعدة المقرر)"], list(D.GRADING_RULE.items()), [12, 100])])


def lab_3_4():
    book("corrective_actions.xlsx", [(
        "الإجراءات التصحيحية", ["النتيجة", "التصحيح الفوري", "لماذا 1", "لماذا 2", "لماذا 3", "لماذا 4", "السبب الجذري", "الإجراء التصحيحي", "المالك", "الموعد", "مقياس النجاح", "كيف يُتحقق منه"],
        [("", "", "", "", "", "", "", "", "", "", "", ""), ("", "", "", "", "", "", "", "", "", "", "", "")],
        [12, 26, 22, 22, 22, 22, 30, 34, 16, 14, 30, 30])])


def main():
    OUT.mkdir(parents=True, exist_ok=True)
    case_brief()
    lab_1_1(); lab_1_2(); lab_1_3(); lab_1_4()
    lab_2_1(); lab_2_2(); lab_2_3(); lab_2_4()
    lab_3_1(); evidence(); lab_3_3(); lab_3_4()
    n = sum(1 for p in OUT.rglob("*") if p.is_file())
    print(f"wrote {n} files into labs/files")


if __name__ == "__main__":
    main()
