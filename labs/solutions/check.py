# /// script
# requires-python = ">=3.11"
# dependencies = ["openpyxl>=3.1"]
# ///
"""Recomputes the reference answer of every lab from the files in labs/files/ and prints PASS per lab.

Run from the course root:  uv run labs/solutions/check.py          (all labs)
                           uv run labs/solutions/check.py 2.2 3.2  (some)
It reads only; it writes nothing and calls no model. The reference answers are in the D*.md files
beside this script; this script proves that the files the participants receive support them.
"""
import csv
import re
import sys
from pathlib import Path

from openpyxl import load_workbook

LABS = Path(__file__).resolve().parents[1]
FILES = LABS / "files"
EVID = FILES / "evidence"
SOL = Path(__file__).resolve().parent
sys.path.insert(0, str(LABS / "_tools"))
import case_data as D  # noqa: E402
from sampling import p_at_least_one, working_days, zero_defect_n  # noqa: E402


def rows(name, sheet=None):
    wb = load_workbook(FILES / name, read_only=True)
    ws = wb[sheet] if sheet else wb.worksheets[0]
    data = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return data[0], data[1:]


def txt(name):
    return (FILES / name).read_text(encoding="utf-8")


def records(name):
    with open(EVID / "records" / name, encoding="utf-8", newline="") as f:
        return list(csv.DictReader(f))


def sol(name):
    return (SOL / name).read_text(encoding="utf-8")


def lab_1_1():
    _, rs = rows("clause_map.xlsx")
    assert [r[0] for r in rs] == [c[0] for c in D.CLAUSES]
    ref = {c: k for c, _, k in D.CLAUSES}
    r = [c for c, k in ref.items() if k == "R"]
    i = [c for c, k in ref.items() if k == "I"]
    print(f"  {len(rs)} clauses; risk: {r}; impact: {i}; neither: {len(rs) - len(r) - len(i)}")
    key = sol("D1_1_clause_map.md")
    return all(f"| {c} |" in key for c in ref) and r == ["6.1.2", "6.1.3", "8.2", "8.3"] and i == ["6.1.4", "8.4"]


def lab_1_2():
    _, rs = rows("systems.xlsx")
    assert [r[0] for r in rs] == [s[0] for s in D.SYSTEMS]
    inside = [s[0] for s in D.SYSTEMS if s[5] == "داخل النطاق"]
    roles = {s[0]: s[4] for s in D.SYSTEMS}
    print(f"  in scope: {inside}; roles: {roles}")
    key = sol("D1_2_roles.md")
    return inside == ["S1", "S2", "S3", "S4", "S5", "S6"] and all(roles[s] in key for s in roles)


def lab_1_3():
    draft = txt("draft_scope.txt")
    bad = [s for s, _, ok in D.DRAFT_SCOPE_EXCLUSIONS if not ok]
    good = [s for s, _, ok in D.DRAFT_SCOPE_EXCLUSIONS if ok]
    assert all(s in draft for s in bad + good) and "S6" not in draft
    obl = txt("obligations.txt")
    missing_obl = [o[0] for o in D.OBLIGATIONS if o[1] not in draft]
    print(f"  invalid exclusions: {bad}; valid: {good}; missing system: S6; obligations not in the draft: {missing_obl}")
    return bad == ["S2", "S5"] and good == ["S7", "S8"] and len(missing_obl) == 4 and "OB1" in obl


def lab_1_4():
    draft = txt("policy_draft.txt")
    lines = re.findall(r"^(P\d)\. ", draft, re.M)
    key = sol("D1_4_policy.md")
    verdicts = dict(re.findall(r"\| (P\d) \| ([^|]+) \|", key))
    print(f"  {len(lines)} draft lines; verdicts: {verdicts}")
    missing = re.findall(r"^- (C\d):", key, re.M)
    print(f"  checklist items the draft does not cover: {missing}")
    return lines == [f"P{i}" for i in range(1, 9)] and set(verdicts) == set(lines) and missing == ["C3", "C4", "C6", "C7", "C8"]


def lab_2_1():
    _, rs = rows("objectives_draft.xlsx")
    weak = [r[0] for r in rs if not (r[2] and r[3] and r[4] and r[5]) or r[3] == "0"]
    _, roles = rows("competence_matrix.xlsx", "مصفوفة الكفاءة")
    _, train = rows("competence_matrix.xlsx", "التدريب المتاح")
    covers = " ".join(t[2] for t in train)
    gap = "التحيز في الفرز الآلي" not in covers
    print(f"  objectives to fix: {weak}; roles: {len(roles)}; bias competence has no training: {gap}")
    key = sol("D2_1_objectives.md")
    fixed = re.findall(r"^\| (O\d) \|[^|]+\| يُعاد صوغه", key, re.M)
    return weak == ["O1", "O3", "O5", "O6"] and gap and fixed == weak


def lab_2_2():
    _, rs = rows("soa_draft.xlsx")
    assert len(rs) == 38
    blank = [r[0] for r in rs if not r[4] and r[5] != "يُستكمل في المختبر 2.3"]
    flagged = [r[0] for r in rs if r[8] and "راجِع" in r[8]]
    for cid in D.SOA_BLANK:
        assert D.soa_reason(cid), cid
    high = D.HIGH_RISK_CONTROLS
    # an exclusion is wrong when the register or a clause/obligation gives the control a reason
    wrong_excl = [c for c in flagged if D.SOA_FLAGGED[c][0] == "مستبعد" and D.soa_reason(c)]
    wrong_status = [c for c in flagged if D.SOA_FLAGGED[c][0] == "مُدرج" and D.SOA_FLAGGED[c][3] != D.status(c)]
    print(f"  blank rows: {blank}; flagged: {flagged}; wrong exclusions: "
          f"{[(c, D.soa_reason(c)) for c in wrong_excl]}; wrong status: {[(c, D.status(c)) for c in wrong_status]}")
    key = sol("D2_2_soa.md")
    in_key = all(f"| {c} |" in key for c in D.SOA_BLANK + flagged)
    print(f"  high-risk controls ({len(high)}): {high}")
    statuses = {c: D.status(c) for c in D.SOA_BLANK}
    print(f"  reference status of blank rows: {statuses}")
    return (blank == D.SOA_BLANK and flagged == ["MC-14", "MC-23", "MC-31"] and len(high) == 13
            and wrong_excl == ["MC-14", "MC-23"] and wrong_status == ["MC-31"] and in_key)


def lab_2_3():
    sheets = txt("supplier_sheets.txt")
    for sid, (fault, ctl, risk) in D.SUPPLIER_FAULTS.items():
        assert sid in sheets
    assert "نعم ما لم يطلب العميل الإيقاف" in sheets and "خارج المملكة" in sheets and "لم يقدم المورّد" in sheets
    agent = txt("agent_sheet.txt")
    assert "لا حد" in agent
    _, tools = rows("staff_tools.xlsx")
    breach = [t[0] for t in tools if t[4] != "عام"]
    print(f"  supplier faults: {[(s, f[0]) for s, f in D.SUPPLIER_FAULTS.items()]}; staff tools breaking the rule: {breach}")
    key = sol("D2_3_suppliers.md")
    return breach == ["T2", "T3", "T7"] and "T2 وT3 وT7" in key and all(f"| {c} |" in key for c in D.SOA_FOR_LAB_2_3)


def lab_2_4():
    _, xs = rows("crosswalk.xlsx", "المقابلة")
    _, gaps = rows("crosswalk.xlsx", "مرشحات الفجوة")
    _, subs = rows("crosswalk.xlsx", "الفئات الفرعية في NIST")
    assert [x[0] for x in xs] == D.XWALK_CONTROLS and len(subs) == 72
    names = {s[0] for s in subs}
    for c, (p, acc) in D.XWALK_REF.items():
        assert p in acc and all(a in names for a in acc), c
    ref_gaps = [g for g in D.GAP_CANDIDATES if g in D.GAPS]
    print(f"  12 controls, 72 subcategories; gaps among the candidates: {ref_gaps}")
    key = sol("D2_4_crosswalk.md")
    rows_ok = all(re.search(rf"^\| {c} \| [^|]+ \| {p} \|", key, re.M) for c, (p, _) in D.XWALK_REF.items())
    return len(ref_gaps) == 6 and len(gaps) == 8 and rows_ok


def lab_3_1():
    _, prog = rows("audit_programme_draft.xlsx")
    covered = {c.split(".")[0] for p in D.DRAFT_PROGRAMME for c in p[2]}
    uncovered = [g for g in D.PROGRAMME_CLAUSE_GROUPS if g not in covered]
    ctl = {c for p in D.DRAFT_PROGRAMME for c in p[3]}
    miss = [c for c in D.HIGH_RISK_CONTROLS if c not in ctl]
    aud = {a[0]: a for a in D.AUDITORS}
    conflicts = []
    for a, _, _, _, fn, who in D.DRAFT_PROGRAMME:
        for x in who:
            if aud[x][2] == D.DEPARTMENT_OF_FUNCTION[a]:
                conflicts.append((a, x))
    print(f"  {len(prog)} audits; uncovered clause groups: {uncovered}; uncovered high-risk controls: {miss}; conflicts: {conflicts}")
    key = sol("D3_1_programme.md")
    in_key = all(x in key for x in ("A6", "A5", "البند 7", "MC-10", "MC-37"))
    return uncovered == ["7"] and miss == ["MC-10", "MC-37"] and conflicts == [("AU-1", "A6"), ("AU-2", "A5")] and in_key


def lab_3_2():
    found = {}
    rel = records("ayyuq_releases.csv")
    found["NC1"] = [r["version"] for r in rel if r["signoff_date"] > r["deploy_date"]]
    ch = records("ghafwa_changes.csv")
    ia = records("impact_assessments.csv")
    last_ia = max(r["date"] for r in ia if r["system"].startswith("S2"))
    found["NC2"] = [r["change_id"] for r in ch if not r["impact_assessment_ref"] and r["date"] > last_ia]
    sa = records("supplier_assessments.csv")
    found["NC3"] = [r["system"] for r in sa if not r["assessment_ref"]]
    ia_systems = {r["system"].split()[0] for r in ia}
    found["NC3"] += [f"{s} بلا تقييم أثر" for s in ("S1", "S2", "S3", "S4", "S5", "S6") if s not in ia_systems]
    # nothing else in the repository breaks a clause of the five procedures
    others = []
    others += [r["system"] for r in sa if r["assessment_ref"] and r["assessment_date"] >= r["contract_date"]]
    others += [r["change_id"] for r in ch if r["change_id"] != "CH-G-08" and not r["impact_assessment_ref"]]
    tr = records("training_records.csv")
    trained = {r["person_id"] for r in tr if r["programme"] == "مراجعة تجاوز درجة غفوة"}
    ov = records("ghafwa_overrides.csv")
    reviewers = sorted({r["reviewer"] for r in ov})
    found["NC4"] = [s for s in reviewers if s not in trained]
    done = {r["person_id"]: r["completed"] for r in tr if r["programme"] == "مراجعة تجاوز درجة غفوة"}
    others += [s for s in reviewers if s in done and done[s] >= min(r["date"] for r in ov if r["reviewer"] == s)]
    val = {r["validation_report"]: r for r in records("ayyuq_validation_summary.csv")}
    others += [r["version"] for r in rel if float(val[r["validation_report"]]["within_2h_share"]) < float(val[r["validation_report"]]["threshold"])]
    by_untrained = sum(r["reviewer"] in found["NC4"] for r in ov)
    cs = {r["incident"]: r for r in records("cs_incidents.csv")}
    reg = {r["incident"]: r for r in records("ai_incident_register.csv")}
    late = []
    for k, r in cs.items():
        if r["ai_related"] == "نعم":
            assert k in reg, k
            if working_days(r["recorded"], reg[k]["received_by_office"]) > D.INCIDENT_LIMIT_WORKING_DAYS:
                late.append((k, working_days(r["recorded"], reg[k]["received_by_office"])))
    found["NC5"] = late
    # red herrings that must stay conforming
    fair = (EVID / "records" / "ghafwa_fairness_check_2026Q3.txt").read_text(encoding="utf-8")
    assert "ضمن العتبة" in fair
    for k, v in found.items():
        print(f"  {k}: {v}")
    print(f"  other breaches of the procedures: {others}")
    print(f"  overrides by untrained reviewers: {by_untrained} of {len(ov)}; random sample of 20 holds one with p = {p_at_least_one(len(ov), by_untrained, 20):.1%}")
    cards = sol("interview_cards.md")
    return (found["NC1"] == ["3.2"] and found["NC2"] == ["CH-G-08"] and found["NC3"] == ["S5", "S5 بلا تقييم أثر"] and others == []
            and found["NC4"] == ["SUP-04", "SUP-07"] and found["NC5"] == [("INC-0412", 13)]
            and by_untrained == D.OVERRIDES_BY_UNTRAINED and len(ov) == D.OVERRIDES_TOTAL
            and all(n in cards for n in ("عبيد", "فرحان", "فيّ")))


def lab_3_3():
    key = sol("D3_3_grading.md")
    grades = {k: v["grade"] for k, v in D.NONCONFORMITIES.items()}
    ok = all(f"| {k} |" in key and v["grade"] in key for k, v in D.NONCONFORMITIES.items())
    print(f"  grades: {grades}; opportunity: 1")
    return ok and list(grades.values()).count("جسيمة") == 2


def lab_3_4():
    _, rs = rows("corrective_actions.xlsx")
    days = [working_days(r, g) for _, r, g in D.FOLLOWUP_INCIDENTS]
    rate = sum(d <= D.INCIDENT_LIMIT_WORKING_DAYS for d in days) / len(days)
    print(f"  template rows: {len(rs)}; follow-up rate {rate:.1%} against target {D.FOLLOWUP_TARGET:.0%}: action not yet effective")
    print(f"  zero-defect samples at 95%: 10% -> {zero_defect_n(0.10)}, 5% -> {zero_defect_n(0.05)}")
    key = sol("D3_4_root_cause.md") + sol("followup_key.md")
    shown = f"{rate:.1%}"
    return len(rs) == 2 and rate < D.FOLLOWUP_TARGET and shown == "88.9%" and shown in key and "8 من 9" in key


LABS_ = {"1.1": lab_1_1, "1.2": lab_1_2, "1.3": lab_1_3, "1.4": lab_1_4,
         "2.1": lab_2_1, "2.2": lab_2_2, "2.3": lab_2_3, "2.4": lab_2_4,
         "3.1": lab_3_1, "3.2": lab_3_2, "3.3": lab_3_3, "3.4": lab_3_4}


def main():
    want = sys.argv[1:] or list(LABS_)
    bad = 0
    for k in want:
        print(f"Lab {k}")
        try:
            ok = LABS_[k]()
        except (AssertionError, FileNotFoundError) as e:
            ok = False
            print(f"  error: {e!r}")
        print(f"  {'PASS' if ok else 'FAIL'}")
        bad += not ok
    print(f"{len(want) - bad} of {len(want)} labs pass")
    sys.exit(1 if bad else 0)


if __name__ == "__main__":
    main()
