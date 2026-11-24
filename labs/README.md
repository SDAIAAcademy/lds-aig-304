# أنظمة إدارة الذكاء الاصطناعي والتدقيق الداخلي: المختبرات

اثنا عشر مختبرًا لمقرر أكاديمية سدايا **LDS-AIG-304**، وملفات المشروع الختامي ومهمة المتابعة. المختبرات صفحات إرشادية بالعربية وملفات جداول ونصوص: الهدف، والملفات، وخطوات بدقائقها، وأسئلة تحقق، وتمرين إضافي ✱، وملاحظات للمدرب. لا كود يكتبه المشارك.

كل المختبرات في شركة واحدة اصطناعية: **شركة مَطايا للخدمات اللوجستية** ومنتجاها «عَيّوق» و«غَفوة»، وثلاث خدمات تشتريها من موردين بلا أسماء. كل الأسماء والسجلات اصطناعية، ومولّدة ببذرة ثابتة. كتالوج الضوابط وقاعدة البيانات العامة لأدوات الموظفين وقاعدة درجات الملاحظات قواعد مطايا أو المقرر، كُتبت للمقرر.

## ما يحتاجه المشارك

- متصفح، وبرنامج جداول يفتح ملفات `xlsx` و`csv` ويكتب من اليمين إلى اليسار، ومحرر نصوص.
- مساعد ذكاء اصطناعي تعتمده جهته، اختياري، مرشدًا في المختبرات 1.1 إلى 2.4 و3.1 و3.4. لا يُنزَّل أي نموذج على الجهاز، ولا يدخل المساعد إلا ملفات المقرر، وهي مصنفة «عامة» لأغراضه.
- المختبران 3.2 و3.3 **دون أي ذكاء اصطناعي**.
- نسخة مرخّصة من ISO/IEC 42001 يوفرها المدرب في القاعة. نص المعيار لا ينسخه المقرر.

## نقطة البداية والحل

كل مختبر يبدأ من ملفات جاهزة في `files/` ولا يعتمد على مخرج مختبر قبله. المختبران 3.3 و3.4 يبنيان على عمل المشارك في المختبر السابق، ومن لم يُكمله يتسلم من المدرب نتائج بديلة (`solutions/fallback_findings.md`). في اليوم الثالث يدقق المشاركون سجلات مطايا في `files/evidence/` كما هي، لا ما كتبوه في اليومين الأولين.

| المختبر | الصفحة | يبدأ من | الحل المرجعي |
|---|---|---|---|
| 1.1 | `D1_1_clause_map.html` | `clause_map.xlsx`، `case_brief.txt` | `solutions/D1_1_clause_map.md` |
| 1.2 | `D1_2_roles.html` | `systems.xlsx` | `solutions/D1_2_roles.md` |
| 1.3 | `D1_3_scope.html` | `draft_scope.txt`، `obligations.txt` | `solutions/D1_3_scope.md` |
| 1.4 | `D1_4_policy.html` | `policy_draft.txt`، `policy_checklist.txt` | `solutions/D1_4_policy.md` |
| 2.1 | `D2_1_objectives.html` | `objectives_draft.xlsx`، `competence_matrix.xlsx` | `solutions/D2_1_objectives.md` |
| 2.2 | `D2_2_soa.html` | `soa_draft.xlsx`، `risk_register.xlsx` | `solutions/D2_2_soa.md` |
| 2.3 | `D2_3_suppliers.html` | `supplier_sheets.txt`، `agent_sheet.txt`، `staff_tools.xlsx` | `solutions/D2_3_suppliers.md` |
| 2.4 | `D2_4_crosswalk.html` | `crosswalk.xlsx` | `solutions/D2_4_crosswalk.md` |
| 3.1 | `D3_1_programme.html` | `audit_programme_draft.xlsx`، `audit_plan_template.txt` | `solutions/D3_1_programme.md` |
| 3.2 | `D3_2_interviews.html` | `evidence/`، `interview_notes.txt`، `findings.xlsx` | `solutions/D3_2_interviews.md`، `solutions/interview_cards.md` |
| 3.3 | `D3_3_grading.html` | `findings.xlsx` | `solutions/D3_3_grading.md`، `solutions/fallback_findings.md` |
| 3.4 | `D3_4_root_cause.html` | `corrective_actions.xlsx` | `solutions/D3_4_root_cause.md`، `solutions/fallback_findings.md` |
| المشروع الختامي | شرائح `slides/capstone.html` | الحزمة من المختبرات | `solutions/capstone_trainer.md`، `solutions/followup_key.md` |

بنود الاختبار القصير المراقب لا تُحفظ في هذا المستودع.

## إعادة توليد الملفات والتحقق من الحلول

يحتاج المدرب `uv` فقط. من مجلد المقرر:

```
uv run --no-project --with openpyxl labs/_tools/prepare_files.py
uv run --no-project --with openpyxl labs/solutions/check.py
uv run --no-project --with openpyxl labs/_tools/sampling.py
```

الأول يعيد كتابة `files/` كلها من `_tools/case_data.py` بالبذرة نفسها، فتخرج الملفات نفسها. والثاني يعيد حساب مرجع كل مختبر من الملفات ويقابله بالحلول، ويطبع «12 of 12 labs pass». والثالث يحسب أحجام العينات واحتمالاتها وأيام العمل (أسبوع العمل من الأحد إلى الخميس) التي تذكرها الشرائح والحلول.

## المصادر والترخيص

- نصوص الفئات الفرعية في `crosswalk.xlsx` و`_tools/nist_ai_rmf_subcategories.json` من NIST AI 100-1 (الإصدار 1.0)، عمل لحكومة الولايات المتحدة.
- أرقام الضوابط المرجعية في كتالوج مطايا من المقابلة التي تستضيفها NIST بين إطارها وISO/IEC 42001؛ عناوين الضوابط بكلمات المقرر. يتحقق المدرب من الأرقام في النسخة المرخّصة.
- كل ما سوى ذلك من ملفات كتبه المقرر.
