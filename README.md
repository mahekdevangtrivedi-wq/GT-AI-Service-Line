# GT Bahrain: AI Service Line Deliverables (October 2026)

These are the deliverables for the AI service lines Grant Thornton Bahrain plans to set up. They cover AI Governance, Risk & Compliance and AI Strategy & Roadmap.

They were built from these inputs:

- the *AI Service Lines for GT Bahrain v3.0* deck
- the original *RCM Frameworks* workbook
- the *AI Ready7* questionnaire and GT Bahrain results report (17 Sept 2026)
- the NSW *aiaf-assessment-tool-01092026* workbook

## 1. AI Governance, Risk & Compliance

| # | Task | Deliverable | What it contains |
|---|---|---|---|
| 1.1 | RCM update | `01_AI_Governance_Risk_Compliance/1.1_RCM_Update/GT_AI_RCM_v2.1_2026.xlsx` | **81 controls in 13 domains** (was 44 in 12). There are **37 new controls** (9 of them added in v2.1 from the ISACA agent-security paper, the ISO 42001 checklist and the AI Security Audit Checklist v1.0), including a new domain, *Protection from AI-Enabled Threats*. All 44 original controls are enhanced with new columns: ISO/IEC 27001:2022 remapping, 2026 framework references, Bahrain/GCC regulatory references, AI Ready7 mapping, applicability (provider/deployer), control type, frequency, and fields for engagement testing. Other sheets: Gap Analysis, AI Ready7 coverage, a 35-entry Regulatory Library and a Change Log. The original sheet is kept unchanged for traceability. |
| | AI security audit | `1.1_RCM_Update/GT_AI_Security_and_Agent_Audit_Checklist_v2.0.xlsx` | **125-item audit programme**: the 85-item AI Security Audit Checklist v1.0 (numbering kept), plus a 23-item AI agent section (ISACA 2026) and a 17-item ISO/IEC 42001 section. Every item maps to RCM v2.1. Scoring by section, the v1.0 rating bands, an auto-filled findings register and sign-off. |
| 1.2 | AI risk assessment framework | `1.2_AI_Risk_Assessment_Framework/GT_AI_Risk_Assessment_Framework_v1.1.docx` | An 18-page methodology based on ISO/IEC 23894, 42001 and 42005, the NIST AI RMF and AI 600-1, the EU AI Act, the NSW AIAF, and Bahrain's PDPL and iGA AI policy. It covers a 7-stage process, 10-factor risk tiering (Low to Critical) with override rules, a **52-risk taxonomy in 12 categories** (v1.1 adds agent memory, sandbox/SSRF, prompt-change and kill-switch risks), 5x5 scoring across 6 impact dimensions, control-effectiveness-based residual risk, minimum requirements by tier, a RACI, a worked example, a secure-by-default baseline for AI agents (Appendix D) and data classification for AI (Appendix E). |
| | | `GT_AI_Risk_Assessment_Toolkit_Template.xlsx` / `..._Worked_Example.xlsx` | Excel toolkit: use-case profile, prohibited-practice screen, automatic tiering, a risk register pre-loaded with the taxonomy and linked to RCM controls, a heat map, a top-10 residual risk list and sign-off. The worked example is a Bahrain retail bank's agentic GenAI assistant. |
| 1.3 | Readiness assessment tool (AIAF-style) v2.0 | `1.3_AI_Readiness_Assessment_Tool/GT_AI_Readiness_Assessment_Tool.xlsx` | **GT-internal engine. Never sent to clients.** It has **20 questions**. Part A has 4 AIAF-based inherent-risk questions. Part B has 16 readiness questions in **5 equally weighted domains**: Strategy & Value; Governance, Risk & Compliance; People & Skills; Data & Technology; and AI Security & Threat Risk (which merges Protect AI and AI Threats). v2.0 adds **5 governance and risk questions**: governance structure, AI risk management framework, risk appetite & KRIs, third-party AI risk, and model risk management. The **3. Report** sheet builds the report automatically. |
| | | `client_questionnaire/GT_AI_Readiness_Questionnaire.docx` | The **client-facing questionnaire** (34 questions: 14 profile, 4 AI use, 16 readiness). It has no scores or weightings and is formatted for **Microsoft Forms Quick import**. The question tags (e.g., `[Q7]`) link answers back to the engine. |
| | | `client_questionnaire/GT_AI_Readiness_Responses_Template.xlsx` | A fallback response sheet with dropdowns, for when Forms cannot be used. |
| | | `src/generate_readiness_reports.py` | **Report generator (GT internal).** It takes the Forms Excel export (or the template), fills the engine, recalculates it with LibreOffice and writes, per organisation, the **PDF report**, the filled engine (internal) and an **Outlook email draft (.eml)** with the PDF attached. |
| | | `GT_INTERNAL_Readiness_Assessment_Operating_Guide.docx` | A step-by-step guide covering Forms set-up, response export, generating the report, quality review, sending, data handling and troubleshooting. |
| | | `GT_AI_Readiness_Assessment_Proposal.pptx` | A **16-slide client proposal** setting out the approach and methodology. It covers why to assess now, what the assessment answers, the 5 steps, the 5 domains, the inherent risk profile, how results are expressed, collection and confidentiality, sample report pages, GT Bahrain's own position (Ready7 vs v2.0 across 5 domains), options and timeline, and deliverables. |
| | | `..._SAMPLE_GT_Bahrain.xlsx` + `SAMPLE_Report_GT_Bahrain.pdf` + `sample_output/` | The GT Bahrain sample: 60%, Informed, inherent tier Medium, exposure Moderate. Also included are the 7-page report, the sample Forms responses and the email draft produced by the generator. |

### The readiness report (sheet *3. Report*)
The report covers:

- an executive summary with headline figures and the recommended pathway
- an auto-generated narrative
- domain scores with a radar chart
- the inherent-risk profile
- coverage of the 8 NSW AIAF principles
- detailed findings (response, potential risk, recommended action, linked RCM controls)
- the top 5 priorities and a phased roadmap
- regulatory considerations
- a mapping to GT services, and a sign-off block

Page breaks are computed from row heights, so sections are not split. Clients never open the workbook, because GT sends them the PDF.

**Printing / PDF:** every generated workbook is post-processed with `src/cache_values.py`, which stores each formula's calculated value in the file. Values therefore display and print even in Protected View. All workbooks use plain A4 page setup with defined print areas.

## 2. AI Strategy & Roadmap

| Deliverable | What it contains |
|---|---|
| `02_AI_Strategy_and_Roadmap/GT_AI_Strategy_and_Roadmap_Proposal.pptx` | A **22-slide proposal covering approach and methodology only**, on the GT template. **Context & starting point:** the questions leadership asks; GT Bahrain's **AI Ready7 position consolidated into the 5 domains vs end-2027 targets**; and where AI fits in GT Bahrain, from the toolkit worked example. **Approach & methodology:** 5 phases, the principles, then the Opportunity Discovery method step by step (position archetypes, task-to-solution families and verdict logic, value/feasibility/risk scoring, outputs), strategy formulation, roadmap & business case, and the standards basis. **Engagement plan:** an 8-week timeline, deliverables, three options (Automation & AI Scan for SMEs, Standard, Strategy + Readiness), and why GT, including the GT tools and the **AI-powered BCP and Financial Statements tools (in preparation)**. |
| `02_AI_Strategy_and_Roadmap/GT_AI_Opportunity_Discovery_Toolkit.xlsx` | **Finds where exactly AI fits, and where automation is enough. It is not a readiness test.** It has the following sheets. **1. AI Position:** 12 questions on Data Position and Scale & Capacity place the organisation in one of 4 archetypes (*Automation is enough - for now*, *Data foundations first*, *Buy and configure AI*, *Build and scale AI*), with a ceiling on solution complexity. **2. Process Library:** 63 cross-industry and Bahrain-sector processes. **3. AI Fit Diagnostic:** for up to 50 processes, the task nature maps to 1 of 10 solution families (levels 1-5). The sheet then applies a data gate and an organisational-fit check, and gives a verdict (AI fit now / Automate - AI not needed / AI later - close data gap / Stretch - partner / Redesign first), value, feasibility, AIAF-style risk tier, quadrant, rank, horizon and hours released. **4. Opportunity Map:** results by function, verdict mix and the top 10. **5. Roadmap:** H1/H2/H3 lanes and the enabling foundations. **6. Use-Case Canvas:** one page per opportunity with minimum controls by tier. **Reference:** standards basis covering OECD AI classification, ISO/IEC 42001, 22989, 23053, 5259, 8183, 42005, 23894 and 5338, NIST AI RMF, NSW AIAF, EU AI Act, PDPL, ISACA 2026, OWASP LLM and APQC PCF. |
| `..._Worked_Example_GT_Bahrain.xlsx` | GT Bahrain applied to itself across 21 processes. The result is *Buy and configure AI* (Data 56% / Scale 44%). 52% of processes are AI fit now, 24% need automation only, and 24% need a partner, data or redesign first. |

## Points to validate before client use
- **Regulatory status** is correct as of September 2026. Three items need confirming against primary sources:
  - the EU AI Act high-risk dates, which depend on the Digital Omnibus outcome
  - the status of the Bahrain draft AI law, which is not yet in force
  - exact module references in the CBB Rulebook for AI/ML
- **Targets and estimates** (Ready7 end-2027 targets; hours released in the Opportunity Discovery Toolkit, which use typical reduction rates per solution family) are indicative and must be validated per engagement.
- **Framework citation in the service-lines deck (slide 8):** it cites "ISO/IEC TR 24028 for bias assessment". The bias technical report is **ISO/IEC TR 24027**; TR 24028 covers trustworthiness.
- **Original RCM ISO/IEC 27001 references** mixed the 2013 and 2022 Annex A numbering. Column I of RCM v2.0 gives a curated 2022 remapping.

## Regenerating the files
The `src/` folder holds the Python scripts that generate every output (openpyxl, python-docx, python-pptx, matplotlib). The content is kept in separate files so it can be maintained in one place: `rcm_data.py`, `risk_data.py`, `tool_data.py`, `opportunity_data.py` and `strategy_data.py`. The decks are built by `build_strategy_proposal.py` and `build_readiness_proposal.py` (shared helpers in `deck_common.py`) from the GT AI Service Lines template, plus page images rendered from the toolkit worked example and the sample report. `src/cache_values.py` must be run on every generated workbook as the final step. `src/recalc_check.py` recalculates a workbook with LibreOffice and reports any formula errors. All the workbooks were checked this way and have **0 formula errors**.
