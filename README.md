# GT Bahrain: AI Service Line Deliverables (September 2026)

These are the deliverables for the AI service lines Grant Thornton Bahrain plans to set up. They cover AI Governance, Risk & Compliance and AI Strategy & Roadmap.

They were built from these inputs:

- the *AI Service Lines for GT Bahrain v3.0* deck
- the original *RCM Frameworks* workbook
- the *AI Ready7* questionnaire and GT Bahrain results report (17 Sept 2026)
- the NSW *aiaf-assessment-tool-01092026* workbook

## 1. AI Governance, Risk & Compliance

| # | Task | Deliverable | What it contains |
|---|---|---|---|
| 1.1 | RCM update | `01_AI_Governance_Risk_Compliance/1.1_RCM_Update/GT_AI_RCM_v2.0_2026.xlsx` | **72 controls in 13 domains** (was 44 in 12). There are **28 new controls**, including a new domain, *Protection from AI-Enabled Threats*. All 44 original controls are enhanced with new columns: ISO/IEC 27001:2022 remapping, 2026 framework references, Bahrain/GCC regulatory references, AI Ready7 mapping, applicability (provider/deployer), control type, frequency, and fields for engagement testing. Other sheets: Gap Analysis, AI Ready7 coverage, a 32-entry Regulatory Library and a Change Log. The original sheet is kept unchanged for traceability. |
| 1.2 | AI risk assessment framework | `1.2_AI_Risk_Assessment_Framework/GT_AI_Risk_Assessment_Framework_v1.0.docx` | A 17-page methodology based on ISO/IEC 23894, 42001 and 42005, the NIST AI RMF and AI 600-1, the EU AI Act, the NSW AIAF, and Bahrain's PDPL and iGA AI policy. It covers a 7-stage process, 10-factor risk tiering (Low to Critical) with override rules, a **48-risk taxonomy in 12 categories**, 5x5 scoring across 6 impact dimensions, control-effectiveness-based residual risk, minimum requirements by tier, a RACI and a worked example. |
| | | `GT_AI_Risk_Assessment_Toolkit_Template.xlsx` / `..._Worked_Example.xlsx` | Excel toolkit: use-case profile, prohibited-practice screen, automatic tiering, a risk register pre-loaded with the taxonomy and linked to RCM controls, a heat map, a top-10 residual risk list and sign-off. The worked example is a Bahrain retail bank's agentic GenAI assistant. |
| 1.3 | Readiness assessment tool (AIAF-style) | `1.3_AI_Readiness_Assessment_Tool/GT_AI_Readiness_Assessment_Tool.xlsx` | **15 questions.** Part A has 4 AIAF-based inherent-risk questions. Part B has 11 readiness questions covering the 7 AI Ready7 pillars, each mapped to RCM v2.0 controls. The **3. Report** sheet builds an in-depth report automatically (see the next section). There is also an Action Plan tracker and a Methodology & mapping sheet. |
| | | `GT_AI_Readiness_Assessment_Tool_SAMPLE_GT_Bahrain.xlsx` + `SAMPLE_Report_GT_Bahrain.pdf` | A copy pre-filled so that it reproduces GT Bahrain's AI Ready7 profile (62% Informed vs. the actual 61.4%), and the 5-page report it produces. |

### The readiness report (sheet *3. Report*)
The report is ready to print on A4, with page breaks between sections. It covers:

- an executive summary with headline figures: readiness %, maturity level, inherent risk tier, exposure and recommended pathway
- an auto-generated narrative
- pillar scores with a radar chart against the 70% target
- the inherent-risk profile
- coverage of the 8 NSW AIAF ethics principles
- detailed findings in the AI Ready7 layout: response, potential risk, recommended action and linked controls
- the top 5 priorities and a phased roadmap
- regulatory applicability (PDPL, iGA/GCC, CBB, EU AI Act, GCC regimes, ISO/IEC 42001, the draft Bahrain AI law)
- a mapping to GT services, and a sign-off block

A panel beside the report, outside the print area, explains how to print or save it as a PDF. It also has a **pre-filled `mailto:` link** that sends the headline results to the email address entered in the Profile sheet. Excel cannot attach files without macros, so the user saves the PDF and attaches it themselves. The workbook has no macros, so corporate email filters won't block it.

## 2. AI Strategy & Roadmap

| Deliverable | What it contains |
|---|---|
| `02_AI_Strategy_and_Roadmap/GT_Bahrain_AI_Strategy_and_Roadmap_2026-2028.pptx` | A **28-slide deck on the GT Bahrain template**, with GT Bahrain as "client zero". **Situational / SWOT analysis:** the Bahrain AI landscape, competitive positioning, the AI Ready7 baseline, a SWOT and a TOWS. **Strategy formulation:** vision, 3 objectives with targets, 5 pillars and principles, AI risk appetite, prioritisation of 12 use cases, service-line go-to-market, operating model and KPIs. **Roadmap:** 3 horizons, an 18-month Gantt with 5 workstreams and 21 activities, a 100-day plan, resourcing, risks and the decisions required. The appendix traces every AI Ready7 gap to a roadmap action. All charts and tables are native and editable. |
| `02_AI_Strategy_and_Roadmap/GT_AI_Strategy_Toolkit.xlsx` | A reusable workbook for client engagements: SWOT/TOWS, weighted use-case scoring (quadrant and rank calculate automatically), an auto-drawing Gantt, a KPI tracker and AI Ready7 traceability. |

## Points to validate before client use
- **Regulatory status** is correct as of September 2026. Three items need confirming against primary sources:
  - the EU AI Act high-risk dates, which depend on the Digital Omnibus outcome
  - the status of the Bahrain draft AI law, which is not yet in force
  - exact module references in the CBB Rulebook for AI/ML
- **Strategy targets** (engagement numbers, time savings, readiness targets) are proposals for leadership to validate. KPI baselines marked "to be baselined" should be measured in the first 100 days.
- **Framework citation in the service-lines deck (slide 8):** it cites "ISO/IEC TR 24028 for bias assessment". The bias technical report is **ISO/IEC TR 24027**; TR 24028 covers trustworthiness.
- **Original RCM ISO/IEC 27001 references** mixed the 2013 and 2022 Annex A numbering. Column I of RCM v2.0 gives a curated 2022 remapping.

## Regenerating the files
The `src/` folder holds the Python scripts that generate every output (openpyxl, python-docx, python-pptx, matplotlib). The content is kept in separate files so it can be maintained in one place: `rcm_data.py`, `risk_data.py`, `tool_data.py` and `strategy_data.py`. `src/recalc_check.py` recalculates a workbook with LibreOffice and reports any formula errors. All the workbooks were checked this way and have **0 formula errors**.
