# -*- coding: utf-8 -*-
"""Content for the GT Bahrain AI Strategy & Roadmap (deck + toolkit)."""

READY7 = [("Business Alignment", 50.0), ("Governance & Ethics", 60.0), ("Training & Development", 68.75), ("Architecture & Technology", 55.0),
          ("Data & Quality", 70.0), ("Protect AI Systems", 55.0), ("Protection from AI Threats", 70.0)]
READY7_OVERALL = 61.4
READY7_TARGET_2027 = [75, 80, 80, 75, 80, 75, 80]

READY7_ITEMS = [  # item, title, result, roadmap response, horizon, workstream
    ("1.1", "Strategy, Culture & Risk Appetite", "Some Met", "Approve AI strategy, executive sponsor and AI risk appetite", "H1", "WS1"),
    ("1.2", "Portfolio Governance", "Some Met", "Use-case intake, prioritisation and shadow-AI conversion", "H1", "WS4"),
    ("1.3", "Value Realisation", "Some Met", "Benefit KPIs and quarterly portfolio value reviews", "H1-H2", "WS4"),
    ("1.4", "Client & Stakeholder Assurance", "Some Met", "Publish responsible-AI statement for clients; ISO/IEC 42001 readiness", "H2", "WS1"),
    ("2.1", "AI Policy & Acceptable Use", "Some Met", "AI acceptable-use policy and approved-tools catalogue", "H1", "WS1"),
    ("2.3", "Adoption Guardrails & Ethical Standards", "Some Met", "AI CoE charter and practical guardrails", "H1", "WS1"),
    ("2.4", "Risk Assessment & Governance Reviews", "Not Met", "Roll out GT AI Risk Assessment Framework to all AI use cases & vendors", "H1", "WS1"),
    ("3.1-3.3", "Literacy, Validation, Security awareness", "Most Met", "Mandatory literacy refresh; deepfake simulations", "H1-H2", "WS2"),
    ("3.4", "Role-Based Capability Development", "Some Met", "Role-based pathways, AI champions, workforce plan", "H2", "WS2"),
    ("4.1", "Enterprise AI Reference Architecture", "Some Met", "Reference architecture v1 (data, models, agents, security)", "H1-H2", "WS3"),
    ("4.2", "Integration & Execution Authority", "Some Met", "Agent execution-authority standard and trust boundaries", "H2", "WS3"),
    ("4.3", "API, Access & Trust Enforcement", "Some Met", "Centralised API gateway / access control for AI", "H2", "WS3"),
    ("4.4", "Observability & Performance", "Some Met", "Monitoring with KRIs, thresholds and mandatory responses", "H2", "WS3"),
    ("4.5", "Cost, Compute & Platform Governance", "Most Met", "Governed pathway from citizen-developed AI to production", "H2", "WS3"),
    ("5.1", "Data Classification, Access & Handling", "Not Met", "Classification rules for AI, DLP on GenAI, data-readiness gate", "H1", "WS3"),
    ("5.2 / 5.5", "Human Oversight; Privacy safeguards", "Most Met", "Human-in-charge standard for deliverables; prompt redaction", "H1", "WS1"),
    ("6.1", "Identity, Access & Privilege", "Most Met", "Non-human / agent identity and credential management", "H2", "WS3"),
    ("6.2", "Data & Model Protection", "Some Met", "Vendor AI contract clauses (no training, residency, audit)", "H1-H2", "WS1"),
    ("6.3", "Secure AI Development & ModelOps", "Some Met", "Secure AI SDLC and ModelOps pipeline", "H2", "WS3"),
    ("6.4", "AI Security Testing", "Some Met", "AI red teaming of internal assistants before scale", "H1-H2", "WS3"),
    ("6.5", "Runtime Protection & Logging", "Some Met", "Runtime guardrails and forensic-ready AI logging", "H2", "WS3"),
    ("7.3", "Identity, Behavioural & Deepfake Defence", "Some Met", "Deepfake-resilient verification for payments & sensitive requests", "H1-H2", "WS3"),
    ("7.1/7.2/7.4/7.5", "Threat intel, SOC, IR, Exposure mgmt", "Most Met", "Sustain; add AI-specific scenarios to SOC and red teaming", "H2-H3", "WS3"),
]

SWOT = dict(
    S=["Strong global brand and GTIL network; AI Ready7 methodology already in use",
       "Deep audit, GRC, cyber and regulatory heritage - a credible trust-first position",
       "Existing relationships in regulated sectors (banking, telecom, government)",
       "Compliance & data foundations in place (Ready7 2.2, 2.5, 5.3, 5.4 fully met)",
       "Good AI literacy (T&D 69%) and threat defence (70%) baseline",
       "Four AI service lines, proprietary tools (RCM v2.1, risk framework, readiness tool, security audit checklist) and two AI apps in preparation (BCP, Financial Statements)"],
    W=["Overall AI readiness 'Informed' (61.4%) - below the 'Repeatable' threshold",
       "No consistent lifecycle AI risk assessment (Ready7 2.4 - Not Met)",
       "Data classification & handling for AI not in place (Ready7 5.1 - Not Met)",
       "Business alignment 50%: strategy, portfolio and value tracking only partially defined",
       "Architecture & AI security at 55%: reference architecture, red teaming, ModelOps immature",
       "Limited in-house AI engineering capacity; reliance on vendor tools"],
    O=["National AI agenda: Vision 2030, iGA AI Policy (2025), UNESCO AI readiness report (2025)",
       "Regulation wave: PDPL, CBB, GCC regimes, EU AI Act reach, forthcoming Bahrain AI law",
       "Emerging ISO/IEC 42001 certification and AI assurance market in the GCC",
       "Sector demand in financial services, insurance, government, telecom, healthcare, hospitality",
       "AI productivity gains in audit, tax and advisory delivery improve margins and quality",
       "Regional scale through the GT network across the GCC"],
    T=["Big Four and tech-native AI consultancies competing aggressively (e.g., EY.ai, niche AI firms)",
       "Fast-moving technology and vendor lock-in risk",
       "AI-enabled cyber threats and deepfake fraud targeting the firm and its clients",
       "Client-confidentiality and professional-liability exposure from AI errors",
       "Scarcity of AI talent across the GCC",
       "Regulatory uncertainty (Bahrain AI law timing, EU Digital Omnibus changes)"],
)
TOWS = [("SO - Lead with trust", "Launch governance-led AI offerings (readiness assessment, AI GRC / RCM review, ISO/IEC 42001 readiness) into regulated sectors, leveraging the GT brand and GRC heritage."),
        ("WO - Become 'client zero'", "Close our own gaps (risk assessment, data classification, strategy & portfolio) using our own tools - and turn the journey into a reference case."),
        ("ST - Differentiate on assurance + security", "Combine AI GRC with cyber (AI red teaming, deepfake resilience) to stand apart from technology-led competitors."),
        ("WT - Partner and protect", "Vendor-agnostic alliances, strict client-data guardrails and targeted hiring / upskilling to contain talent, liability and lock-in risks.")]

PILLARS = [("1", "Trusted & Governed AI", "Responsible by design: policy, AI register, lifecycle risk assessment, CoE guardrails, ISO/IEC 42001-aligned AIMS.", "Governance & Ethics; Data & Quality"),
           ("2", "AI-Fluent People", "Every professional AI-literate; specialists, leaders and champions developed through role-based pathways.", "Training & Development"),
           ("3", "The Intelligent Firm", "Prioritised internal use cases in audit, tax, advisory and operations with measured value.", "Business Alignment"),
           ("4", "Secure AI Foundations", "Reference architecture, protected data, secure-by-default AI agents, AI security testing, monitoring and deepfake-resilient processes.", "Architecture; Protect AI; AI Threats"),
           ("5", "Market-Leading AI Advisory", "Four AI service lines packaged, launched and scaled into priority sectors across Bahrain and the GCC.", "Growth")]

PRINCIPLES = ["Human-in-charge: professionals remain accountable for every AI-assisted output",
              "Client confidentiality first: no client data in unapproved AI tools",
              "Value before technology: every use case has an owner and a measurable benefit",
              "Proportionate risk: controls scale with the risk tier of each use case",
              "Vendor-agnostic and secure by design",
              "Practise what we advise: GT tools and RCM applied to ourselves first"]

APPETITE = [("Client confidentiality & data protection", "Very low", "No client or personal data in unapproved AI tools; enterprise tools with no-training and in-region / approved hosting only."),
            ("Regulatory & professional-standards compliance", "Very low", "AI use must comply with PDPL, IESBA/ISQM 1 quality management and client contractual terms."),
            ("Quality of client deliverables", "Low", "Human review and sign-off mandatory for all AI-assisted client outputs; no autonomous client-facing AI without CoE approval."),
            ("Operational & productivity experimentation", "Moderate", "Experimentation encouraged within the approved sandbox, with synthetic / non-confidential data."),
            ("Innovation & new AI services", "Moderate-High", "Invest in new AI offerings and pilots through stage-gates with defined success criteria.")]

VALUE_CRIT = [("Strategic alignment", 0.25), ("Efficiency / financial benefit", 0.35), ("Quality & client impact", 0.25), ("Scalability / reuse", 0.15)]
FEAS_CRIT = [("Data readiness", 0.30), ("Technical simplicity", 0.25), ("Ease of change & adoption", 0.20), ("Tool / vendor availability", 0.25)]
SUBSCORES = {1: ([4, 4, 4, 4], [4, 5, 4, 5]), 2: ([3, 3, 3, 4], [5, 5, 4, 5]), 3: ([5, 4, 4, 4], [3, 4, 4, 4]), 4: ([5, 5, 5, 4], [3, 2, 3, 4]),
             5: ([4, 4, 4, 4], [3, 4, 4, 4]), 6: ([4, 4, 4, 3], [3, 3, 3, 4]), 7: ([3, 4, 3, 3], [3, 3, 3, 3]), 8: ([5, 4, 3, 4], [3, 3, 4, 4]),
             9: ([2, 4, 2, 3], [4, 4, 3, 5]), 10: ([2, 2, 3, 3], [5, 5, 5, 4]), 11: ([5, 5, 4, 4], [2, 1, 2, 3]), 12: ([5, 4, 4, 5], [3, 3, 4, 3])}
# id, name, value(1-5), feasibility(1-5), risk tier, horizon, area, description  (value / feasibility recomputed from SUBSCORES below)
USE_CASES = [
    (1, "Proposal & report drafting assistant", 4.0, 4.6, "Medium", "H1", "All service lines", "Enterprise GenAI with GT templates for proposals, reports and client letters; mandatory human review."),
    (2, "Meeting transcription & summaries", 3.2, 4.8, "Low", "H1", "Firm-wide", "Automated minutes, action tracking and summaries for internal meetings and client workshops (with consent)."),
    (3, "Knowledge assistant (methodology RAG)", 4.2, 4.0, "Medium", "H1", "Firm-wide", "Retrieval-augmented assistant over GT methodologies, standards, Bahrain laws and prior (sanitised) work."),
    (4, "Audit analytics & anomaly detection", 4.8, 3.0, "High", "H2", "Audit & Assurance", "AI-assisted journal-entry testing, anomaly detection and full-population analytics under ISA / ISQM 1."),
    (5, "Tax research & regulatory Q&A", 4.1, 3.8, "Medium", "H2", "Tax", "Grounded assistant for Bahrain VAT, corporate / top-up tax and GCC tax research with citations."),
    (6, "Contract & document review", 3.8, 3.1, "Medium", "H2", "Advisory / Due diligence", "Clause extraction, risk flagging and summarisation for due diligence and compliance reviews."),
    (7, "Client onboarding / KYC screening", 3.4, 3.0, "High", "H3", "Risk & Quality", "Automated entity screening, adverse-media review and independence checks."),
    (8, "AI-assisted control testing (RCM)", 3.9, 3.45, "Medium", "H2", "Risk Advisory / AI GRC", "Evidence review and control-test support using the GT AI RCM v2.1 test plans and AI security audit checklist."),
    (9, "Finance, timesheet & billing automation", 3.0, 4.0, "Low", "H2", "Operations", "RPA + AI for WIP, billing narratives and expense processing."),
    (10, "Marketing & thought-leadership content", 2.4, 4.7, "Low", "H1", "Marketing", "Drafting of insights, social posts and event content with brand review."),
    (11, "Agentic engagement workflow assistant", 4.3, 2.0, "High", "H3", "Firm-wide", "Multi-step agents orchestrating engagement set-up, requests-for-information and status reporting."),
    (12, "Client AI readiness self-assessment portal", 4.3, 3.3, "Medium", "H2", "AI service line", "Web version of the GT AI Readiness Assessment Tool as a lead-generation and delivery accelerator."),
]

OFFERINGS = [
    ("AI Governance, Risk & Compliance", ["AI Readiness & Risk Assessment (20 questions, 5 domains, auto-generated report)", "AI governance framework, policies & AI register", "RCM v2.1 compliance / gap review (EU AI Act, PDPL, CBB)", "Use-case AI risk & impact assessment", "ISO/IEC 42001 readiness & internal audit"], "Banks & insurers, government, telecom"),
    ("AI Strategy & Roadmap", ["AI strategy sprint (6-8 weeks): situational / SWOT, strategy, roadmap", "Use-case discovery & prioritisation workshops", "AI risk appetite & value-realisation framework"], "Mid-to-large corporates, family groups, government"),
    ("AI-Driven Automation & Apps", ["Technology evaluation & selection", "Process automation & AI integration", "Industry-specific solutions", "AI-powered BCP tool (in preparation)", "AI-powered Financial Statements tool (in preparation)"], "Financial services, SMEs, healthcare, hospitality"),
    ("Training & Awareness", ["AI literacy academy (all staff; EU AI Act Art. 4-aligned)", "Board & leadership AI governance sessions", "Immersive labs & AI-driven gamification", "Deepfake / AI-phishing simulations"], "All sectors; entry offer"),
]

KPIS = [("AI Ready7 overall readiness", "61.4% (Informed)", ">= 75% (Repeatable)", ">= 85%", "AI CoE lead"),
        ("Governance & Ethics pillar", "60%", ">= 80%", ">= 90%", "Risk & Quality partner"),
        ("AI use cases risk-tiered & assessed (AI register)", "Register exists; assessment not consistent", "100% by Q1 2027", "100% + quarterly reconciliation", "AI CoE"),
        ("Staff completing AI literacy", "To be baselined", ">= 90% by Q2 2027", "100% (annual refresh)", "HR / L&D"),
        ("Production AI use cases with measured benefits", "0 measured", "6", "10+", "Service-line leaders"),
        ("Time saved in targeted processes", "To be baselined", "10%", "15%", "CoE + Finance"),
        ("AI advisory engagements delivered (proposed)", "Pilot stage", "12", "25", "AI service line lead"),
        ("GT AI apps (BCP tool, Financial Statements tool)", "2 in preparation", "2 piloted with clients", "2 launched commercially", "AI service line lead"),
        ("Client-data leakage incidents via AI", "-", "0", "0", "CISO / DPO"),
        ("ISO/IEC 42001 status", "Not started", "Certification-ready", "Certified (optional)", "Risk & Quality partner")]

# workstream, activity, start month (0 = Oct 2026), duration (months)
GANTT = [
    ("WS1 Governance & Risk", "AI strategy, risk appetite & acceptable-use policy", 0, 2),
    ("WS1 Governance & Risk", "AI Steering Committee & CoE charter", 0, 2),
    ("WS1 Governance & Risk", "AI register refresh & risk-framework rollout", 1, 4),
    ("WS1 Governance & Risk", "Vendor AI contract clauses & due diligence", 2, 4),
    ("WS1 Governance & Risk", "ISO/IEC 42001 gap assessment & AIMS build", 5, 8),
    ("WS1 Governance & Risk", "Certification-readiness audit", 13, 3),
    ("WS2 People & Skills", "AI literacy baseline (all staff)", 1, 5),
    ("WS2 People & Skills", "Role-based pathways & AI champions network", 4, 7),
    ("WS2 People & Skills", "Leadership & client-facing enablement", 6, 6),
    ("WS2 People & Skills", "Annual refresh & workforce plan", 12, 6),
    ("WS3 Data, Platforms & Security", "Data classification for AI & GenAI DLP", 0, 4),
    ("WS3 Data, Platforms & Security", "Enterprise GenAI platform & reference architecture", 1, 6),
    ("WS3 Data, Platforms & Security", "AI red teaming & deepfake-resilient processes", 3, 6),
    ("WS3 Data, Platforms & Security", "Monitoring, KRIs, logging & ModelOps", 6, 7),
    ("WS4 Use Cases & Value", "Quick wins: drafting, transcription, knowledge assistant", 1, 5),
    ("WS4 Use Cases & Value", "Pilots: audit analytics, tax assistant, document review", 5, 6),
    ("WS4 Use Cases & Value", "Scale proven use cases; agentic pilots", 11, 7),
    ("WS5 AI Service Line & Apps", "Package offerings & tools; pricing approach", 0, 3),
    ("WS5 AI Service Line & Apps", "AI apps (BCP, Financial Statements): design & MVP", 1, 6),
    ("WS5 AI Service Line & Apps", "AI apps: risk assessment, internal pilot & client launch", 7, 7),
    ("WS5 AI Service Line & Apps", "Launch & first pilot clients", 3, 4),
    ("WS5 AI Service Line & Apps", "Sector campaigns (FS, government, telecom)", 6, 7),
    ("WS5 AI Service Line & Apps", "Regional scale via the GT network", 12, 6),
]

HUNDRED_DAYS = [
    ("Days 0-30", [("Approve AI strategy, objectives & risk appetite", "Managing Partner"),
                   ("Constitute AI Steering Committee and AI CoE", "Managing Partner"),
                   ("Issue AI acceptable-use policy & approved-tools list", "Risk & Quality"),
                   ("Enforce data-classification rules & DLP for GenAI", "IT / InfoSec")]),
    ("Days 31-60", [("Refresh AI register; tier & assess all AI tools", "AI CoE"),
                    ("Launch AI literacy baseline for all staff", "HR / L&D"),
                    ("Start quick-win pilots (drafting, transcription, knowledge)", "AI CoE + service lines"),
                    ("Package entry offers; confirm MVP scope of AI BCP & Financial Statements apps", "AI service line lead")]),
    ("Days 61-100", [("Reference architecture v1 and AI logging", "IT"),
                     ("AI red-team of internal GenAI assistant", "Cyber team"),
                     ("First 3 client pilots of the readiness assessment", "AI service line lead"),
                     ("KPI baseline and AI Ready7 checkpoint to leadership", "AI CoE")]),
]

STRAT_RISKS = [("Low adoption / change resistance", "Visible leadership sponsorship, AI champions, quick wins communicated firm-wide"),
               ("Client-data leakage through AI tools", "Acceptable-use policy, enterprise tools only, DLP, no-training contract clauses"),
               ("Investment without measurable return", "Stage-gates, benefit KPIs per use case, quarterly portfolio review"),
               ("Shortage of AI talent", "Targeted hiring, GTIL expertise sharing, technology partners, upskilling"),
               ("Reputational harm from AI errors", "Human-in-charge sign-off, QA sampling, professional-standards alignment"),
               ("Regulatory change (Bahrain AI law, EU AI Act)", "Obligations register and horizon scanning (RCM RO-5)"),
               ("Liability from GT AI apps (BCP, financial statements outputs)", "Apps assessed under the GT AI Risk Assessment Framework; qualified professional review of outputs; clear terms of use; independence checks for audit clients"),
               ("Vendor lock-in / platform change", "Multi-model reference architecture, exit clauses, periodic re-evaluation")]


def _w(sc, crit):
    return round(sum(a * w for a, (_, w) in zip(sc, crit)), 2)


USE_CASES = [(u[0], u[1], _w(SUBSCORES[u[0]][0], VALUE_CRIT), _w(SUBSCORES[u[0]][1], FEAS_CRIT)) + tuple(u[4:]) for u in USE_CASES]


def quadrant(v, f):
    return "Quick win" if v >= 3.5 and f >= 3.5 else "Strategic bet" if v >= 3.5 else "Fill-in" if f >= 3.5 else "Deprioritise"


APPS = [  # name, purpose, service line, status, next milestone
    ("GT AI Readiness & Risk Assessment Tool", "20-question, 5-domain organisational AI readiness & risk-exposure assessment with auto-generated report", "AI GRC", "Available (v2.0)", "Client pilots in H1"),
    ("GT AI RCM v2.1 + AI Security & Agent Audit Checklist", "81-control AI risk & control matrix and 125-item security / agent audit programme", "AI GRC / Cyber", "Available", "First engagements in H1"),
    ("GT AI Risk Assessment Framework & Toolkit", "Use-case tiering, 52-risk taxonomy, register, heat map and sign-off", "AI GRC", "Available (v1.1)", "Embed in AI CoE intake"),
    ("AI-powered BCP tool", "AI-assisted business impact analysis, BCP / DR plan drafting and tabletop scenarios", "AI Automation / Risk Advisory", "In preparation", "MVP & internal pilot (indicative H1-H2)"),
    ("AI-powered Financial Statements tool", "AI-assisted preparation and review of IFRS financial statements and disclosures", "AI Automation / Accounting Advisory", "In preparation", "MVP & internal pilot (indicative H1-H2)"),
    ("Client AI self-assessment portal", "Web version of the readiness tool for lead generation and delivery", "AI GRC", "Planned", "Build in H2"),
]
APP_DETAIL = [
    ("AI-powered BCP Tool", [
        ("Why", "Business continuity plans are often generic, outdated and slow to build; regulated clients (e.g., CBB licensees) must evidence tested BCM."),
        ("What it does", "Guided business impact analysis with AI-suggested impacts, RTO / RPO and dependencies; auto-drafted BCP / DR plans aligned to ISO 22301; AI-generated scenarios and tabletop exercises; gap scoring and maintenance reminders."),
        ("Who for", "Banks & insurers, government entities, telecom, healthcare, hospitality and SMEs; links to Training & Awareness BC gamification."),
        ("Guardrails", "Risk tier Medium; human validation of BIA and plan content; client data in approved hosting; no autonomous actions; covered by RCM v2.1 controls."),
    ]),
    ("AI-powered Financial Statements Tool", [
        ("Why", "Preparing and reviewing IFRS financial statements is manual and error-prone; AI can speed mapping, drafting and consistency checks."),
        ("What it does", "Trial-balance import with AI-assisted mapping to statement line items; draft primary statements and notes from GT templates; IFRS disclosure checklist with AI suggestions; tie-outs, anomaly and variance flags."),
        ("Who for", "SMEs and family groups, non-audit clients, GT accounting-advisory teams; internal review aid for engagement teams."),
        ("Guardrails", "Risk tier High (financial-reporting impact): qualified accountant review and sign-off mandatory; full audit trail of AI suggestions; independence rules - not offered as a preparation service to audit clients where prohibited (IESBA Code)."),
    ]),
]
