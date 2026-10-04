# -*- coding: utf-8 -*-
"""Content for the GT AI Opportunity Discovery Toolkit - "where exactly does AI fit in this organisation?".

The toolkit is deliberately NOT a readiness test (that is the GT AI Readiness & Risk Assessment). It answers:
  1. Position  - does the organisation have the data and the scale / capacity for AI, or is automation enough for now?
  2. Discover  - which processes are candidates (process inventory, supported by a cross-industry process library)?
  3. Diagnose  - for each process: what kind of solution fits (digitise, automate, off-the-shelf AI, configured AI,
                 custom AI / agents), is the data there, what is the risk tier?
  4. Prioritise- value x feasibility, risk-adjusted priority, horizon.
  5. Plan      - opportunity map, roadmap lanes and one-page use-case canvases.
"""

# ------------------------------------------------------------------ Step 1 - organisational AI position
# (id, axis, question, guidance, 4 options from lowest to highest position)
POSITION_Q = [
    ("D1", "Data", "How are your core business processes run today?",
     "Think of the 3-5 processes that generate most of your revenue or workload.",
     ["Mostly on paper or manually", "Mainly spreadsheets and email",
      "Core systems (e.g., ERP, CRM, core banking) for the main processes",
      "Integrated systems end-to-end with little manual re-keying"]),
    ("D2", "Data", "How much digital history do you keep for your core activities (transactions, customers, cases, operations)?",
     "AI that predicts, scores or detects patterns learns from history.",
     ["Less than 1 year, or scattered across files", "1-2 years, held in different places",
      "3+ years in core systems", "3+ years consolidated in a data warehouse / data platform"]),
    ("D3", "Data", "How reliable is that data, and does it have owners?",
     "Data quality and ownership - ISO/IEC 5259 / ISO/IEC 42001 Annex A.7.",
     ["No defined owners; frequent errors and duplicates", "Some owners; known quality issues not tracked",
      "Owners for key data; quality checked periodically", "Governed data with owners, definitions and quality metrics"]),
    ("D4", "Data", "Where do your documents and knowledge live (contracts, policies, reports, emails, manuals)?",
     "Generative AI and knowledge assistants work on documents rather than tables.",
     ["Paper or personal drives", "Shared drives / email, not organised",
      "A document-management system or SharePoint with structure", "Well-organised repository with metadata and access control"]),
    ("D5", "Data", "Can data move between your systems?",
     "Integration determines whether AI outputs can be used inside the workflow.",
     ["No - data is re-keyed manually", "Manual exports / imports (e.g., CSV)",
      "Some APIs or integrations between key systems", "Integration platform / APIs and a central data layer"]),
    ("D6", "Data", "Do you record outcomes against your data (e.g., fraud confirmed, claim paid, customer churned, invoice disputed)?",
     "Recorded outcomes ('labels') are what supervised machine learning needs.",
     ["No", "Occasionally / informally", "For some key processes", "Systematically for most key processes"]),
    ("S1", "Scale", "How many employees does the organisation have?", "Scale drives the size of the prize and what can be sustained.",
     ["Fewer than 50", "50 - 249", "250 - 999", "1,000 or more"]),
    ("S2", "Scale", "How many repetitive transactions / cases does your largest process handle per month?",
     "e.g., invoices, applications, claims, customer requests, bookings.",
     ["Fewer than 500", "500 - 5,000", "5,000 - 50,000", "More than 50,000"]),
    ("S3", "Scale", "What technology capability do you have in-house?", "Who would configure, integrate and run a solution?",
     ["No IT staff / basic outsourced support", "Small IT team (support and vendor management)",
      "IT team with integration / development skills", "IT plus a data / analytics team"]),
    ("S4", "Scale", "What budget could be made available for digital / AI initiatives in the next 12 months?",
     "Indicative only - used to size the type of solution that is realistic.",
     ["Below BHD 10,000", "BHD 10,000 - 50,000", "BHD 50,000 - 250,000", "Above BHD 250,000"]),
    ("S5", "Scale", "How many customer / citizen interactions do you handle through digital channels?",
     "Web, app, email, WhatsApp, call centre.",
     ["Few - mostly in person", "Moderate", "High - most interactions are digital", "Very high - omnichannel at scale"]),
    ("S6", "Scale", "What capacity do you have to run change projects alongside day-to-day work?",
     "Leadership bandwidth, project management, change management.",
     ["None at the moment", "One project at a time", "Several projects in parallel", "Dedicated transformation / innovation team"]),
]

# code, archetype, solution ceiling (level), headline, what it means
ARCHETYPES = [
    ("A", "Automation is enough - for now", 3,
     "Digitise and automate first; use AI that is already built into everyday tools.",
     "Data and scale do not yet justify building AI. Most value comes from digitising paper steps, workflow / RPA automation, "
     "dashboards and safe use of off-the-shelf AI (e.g., AI features in Microsoft 365, accounting or CRM software) under an acceptable-use policy."),
    ("B", "Data foundations first", 3,
     "Scale is there but the data is not: fix data foundations while automating and adopting off-the-shelf AI.",
     "Volumes make AI attractive, but processes and data are not yet digital, joined-up or reliable enough. Prioritise data ownership, "
     "integration and capturing outcomes; automate high-volume steps; adopt packaged AI. Configured and custom AI become viable once the data gaps close."),
    ("C", "Buy and configure AI", 4,
     "Good data, limited scale / capacity: adopt and configure proven AI products rather than build.",
     "The data supports AI, but in-house capacity and scale favour buying. Configure AI products on your own data (knowledge assistants, "
     "intelligent document processing, conversational AI, vendor ML modules). Custom models and autonomous agents only through partners and pilots."),
    ("D", "Build and scale AI", 5,
     "Data and scale support custom AI: build a portfolio, including advanced analytics and governed agentic AI.",
     "The organisation can sustain a portfolio of AI use cases, including custom machine-learning models, retrieval-augmented assistants and "
     "agentic workflows - provided governance, MLOps and AI security scale with it."),
]

LEVELS = [(1, "Digitise / redesign"), (2, "Automate (no AI)"), (3, "Off-the-shelf / embedded AI"), (4, "Configured AI on own data"),
          (5, "Custom AI / agentic AI")]

# ------------------------------------------------------------------ Step 3 - task nature -> solution family
# nature (what the work is), family, level, AI? (No / Light / Yes), data required (0-3), indicative effort reduction,
# typical solutions, example KPIs, key design notes / standards
FAMILIES = [
    ("Move or re-key data between systems following fixed rules", "Workflow automation / RPA", 2, "No", 1, 0.60,
     "Power Automate, UiPath, Zapier / Make, ERP workflow", "Cycle time; manual touches per case; error rate",
     "Rules are stable and documented - AI adds little. Automate first."),
    ("Check or approve items against fixed criteria (rules, thresholds, checklists)", "Rules engine / workflow automation", 2, "No", 1, 0.50,
     "Business-rules engines, ERP / core-system workflow, low-code apps", "Turnaround time; exceptions handled; compliance rate",
     "Encode the criteria; route exceptions to people. Use AI only for the unstructured part, if any."),
    ("Produce reports, dashboards or reconciliations from existing data", "BI, dashboards & reconciliation automation", 2, "No", 2, 0.25,
     "Power BI, Tableau, Excel automation, reconciliation tools", "Report preparation time; timeliness; reconciliation breaks",
     "Usually a data-modelling problem, not an AI problem. GenAI 'ask your data' can be layered on later."),
    ("Extract information from documents, forms or emails", "Intelligent document processing (IDP)", 3, "Light", 1, 0.50,
     "Azure AI Document Intelligence, AWS Textract, ABBYY, Rossum; OCR + validation rules", "Straight-through processing rate; extraction accuracy; cost per document",
     "Needs digital or scanned documents and a validation step. Arabic / bilingual accuracy to be tested."),
    ("Draft, summarise, translate or review text", "Generative AI assistant (off-the-shelf)", 3, "Yes", 0, 0.30,
     "Microsoft 365 Copilot, ChatGPT Enterprise, Gemini for Workspace, Claude for Enterprise", "Drafting time; adoption; quality-review findings",
     "No own data needed, but needs acceptable-use policy, approved enterprise tools and human review (accuracy, confidentiality)."),
    ("Answer questions from internal knowledge, policies or past work", "GenAI knowledge assistant (RAG on own content)", 4, "Yes", 1, 0.25,
     "Copilot Studio, Azure AI Search + LLM, Glean, vendor knowledge bots", "Time to answer; deflected queries; answer accuracy",
     "Quality depends on organised, current, access-controlled content (D4). Test for hallucination and data leakage."),
    ("Converse with customers, citizens or staff (chat, voice, WhatsApp)", "Conversational AI / virtual agent", 4, "Yes", 1, 0.30,
     "Copilot Studio, Dialogflow CX, Amazon Lex, contact-centre AI platforms", "Containment rate; CSAT; average handling time",
     "Transparency to users (EU AI Act Art. 50); clear hand-off to humans; bilingual Arabic / English design."),
    ("Recognise or inspect images, video or speech", "Computer vision / speech AI", 4, "Yes", 1, 0.40,
     "Packaged vision / speech services, industry inspection products", "Inspection throughput; detection accuracy; false alarms",
     "Biometric and surveillance uses can be prohibited or high-risk (EU AI Act Art. 5 / Annex III; PDPL sensitive data)."),
    ("Forecast, score, classify or detect anomalies from historical data", "Predictive analytics / machine learning", 5, "Yes", 2, 0.20,
     "Azure ML / Databricks / SageMaker, vendor ML modules (fraud, demand, churn), AutoML", "Forecast accuracy; detection rate; losses avoided",
     "Needs history (D2) and ideally recorded outcomes (D6). Scoring of individuals (credit, hiring, insurance) is high-risk: impact assessment and bias testing."),
    ("Coordinate multi-step work across systems that needs judgement", "Agentic AI (supervised agents)", 5, "Yes", 2, 0.40,
     "Copilot Studio agents, Azure AI Foundry Agent Service, Amazon Bedrock Agents, LangGraph", "Cases completed end-to-end; human interventions; error / rollback rate",
     "Least agency and least privilege, human approval for high-impact actions, tool allow-lists, logging (ISACA Securing AI Agents 2026; OWASP LLM06 Excessive Agency)."),
]

DATA_AVAIL = [("None - not captured", 0), ("Partial or paper-based", 1), ("Digital history available", 2),
              ("Digital history with recorded outcomes", 3)]
VOLUME = [("Low (fewer than 100 a month)", 1), ("Medium (100 - 1,000 a month)", 3), ("High (more than 1,000 a month)", 5)]
STANDARD = [("Yes - documented and consistent", 5), ("Partly - common steps, some variation", 3), ("No - varies by person / team", 1)]
AUTONOMY = [("Assist - AI suggests, a person does the work", 0), ("Human approves each AI output", 1), ("Fully automated - no human review", 2)]
YESNO = ["Yes", "No"]
VERDICTS = [  # verdict, meaning, next step
    ("AI fit now", "AI is the right tool, the data exists and it is within the organisation's current position.",
     "Shape the use case (canvas), run the GT AI risk assessment for its tier, select a tool and pilot within 3 months with KPIs."),
    ("Automate - AI not needed", "The work follows fixed rules or is a data / reporting problem; automation or BI delivers the value with less risk and cost.",
     "Map the process, automate with workflow / RPA / BI and measure cycle-time savings. Revisit AI later for the exceptions."),
    ("AI later - close data gap first", "AI would fit, but the data it needs is not captured, digital or labelled yet.",
     "Start capturing digital history and outcomes, assign a data owner and fix quality; re-diagnose in 6-12 months."),
    ("Stretch - partner or later stage", "AI fits and the data exists, but the solution is more complex than the organisation can build or run today.",
     "Use a vendor product or delivery partner, or run a time-boxed pilot; build capability before scaling."),
    ("Redesign / standardise first", "The process varies by person or team; automating or adding AI would scale inconsistency.",
     "Standardise and document the process (and digitise steps); then re-diagnose - it often becomes an automation quick win."),
]
HORIZONS = ["H1 - Now (0-6 months)", "H2 - Next (6-12 months)", "H3 - Later (12-24 months)", "Backlog / not prioritised"]

CONTROLS = [  # risk tier, minimum governance before go-live (GT AI Risk Assessment Framework / RCM v2.1)
    ("Low", "Approved tool on the AI register; acceptable-use policy; user guidance; spot-check outputs; vendor terms checked (no training on your data)."),
    ("Medium", "As Low, plus: documented use-case risk assessment (GT AI Risk Assessment Framework, NSW AIAF-aligned); named owner; human review of outputs; "
               "PDPL check where personal data is used; testing before go-live; monitoring KPIs / KRIs."),
    ("High", "As Medium, plus: AI system impact assessment (ISO/IEC 42005) and DPIA; bias, robustness and security testing (incl. AI red teaming); "
             "human-in-the-loop for decisions about individuals; EU AI Act / CBB obligations checked; executive / AI committee approval; GT AI RCM v2.1 controls applied."),
]

# ------------------------------------------------------------------ standards basis (step, standard, how it is used)
STANDARDS = [
    ("1 Position", "OECD Framework for the Classification of AI Systems (2022)", "Economic Context and Data & Input dimensions shape the position screen: scale, sector, data provenance, structure and availability."),
    ("1 Position", "ISO/IEC 42001:2023 cl. 4.1 - 4.2", "Understanding the organisation, its context and interested parties before deciding where AI is used."),
    ("1 Position", "NIST AI RMF 1.0 - MAP 1.3, 1.4, 1.6", "Organisational mission, business value and system requirements established before AI is selected."),
    ("2 Discover", "APQC Process Classification Framework (cross-industry)", "Function taxonomy of the 240-entry GT AI Use-Case Library (11 functions mapped to APQC level-1 categories)."),
    ("2 Discover", "Sector regulation (CBB Rulebook, NHRA, TRA, iGA AI Policy, PDPL)", "Sector entries of the library carry the relevant Bahrain regulator and rule references and a minimum risk tier."),
    ("2 Discover", "NIST AI RMF 1.0 - MAP 1.1", "Intended purpose, context of use and users documented for each candidate process."),
    ("3 Diagnose", "ISO/IEC 22989:2022 and ISO/IEC 23053:2022", "AI concepts, tasks and ML system components - basis for the task-nature to solution-family mapping."),
    ("3 Diagnose", "OECD AI classification - AI Model and Task & Output dimensions", "Distinguishes recognition, detection, forecasting, personalisation, goal-driven optimisation and generation tasks."),
    ("3 Diagnose", "ISO/IEC 5259 series (2024-2025) and ISO/IEC 8183:2023", "Data quality and data life cycle for analytics and ML - basis for the per-process data gate."),
    ("3 Diagnose", "ISO/IEC 42001:2023 Annex A.7 (Data for AI systems)", "Data acquisition, quality, provenance and preparation expected before AI is deployed."),
    ("3 Diagnose", "NSW AI Assessment Framework (AIAF)", "Inherent-risk questions (impact on people, autonomy, data sensitivity) behind the risk tier."),
    ("3 Diagnose", "EU AI Act (Reg. 2024/1689) Art. 5, 6, 50 and Annex III", "Library minimum risk tiers: high-risk uses (credit scoring, life & health insurance pricing, employment, education, public assistance, emergency triage) are floored at High; prohibited practices and transparency duties noted."),
    ("3 Diagnose", "Bahrain PDPL (Law No. 30 of 2018)", "Personal and sensitive data flag; lawful basis and DPIA expectations for higher-risk processing."),
    ("3 Diagnose", "ISACA - Cybersecurity Recommendations for Securing AI Agents (2026); OWASP Top 10 for LLM Applications (2025)", "Autonomy rating and agentic guardrails: least agency, human approval for high-impact actions, excessive-agency risk."),
    ("4 Prioritise", "NIST AI RMF 1.0 - MAP 3.1, 3.2", "Expected benefits and costs (incl. of errors) weighed per use case."),
    ("4 Prioritise", "ISO/IEC 42001:2023 cl. 6.2", "AI objectives that are measurable - feed the KPIs on each canvas."),
    ("5 Plan", "ISO/IEC 42005:2025 and ISO/IEC 23894:2023", "AI system impact assessment and risk management for Medium / High-tier use cases before go-live."),
    ("5 Plan", "ISO/IEC 5338:2023 and ISO/IEC 42001 Annex A.6", "AI system life cycle from design to retirement for each roadmap item."),
    ("5 Plan", "GT AI Risk Assessment Framework v1.1 and GT AI RCM v2.1", "Tier-based minimum controls on each use-case canvas; control testing once in production."),
    ("Complementary", "GT AI Ready7 / GT AI Readiness & Risk Assessment", "Measures governance and capability maturity. Not repeated here - this toolkit asks WHERE AI fits, not HOW READY the organisation is."),
]

# ------------------------------------------------------------------ worked examples (illustrative inputs)
# position answers are option indices 0-3; processes reference the library by sector code + process name:
# (library process, volume idx, hrs / month, standardised idx, data idx, strategic, impact, affects rights, personal data, autonomy idx)
EXAMPLES = {
 "gt": dict(
    file_suffix="Worked_Example_GT_Bahrain",
    profile={"Organisation": "Grant Thornton Bahrain (illustrative)", "Sector": "Professional services", "Scope": "Whole firm - audit, tax, advisory and support functions",
             "Assessment date": "04/10/2026", "Facilitated by": "GT AI Service Line"},
    position={"D1": 2, "D2": 2, "D3": 1, "D4": 2, "D5": 2, "D6": 1, "S1": 1, "S2": 1, "S3": 1, "S4": 2, "S5": 1, "S6": 2},
    processes=[
     ("Audit - journal-entry testing (full population)", 2, 300, 1, 2, 5, 5, "No", "Yes", 1),
     ("Audit - working papers & memo drafting", 1, 400, 0, 2, 4, 4, "No", "Yes", 1),
     ("Audit - bank & external confirmations", 2, 120, 0, 2, 3, 3, "No", "Yes", 1),
     ("Tax research & technical Q&A", 1, 160, 1, 2, 4, 4, "No", "No", 0),
     ("Tax - VAT / tax return data preparation", 2, 200, 0, 2, 3, 4, "No", "Yes", 1),
     ("Predictive tax scenario modelling", 0, 80, 1, 2, 4, 4, "No", "Yes", 0),
     ("Proposal, quotation & tender responses", 1, 350, 1, 2, 4, 4, "No", "No", 0),
     ("Advisory - due-diligence data-room review", 1, 220, 1, 2, 4, 4, "No", "Yes", 1),
     ("Meeting minutes & action tracking", 2, 250, 0, 1, 3, 3, "No", "No", 0),
     ("Methodology & knowledge Q&A", 2, 200, 1, 2, 4, 3, "No", "No", 0),
     ("Client acceptance & independence checks", 1, 140, 0, 2, 4, 4, "No", "Yes", 1),
     ("Timesheets, WIP & billing", 2, 180, 0, 3, 3, 3, "No", "No", 1),
     ("Expense claims audit", 1, 60, 0, 2, 2, 2, "No", "Yes", 1),
     ("CV screening & shortlisting", 1, 80, 1, 1, 3, 3, "Yes", "Yes", 1),
     ("Employee HR queries (leave, policies, payslips)", 1, 60, 0, 2, 2, 3, "No", "Yes", 0),
     ("Marketing content & campaigns", 1, 80, 0, 0, 3, 2, "No", "No", 0),
     ("Engagement workflow coordination", 1, 150, 2, 1, 4, 3, "No", "Yes", 1),
     ("KPI dashboards", 0, 100, 1, 2, 4, 3, "No", "No", 1),
     ("Assessment & maturity report production", 0, 40, 0, 2, 4, 4, "No", "No", 1),
     ("Client BCP & business impact analysis (GT AI BCP tool - in preparation)", 0, 200, 1, 1, 5, 4, "No", "Yes", 1),
     ("Financial statements preparation (GT AI FS tool - in preparation)", 1, 260, 0, 2, 5, 5, "No", "Yes", 1),
     ("Customer churn prediction", 1, 200, 1, 1, 4, 4, "No", "Yes", 0),
    ]),
 "bank": dict(
    file_suffix="Worked_Example_Retail_Bank",
    profile={"Organisation": "Illustrative Bahrain retail bank", "Sector": "Banking & payments", "Scope": "Retail and SME banking, operations and support functions",
             "Assessment date": "04/10/2026", "Facilitated by": "GT AI Service Line"},
    position={"D1": 2, "D2": 2, "D3": 2, "D4": 2, "D5": 2, "D6": 1, "S1": 2, "S2": 3, "S3": 2, "S4": 3, "S5": 2, "S6": 2},
    processes=[
     ("Retail customer onboarding (eKYC)", 2, 400, 0, 2, 5, 5, "Yes", "Yes", 1),
     ("Sanctions & PEP screening alert review", 2, 600, 0, 3, 5, 4, "No", "Yes", 1),
     ("AML transaction-monitoring alert triage", 2, 1200, 1, 3, 5, 5, "No", "Yes", 1),
     ("Suspicious transaction report (STR) narratives", 1, 160, 1, 2, 4, 4, "No", "Yes", 1),
     ("Retail credit decisioning", 2, 500, 0, 3, 5, 5, "Yes", "Yes", 1),
     ("SME credit memo drafting", 1, 300, 1, 2, 4, 4, "No", "Yes", 0),
     ("Trade-finance document checking (LCs)", 1, 350, 0, 2, 4, 4, "No", "Yes", 1),
     ("Card & payment fraud detection", 2, 300, 0, 3, 5, 5, "No", "Yes", 2),
     ("Payment exceptions & repairs", 2, 250, 0, 2, 3, 4, "No", "Yes", 2),
     ("Banking virtual assistant", 2, 900, 1, 2, 4, 5, "No", "Yes", 0),
     ("Complaints handling (CBB timelines)", 1, 200, 2, 2, 4, 5, "No", "Yes", 1),
     ("Collections prioritisation", 2, 400, 1, 1, 4, 4, "Yes", "Yes", 1),
     ("Regulatory reporting (CBB returns)", 0, 300, 1, 2, 5, 3, "No", "No", 1),
     ("ATM & branch cash forecasting", 1, 80, 0, 3, 3, 3, "No", "No", 1),
     ("Corporate onboarding (KYB)", 1, 250, 1, 2, 4, 4, "No", "Yes", 1),
     ("Card dispute & chargeback handling", 2, 300, 0, 1, 3, 4, "No", "Yes", 1),
     ("Periodic KYC refresh", 2, 500, 1, 2, 4, 4, "No", "Yes", 1),
     ("Mortgage document verification", 1, 120, 0, 2, 3, 4, "No", "Yes", 1),
     ("IFRS 9 ECL staging & overlays support", 0, 200, 1, 1, 4, 4, "No", "Yes", 1),
     ("IT service desk", 2, 300, 0, 2, 3, 3, "No", "No", 0),
     ("Security alert triage (SOC)", 2, 400, 1, 2, 5, 4, "No", "Yes", 1),
     ("Employee HR queries (leave, policies, payslips)", 1, 80, 0, 2, 2, 3, "No", "Yes", 0),
     ("Supplier invoice capture & posting", 1, 150, 0, 2, 2, 2, "No", "No", 1),
     ("Regulatory change monitoring", 0, 120, 1, 2, 4, 3, "No", "No", 0),
     ("Meeting minutes & action tracking", 1, 150, 0, 1, 3, 2, "No", "No", 0),
    ]),
}
