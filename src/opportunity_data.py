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

SECTORS = ["Banking & financial services", "Insurance", "Government & public sector", "Healthcare", "Hospitality & tourism",
           "Retail & consumer", "Manufacturing & industrial", "Logistics & transport", "Telecom & technology", "Real estate & construction",
           "Professional services", "Education", "Oil, gas & energy", "Other"]

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
FUNCTIONS = ["Core service delivery / operations", "Customer service & sales", "Marketing & communications", "Finance & accounting", "HR & people",
             "Procurement & supply chain", "Risk, compliance & legal", "IT & security", "Management & reporting", "Knowledge & collaboration"]

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

# ------------------------------------------------------------------ Process library (cross-industry + Bahrain priority sectors)
# function, process / activity, typical pain point, typical task nature (index into FAMILIES), sectors
LIBRARY = [
    ("Finance & accounting", "Supplier invoice processing (AP)", "Manual keying of invoices from PDF / email into the ERP", 3, "All"),
    ("Finance & accounting", "Three-way match and payment approval", "Checking PO, GRN and invoice by hand", 1, "All"),
    ("Finance & accounting", "Bank and intercompany reconciliations", "Spreadsheet-based matching each month", 2, "All"),
    ("Finance & accounting", "Month-end close and management reporting", "Late, manual report packs", 2, "All"),
    ("Finance & accounting", "Expense claims", "Receipt checking and policy compliance", 3, "All"),
    ("Finance & accounting", "Cash-flow forecasting", "Forecasts in spreadsheets, poor accuracy", 8, "All"),
    ("Finance & accounting", "Credit control and collections", "Chasing overdue invoices manually", 0, "All"),
    ("Finance & accounting", "VAT return preparation", "Collating transactions and VAT codes manually", 0, "All"),
    ("Finance & accounting", "Financial statements and disclosure drafting", "Mapping trial balance and drafting notes manually", 4, "All"),
    ("HR & people", "Recruitment - CV screening and shortlisting", "High application volumes, slow shortlisting", 8, "All"),
    ("HR & people", "Employee HR queries (leave, policies, payslips)", "HR team answering the same questions", 6, "All"),
    ("HR & people", "Onboarding and offboarding", "Many hand-offs across HR, IT and facilities", 0, "All"),
    ("HR & people", "Payroll inputs and validation", "Manual collation of attendance and allowances", 1, "All"),
    ("HR & people", "Job descriptions, policies and training content", "Time-consuming drafting", 4, "All"),
    ("HR & people", "Attrition risk analysis", "Losing key staff without warning", 8, "Medium-large organisations"),
    ("Customer service & sales", "Customer enquiries (web, WhatsApp, phone)", "Long waits for routine questions", 6, "All"),
    ("Customer service & sales", "Complaint handling and categorisation", "Manual triage and late responses", 4, "All"),
    ("Customer service & sales", "Proposal, quotation and tender responses", "Re-writing similar content", 4, "All"),
    ("Customer service & sales", "Lead scoring and next-best offer", "Sales effort spread evenly across leads", 8, "Banking, telecom, retail"),
    ("Customer service & sales", "Customer churn prediction", "Customers leave without warning", 8, "Banking, telecom, insurance"),
    ("Customer service & sales", "Call summarisation and quality review", "Supervisors sample only a few calls", 4, "Contact centres"),
    ("Marketing & communications", "Marketing content and social media", "Slow content production", 4, "All"),
    ("Marketing & communications", "Arabic / English translation", "External translation cost and delays", 4, "All"),
    ("Marketing & communications", "Customer review and sentiment analysis", "Feedback not analysed", 4, "Hospitality, retail, healthcare"),
    ("Procurement & supply chain", "Purchase requisition to order", "Approvals chased by email", 1, "All"),
    ("Procurement & supply chain", "Supplier onboarding and due diligence", "Collecting documents and checks manually", 3, "All"),
    ("Procurement & supply chain", "Contract review and obligation tracking", "Missed renewal dates and terms", 4, "All"),
    ("Procurement & supply chain", "Demand forecasting and inventory planning", "Stock-outs and excess stock", 8, "Retail, manufacturing, hospitality"),
    ("Procurement & supply chain", "Route and delivery optimisation", "Inefficient routes, late deliveries", 8, "Logistics, retail"),
    ("Risk, compliance & legal", "KYC / customer due diligence", "Manual document checks and screening", 3, "Banking, insurance, professional services"),
    ("Risk, compliance & legal", "AML transaction monitoring alert triage", "High false-positive alert volumes", 8, "Banking, payments"),
    ("Risk, compliance & legal", "Regulatory change monitoring (CBB, PDPL, NBR)", "Tracking new rules manually", 5, "Regulated sectors"),
    ("Risk, compliance & legal", "Policy and procedure Q&A", "Staff unsure where rules are", 5, "All"),
    ("Risk, compliance & legal", "Internal audit testing and sampling", "Sample-based testing of large populations", 8, "Medium-large organisations"),
    ("Risk, compliance & legal", "Business continuity plans and BIA", "Plans outdated and generic", 4, "Regulated sectors"),
    ("IT & security", "IT service desk tickets", "Repetitive password and access requests", 6, "All"),
    ("IT & security", "User access reviews", "Manual reconciliation of access lists", 1, "All"),
    ("IT & security", "Security alert triage (SOC)", "Alert fatigue", 9, "Medium-large organisations"),
    ("IT & security", "Code and script development", "Slow development, little documentation", 4, "Organisations with IT teams"),
    ("Management & reporting", "Board and management packs", "Data gathered by hand from many sources", 2, "All"),
    ("Management & reporting", "KPI dashboards", "No single view of performance", 2, "All"),
    ("Management & reporting", "Meeting minutes and action tracking", "Minutes late, actions lost", 4, "All"),
    ("Knowledge & collaboration", "Internal knowledge search", "Time lost looking for documents", 5, "All"),
    ("Knowledge & collaboration", "Email and document drafting", "Routine drafting effort", 4, "All"),
    ("Core service delivery / operations", "Banking - trade-finance document checking", "Manual checking of LCs and shipping documents", 3, "Banking"),
    ("Core service delivery / operations", "Banking - credit application assessment", "Slow credit decisions", 8, "Banking (high-risk: individuals)"),
    ("Core service delivery / operations", "Insurance - claims intake (FNOL)", "Claims forms keyed manually", 3, "Insurance"),
    ("Core service delivery / operations", "Insurance - claims fraud detection", "Fraud found late or not at all", 8, "Insurance"),
    ("Core service delivery / operations", "Insurance - underwriting submissions", "Broker submissions read manually", 3, "Insurance"),
    ("Core service delivery / operations", "Government - permit and licence applications", "Paper-heavy applications and checks", 1, "Government"),
    ("Core service delivery / operations", "Government - citizen enquiries", "Call-centre peaks", 6, "Government"),
    ("Core service delivery / operations", "Healthcare - appointment scheduling and reminders", "No-shows and phone booking", 0, "Healthcare"),
    ("Core service delivery / operations", "Healthcare - clinical documentation", "Clinicians typing notes", 4, "Healthcare (patient data)"),
    ("Core service delivery / operations", "Healthcare - medical-image triage", "Radiology backlogs", 7, "Healthcare (high-risk)"),
    ("Core service delivery / operations", "Hospitality - dynamic pricing and revenue management", "Prices set manually", 8, "Hospitality"),
    ("Core service delivery / operations", "Hospitality - guest messaging and concierge", "Front desk overloaded", 6, "Hospitality"),
    ("Core service delivery / operations", "Manufacturing - predictive maintenance", "Unplanned downtime", 8, "Manufacturing, oil & gas"),
    ("Core service delivery / operations", "Manufacturing - visual quality inspection", "Manual inspection misses defects", 7, "Manufacturing"),
    ("Core service delivery / operations", "Telecom - network fault prediction", "Outages detected by customers", 8, "Telecom"),
    ("Core service delivery / operations", "Real estate - lease abstraction", "Lease terms extracted by hand", 3, "Real estate"),
    ("Core service delivery / operations", "Professional services - audit journal-entry testing", "Sample-based testing", 8, "Audit firms"),
    ("Core service delivery / operations", "Professional services - engagement workflow coordination", "Status chased across teams", 9, "Professional services"),
    ("Core service delivery / operations", "Education - student queries and course support", "Repetitive queries", 6, "Education"),
]

# ------------------------------------------------------------------ standards basis (step, standard, how it is used)
STANDARDS = [
    ("1 Position", "OECD Framework for the Classification of AI Systems (2022)", "Economic Context and Data & Input dimensions shape the position screen: scale, sector, data provenance, structure and availability."),
    ("1 Position", "ISO/IEC 42001:2023 cl. 4.1 - 4.2", "Understanding the organisation, its context and interested parties before deciding where AI is used."),
    ("1 Position", "NIST AI RMF 1.0 - MAP 1.3, 1.4, 1.6", "Organisational mission, business value and system requirements established before AI is selected."),
    ("2 Discover", "APQC Process Classification Framework (cross-industry)", "Function / process taxonomy behind the process library and inventory."),
    ("2 Discover", "NIST AI RMF 1.0 - MAP 1.1", "Intended purpose, context of use and users documented for each candidate process."),
    ("3 Diagnose", "ISO/IEC 22989:2022 and ISO/IEC 23053:2022", "AI concepts, tasks and ML system components - basis for the task-nature to solution-family mapping."),
    ("3 Diagnose", "OECD AI classification - AI Model and Task & Output dimensions", "Distinguishes recognition, detection, forecasting, personalisation, goal-driven optimisation and generation tasks."),
    ("3 Diagnose", "ISO/IEC 5259 series (2024-2025) and ISO/IEC 8183:2023", "Data quality and data life cycle for analytics and ML - basis for the per-process data gate."),
    ("3 Diagnose", "ISO/IEC 42001:2023 Annex A.7 (Data for AI systems)", "Data acquisition, quality, provenance and preparation expected before AI is deployed."),
    ("3 Diagnose", "NSW AI Assessment Framework (AIAF)", "Inherent-risk questions (impact on people, autonomy, data sensitivity) behind the risk tier."),
    ("3 Diagnose", "EU AI Act (Reg. 2024/1689) Art. 5, 6, 50 and Annex III", "Screens out prohibited practices and flags high-risk uses (credit, employment, essential services, biometrics) and transparency duties."),
    ("3 Diagnose", "Bahrain PDPL (Law No. 30 of 2018)", "Personal and sensitive data flag; lawful basis and DPIA expectations for higher-risk processing."),
    ("3 Diagnose", "ISACA - Cybersecurity Recommendations for Securing AI Agents (2026); OWASP Top 10 for LLM Applications (2025)", "Autonomy rating and agentic guardrails: least agency, human approval for high-impact actions, excessive-agency risk."),
    ("4 Prioritise", "NIST AI RMF 1.0 - MAP 3.1, 3.2", "Expected benefits and costs (incl. of errors) weighed per use case."),
    ("4 Prioritise", "ISO/IEC 42001:2023 cl. 6.2", "AI objectives that are measurable - feed the KPIs on each canvas."),
    ("5 Plan", "ISO/IEC 42005:2025 and ISO/IEC 23894:2023", "AI system impact assessment and risk management for Medium / High-tier use cases before go-live."),
    ("5 Plan", "ISO/IEC 5338:2023 and ISO/IEC 42001 Annex A.6", "AI system life cycle from design to retirement for each roadmap item."),
    ("5 Plan", "GT AI Risk Assessment Framework v1.1 and GT AI RCM v2.1", "Tier-based minimum controls on each use-case canvas; control testing once in production."),
    ("Complementary", "GT AI Ready7 / GT AI Readiness & Risk Assessment", "Measures governance and capability maturity. Not repeated here - this toolkit asks WHERE AI fits, not HOW READY the organisation is."),
]

# ------------------------------------------------------------------ worked example - GT Bahrain (illustrative)
EXAMPLE_PROFILE = {"Organisation": "Grant Thornton Bahrain", "Sector": "Professional services", "Assessment date": "03/10/2026",
                   "Facilitated by": "GT AI Service Line", "Scope": "Whole firm - audit, tax, advisory and support functions"}
EXAMPLE_POSITION = {"D1": 2, "D2": 2, "D3": 1, "D4": 2, "D5": 2, "D6": 1, "S1": 1, "S2": 1, "S3": 1, "S4": 2, "S5": 1, "S6": 2}  # option index 0-3
# function idx, process, pain, nature idx, volume idx, hrs/month, standardised idx, data idx, strategic, customer impact, rights, personal, autonomy idx
EXAMPLE_PROCESSES = [
    (0, "Audit - journal-entry testing", "Sample-based testing of large ledgers; anomalies missed", 8, 2, 300, 1, 2, 5, 5, "No", "Yes", 1),
    (0, "Audit - working-paper and memo drafting", "Senior time spent drafting standard documentation", 4, 1, 400, 0, 2, 4, 4, "No", "Yes", 1),
    (0, "Audit - bank and confirmation letters", "Confirmations read and keyed manually", 3, 2, 120, 0, 2, 3, 3, "No", "Yes", 1),
    (0, "Tax - VAT and corporate tax research", "Research across NBR guidance, laws and prior advice", 5, 1, 160, 1, 2, 4, 4, "No", "No", 0),
    (0, "Tax - VAT return data preparation", "Collating client transaction data and VAT codes", 0, 2, 200, 0, 2, 3, 4, "No", "Yes", 1),
    (1, "Proposals and tender responses", "Re-writing similar credentials and methodology", 4, 1, 350, 1, 2, 4, 4, "No", "No", 0),
    (0, "Advisory - contract and document review (due diligence)", "Large data rooms reviewed by hand", 4, 1, 220, 1, 2, 4, 4, "No", "Yes", 1),
    (8, "Meeting minutes and action tracking", "Minutes late; actions lost", 4, 2, 250, 0, 1, 3, 3, "No", "No", 0),
    (9, "Methodology and knowledge Q&A", "Staff search across manuals and past work", 5, 2, 200, 1, 2, 4, 3, "No", "No", 0),
    (6, "Client acceptance, KYC and independence checks", "Checks collated by email and spreadsheets", 1, 1, 140, 0, 2, 4, 4, "No", "Yes", 1),
    (3, "Timesheets, WIP and billing", "Chasing timesheets; manual billing narratives", 0, 2, 180, 0, 3, 3, 3, "No", "No", 1),
    (3, "Expense claims", "Receipts checked manually", 3, 1, 60, 0, 2, 2, 2, "No", "Yes", 1),
    (4, "Recruitment - CV screening", "Peaks of applications during graduate intake", 8, 1, 80, 1, 1, 3, 3, "Yes", "Yes", 1),
    (4, "HR policy and leave queries", "HR answering repeat questions", 6, 1, 60, 0, 2, 2, 3, "No", "Yes", 0),
    (2, "Thought-leadership and social content", "Slow content production", 4, 1, 80, 0, 0, 3, 2, "No", "No", 0),
    (0, "Engagement set-up and status coordination", "Status chased across teams; steps vary by partner", 9, 1, 150, 2, 1, 4, 3, "No", "Yes", 1),
    (8, "Management KPI reporting", "Partner packs compiled by hand", 2, 0, 100, 1, 2, 4, 3, "No", "No", 1),
    (1, "Client cross-sell and retention insights", "No view of which clients need which services or are at risk of leaving", 8, 1, 200, 1, 1, 4, 4, "No", "Yes", 0),
    (0, "AI readiness assessment report production", "Report built manually per client", 0, 0, 40, 0, 2, 4, 4, "No", "No", 1),
    (6, "Client BCP and business impact analysis (AI-powered BCP tool - in preparation)", "Generic, outdated plans; slow BIA workshops", 4, 0, 200, 1, 1, 5, 4, "No", "Yes", 1),
    (0, "Financial statements preparation (AI-powered FS tool - in preparation)", "Mapping trial balances and drafting notes manually", 8, 1, 260, 0, 2, 5, 5, "No", "Yes", 1),
]
