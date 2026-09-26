# -*- coding: utf-8 -*-
"""Shared content for the GT AI Risk Assessment Framework (document + toolkit)."""

CATEGORIES = [
    ("GOV", "Governance & Accountability"),
    ("LEG", "Legal, Regulatory & Compliance"),
    ("FAI", "Fairness, Bias & Human Rights"),
    ("PRI", "Privacy & Data Protection"),
    ("DAT", "Data Quality & Integrity"),
    ("PER", "Model Performance, Reliability & Robustness"),
    ("TRA", "Transparency, Explainability & Contestability"),
    ("SEC", "Security of AI Systems"),
    ("HUM", "Human Oversight & Autonomy"),
    ("TPR", "Third-Party & Supply Chain"),
    ("OPS", "Operational, Financial & Sustainability"),
    ("SOC", "Societal, Reputational & AI-Enabled Threats"),
]

# (id, cat, risk, description / scenario, key sources, RCM controls, example KRI)
TAXONOMY = [
    ("GOV-01", "GOV", "Unclear accountability for AI outcomes", "No named business or technical owner for an AI system; decisions and incidents cannot be attributed or escalated.", "ISO/IEC 42001 5.3; NIST AI RMF Govern 2.1", "GL-1, GL-2, GL-8", "% AI systems in register without named owner"),
    ("GOV-02", "GOV", "Shadow / unsanctioned AI use", "Staff use unapproved AI tools, bypassing risk assessment and exposing confidential data.", "AI Ready7 1.2; NSW AIAF", "GL-5, GL-8", "Unsanctioned AI services detected per month"),
    ("GOV-03", "GOV", "AI deployed without adequate risk assessment", "Use case moves to production without tiering, impact assessment or approval proportionate to risk.", "ISO/IEC 23894; EU AI Act Art. 9", "RM-2, RM-5, GL-7", "% production AI systems with current assessment"),
    ("GOV-04", "GOV", "Misaligned or low-value AI investment", "AI initiatives lack defined benefits, leading to wasted spend and continued investment in ineffective or risky solutions.", "AI Ready7 1.3; NIST AI RMF Map 3.1", "GL-3, GL-6", "% AI initiatives meeting benefit KPIs"),
    ("LEG-01", "LEG", "Use of a prohibited AI practice", "AI system falls within a prohibited category (e.g., manipulative techniques, social scoring, workplace emotion recognition).", "EU AI Act Art. 5", "RM-5, RS-2", "Number of use cases failing prohibited-practice screen"),
    ("LEG-02", "LEG", "Non-compliance with high-risk AI obligations", "Provider or deployer duties (risk management, data governance, logging, oversight, FRIA, registration) not met for a high-risk system.", "EU AI Act Art. 8-27, 43, 49", "RO-1, RO-6, RM-5", "Open compliance gaps on high-risk systems"),
    ("LEG-03", "LEG", "Breach of sector or local regulation", "AI use breaches Bahrain PDPL, CBB Rulebook, iGA policy or other applicable regimes (incl. KSA/UAE where operating).", "Bahrain PDPL; CBB Rulebook; iGA AI Policy", "RO-5, PR-1, TP-3", "Regulatory findings related to AI"),
    ("LEG-04", "LEG", "Intellectual property / copyright infringement", "Training data, prompts or outputs infringe third-party IP, licence terms or client confidentiality.", "EU AI Act Art. 53(1)(c); Bahrain Copyright Law", "LC-11, LC-6", "IP complaints / takedown requests"),
    ("LEG-05", "LEG", "Liability for AI-caused harm", "Organisation is liable for damage caused by AI outputs or actions (contractual, tort or product liability).", "EU Product Liability Directive (2024/2853); Bahrain draft AI law", "TP-3, CO-3, IM-2", "Claims / complaints attributed to AI"),
    ("FAI-01", "FAI", "Discriminatory or biased outcomes", "Model produces systematically different outcomes for protected or vulnerable groups (e.g., gender, nationality, age).", "ISO/IEC TR 24027; NIST SP 1270; EU AI Act Art. 10", "RS-5, LC-1", "Disparate impact ratio across cohorts"),
    ("FAI-02", "FAI", "Unrepresentative training or evaluation data", "Data does not reflect the population or context of use (e.g., Arabic dialects, local customer segments).", "EU AI Act Art. 10(3)-(4); ISO/IEC 5259", "LC-1, RS-7", "Coverage gaps by segment in evaluation set"),
    ("FAI-03", "FAI", "Infringement of fundamental rights / dignity", "AI use undermines autonomy, dignity, freedom of expression or access to essential services.", "EU AI Act Art. 27; CoE HUDERIA; UNESCO", "RM-2, RO-6", "FRIAs outstanding for in-scope systems"),
    ("FAI-04", "FAI", "Exclusion and accessibility barriers", "Persons with disabilities, low digital literacy or non-English speakers cannot use or benefit from AI services.", "EU AI Act Art. 16(l); GCC AI Ethics Manual", "RS-7", "Accessibility defects open"),
    ("PRI-01", "PRI", "Unlawful processing of personal data", "Personal data used for AI without lawful basis, beyond original purpose, or without notice.", "Bahrain PDPL; GDPR Art. 5-6", "PR-1, PR-2", "AI systems processing personal data without DPIA"),
    ("PRI-02", "PRI", "Leakage of confidential / personal data via AI", "Sensitive data entered into prompts, memorised by models or exposed in outputs / logs.", "OWASP LLM02:2025; NIST AI 600-1 (Data privacy)", "GL-5, SE-4, PR-4", "DLP events involving AI services"),
    ("PRI-03", "PRI", "Unassessed cross-border data transfer", "AI service processes data outside Bahrain without transfer assessment or safeguards.", "Bahrain PDPL; Decree No. 56/2018", "PR-6", "AI services hosted outside approved regions"),
    ("PRI-04", "PRI", "Inability to honour data-subject rights", "Access, objection, rectification or erasure requests cannot be fulfilled for data embedded in models or vector stores.", "Bahrain PDPL; GDPR Art. 15-22", "PR-5", "DSRs involving AI exceeding SLA"),
    ("PRI-05", "PRI", "Re-identification / inference of sensitive attributes", "Model infers sensitive attributes or re-identifies individuals from anonymised data.", "ISO/IEC 27559; NIST AI 100-2 (inference attacks)", "PR-4, SE-6", "Re-identification test failures"),
    ("DAT-01", "DAT", "Poor data quality", "Incomplete, inaccurate, stale or inconsistent data degrades model outputs.", "ISO/IEC 5259; ISO/IEC 42001 A.7.4", "LC-1", "Data-quality rule failures"),
    ("DAT-02", "DAT", "Lack of data provenance and lineage", "Origin, transformations and rights of data cannot be traced, undermining defensibility.", "ISO/IEC 42001 A.7.5; AI Ready7 5.4", "LC-1, LC-11", "% datasets with documented lineage"),
    ("DAT-03", "DAT", "Data / model drift", "Changes in input data distribution or real-world relationships reduce performance over time.", "NIST AI RMF Measure 2.4; ISO/IEC 5338", "OM-1, RM-6", "Population stability index / drift score"),
    ("PER-01", "PER", "Inaccurate or unreliable outputs", "Model error rates exceed acceptable levels for the intended purpose.", "EU AI Act Art. 15; ISO/IEC 25059", "RS-6, LC-9", "Error rate vs threshold"),
    ("PER-02", "PER", "Hallucination / confabulation", "Generative AI produces plausible but false content (facts, citations, calculations).", "NIST AI 600-1 (Confabulation)", "LC-7", "Hallucination rate in evaluation / QA sampling"),
    ("PER-03", "PER", "Lack of robustness to edge cases", "Performance collapses on unusual inputs, adversarial perturbations or out-of-distribution data.", "ISO/IEC 24029; EU AI Act Art. 15(4)", "RS-3, AA-3", "Robustness test pass rate"),
    ("PER-04", "PER", "Uncontrolled model or provider changes", "Provider updates foundation model or system is changed without re-validation, altering behaviour.", "EU AI Act Art. 3(23); ISO/IEC 42001 A.6.2.5", "LC-5, LC-6", "Unvalidated model version changes"),
    ("PER-05", "PER", "Uncontrolled prompt / configuration changes", "Changes to prompts, guardrails, tool permissions or retrieval sources degrade safety or behaviour without regression testing.", "ISACA Securing AI Agents (2026)", "LC-12", "Releases without passing LLM regression suite"),
    ("TRA-01", "TRA", "Undisclosed AI interaction or content", "Users are unaware they interact with AI or that content is AI-generated.", "EU AI Act Art. 50", "CO-1, LC-7", "AI touchpoints lacking disclosure"),
    ("TRA-02", "TRA", "Unexplainable decisions", "Organisation cannot explain AI-supported decisions to affected persons, regulators or auditors.", "EU AI Act Art. 13, 86; NIST IR 8312", "RS-4, LC-4", "Explanation requests not satisfied"),
    ("TRA-03", "TRA", "No effective contest or redress route", "Affected persons cannot challenge or obtain human review of AI-supported decisions.", "EU AI Act Art. 86; NSW AIAF", "CO-3, PR-5", "Appeals exceeding SLA"),
    ("TRA-04", "TRA", "Inadequate documentation and records", "Technical documentation, logs or decisions are insufficient for audit or regulatory inspection.", "EU AI Act Art. 11-12, 18-19", "LC-4, OM-2, RO-3", "Systems with incomplete documentation pack"),
    ("SEC-01", "SEC", "Prompt injection and jailbreak", "Attacker manipulates model behaviour via direct or indirect (embedded content) instructions.", "OWASP LLM01:2025; MITRE ATLAS", "SE-6", "Blocked injection attempts / bypasses in testing"),
    ("SEC-02", "SEC", "Data or model poisoning", "Training, fine-tuning or retrieval data is manipulated to corrupt outputs or insert backdoors.", "OWASP LLM04:2025; NIST AI 100-2", "SE-6, LC-1, SE-7", "Integrity check failures"),
    ("SEC-03", "SEC", "Model theft, extraction or inversion", "Model weights, system prompts or training data are extracted.", "OWASP LLM07/LLM10:2025; MITRE ATLAS", "SE-4, SE-6", "Anomalous query volumes"),
    ("SEC-04", "SEC", "Insecure integration and excessive privilege", "AI components have excessive permissions or insecure APIs/plugins, enabling unauthorised actions.", "OWASP LLM05/LLM06:2025; AI Ready7 4.2-4.3", "SE-2, LC-8", "Agents / integrations with privileged standing access"),
    ("SEC-06", "SEC", "Agent memory poisoning & cross-session leakage", "Poisoned or unauthorised content persists in agent memory / vector stores, or context leaks across users, sessions or tenants.", "ISACA Securing AI Agents (2026); OWASP LLM08:2025", "SE-8", "Cross-tenant retrieval test failures"),
    ("SEC-07", "SEC", "Unsafe code execution, SSRF & sandbox escape", "Agent-generated code or tool calls reach internal networks, cloud metadata services or production secrets.", "ISACA Securing AI Agents (2026); CWS ISO 42001 checklist 33", "SE-9, SE-10", "Blocked egress / SSRF attempts"),
    ("SEC-05", "SEC", "Compromised AI supply chain", "Malicious or vulnerable third-party models, libraries, plugins or MCP servers.", "OWASP LLM03:2025; CISA AI Data Security", "SE-7", "Unscanned model artefacts"),
    ("HUM-01", "HUM", "Automation bias and over-reliance", "Users accept AI outputs without adequate scrutiny, eroding professional judgement.", "EU AI Act Art. 14(4)(b); NIST AI 600-1 (Human-AI configuration)", "RS-1, GL-4, LC-7", "Override / challenge rate"),
    ("HUM-02", "HUM", "Ineffective human oversight", "Oversight persons lack competence, authority, time or information to intervene.", "EU AI Act Art. 14, 26(2)", "RS-1, GL-4", "Oversight roles without training"),
    ("HUM-03", "HUM", "Unbounded autonomous agent actions", "AI agents take high-impact actions (payments, communications, code execution) without adequate limits or approval.", "OWASP Agentic Top 10; Singapore Agentic AI framework", "LC-8", "High-impact agent actions without approval"),
    ("HUM-04", "HUM", "Adverse workforce impact", "AI adoption causes skills erosion, job anxiety or unconsulted changes to working conditions.", "EU AI Act Art. 26(7); AI Ready7 3.4", "RS-9, GL-4", "Staff sentiment / attrition in affected roles"),
    ("TPR-01", "TPR", "Vendor lock-in and concentration", "Critical dependence on a single AI provider or cloud creates resilience and bargaining risk.", "ISO/IEC 42001 A.10; CBB outsourcing", "TP-2, TP-3", "% critical AI workloads on single provider"),
    ("TPR-02", "TPR", "Inadequate vendor transparency and assurance", "Vendor cannot provide documentation, testing evidence or incident information needed for compliance.", "EU AI Act Art. 25(4); ISO/IEC 42001 A.10.3", "TP-1, TP-3, LC-6", "Vendors without current assurance evidence"),
    ("TPR-03", "TPR", "Vendor use of organisation data", "Provider uses prompts, outputs or client data to train models or retains data beyond agreed terms.", "Contract / DPA; Bahrain PDPL", "TP-3, PR-6", "Contracts lacking no-training clause"),
    ("OPS-01", "OPS", "AI service outage or degradation", "Business processes dependent on AI are disrupted by service unavailability or latency.", "ISO/IEC 27001 A.5.30; CBB OM", "OM-1, SE-1", "AI service availability vs SLA"),
    ("OPS-02", "OPS", "Uncontrolled AI cost and consumption", "Compute and API spend escalate without governance (incl. denial-of-wallet attacks).", "AI Ready7 4.5; OWASP LLM10:2025", "GL-6, SE-6", "AI spend vs budget"),
    ("OPS-04", "OPS", "No safe degradation or kill switch", "A failing or compromised AI system cannot be stopped quickly or the business process cannot revert to manual operation.", "ISACA Securing AI Agents (2026); ISO 22301", "OM-4", "Kill-switch test age (days)"),
    ("OPS-03", "OPS", "Environmental impact of AI", "Energy, water and carbon footprint of AI workloads conflicts with sustainability commitments.", "ISO/IEC TR 20226; ISO/IEC 42001", "RS-8", "kgCO2e per 1,000 inferences"),
    ("SOC-01", "SOC", "Harmful, toxic or unsafe content", "AI generates offensive, dangerous or culturally inappropriate content.", "NIST AI 600-1; GCC AI Ethics Manual", "RS-2, SE-6", "Content-policy violations per 1,000 outputs"),
    ("SOC-02", "SOC", "Misinformation and loss of stakeholder trust", "AI-generated errors or undisclosed AI use damage client, regulator or public trust.", "NIST AI 600-1 (Information integrity); AI Ready7 1.4", "LC-7, CO-1, CO-2", "Client complaints linked to AI"),
    ("SOC-03", "SOC", "Deepfake and AI-enabled fraud against the organisation", "Voice/video impersonation or AI-crafted phishing leads to fraudulent payments or data release.", "FinCEN FIN-2024-Alert004; AI Ready7 7.3", "AT-3, GL-4", "Deepfake / BEC attempts detected"),
    ("SOC-04", "SOC", "AI-accelerated cyber attacks", "Adversaries use AI for automated discovery, exploit chaining and scaled social engineering, outpacing defences.", "AI Ready7 7.1-7.5; MITRE ATLAS", "AT-1, AT-2, AT-4", "Mean time to remediate critical vulnerabilities"),
]

# Inherent-risk tiering factors: (id, factor, weight, [level0..level4 descriptions])
TIER_FACTORS = [
    ("T1", "Impact of decisions on individuals", 3, [
        "No decisions about individuals (e.g., internal content drafting, code assistance)",
        "Minor internal operational decisions; negligible effect on individuals",
        "Influences decisions about individuals with limited, easily reversible effect",
        "Materially influences decisions affecting individuals' access to services, employment or finances",
        "Makes or determines decisions with legal or similarly significant effects (credit, hiring, eligibility, health, law enforcement)"]),
    ("T2", "Autonomy & human oversight", 3, [
        "AI provides information only; humans make all decisions",
        "AI recommends; a human reviews every output before use",
        "AI recommends; humans review samples or exceptions only",
        "AI decides or acts automatically; humans can intervene after the fact",
        "AI agent acts autonomously on systems / external parties with no practical real-time intervention"]),
    ("T3", "Data sensitivity", 2, [
        "Public or synthetic data only",
        "Internal, non-confidential business data",
        "Confidential business or client data",
        "Personal data",
        "Sensitive personal data (health, biometric, financial, children, criminal) or regulated/classified data"]),
    ("T4", "Scale & reach", 1, [
        "Fewer than 10 internal users; pilot",
        "Single team / department",
        "Organisation-wide internal use",
        "External customers / public (up to 10,000 people)",
        "Large-scale public or market-wide impact (> 10,000 people)"]),
    ("T5", "Reversibility of outcomes", 2, [
        "Outputs are advisory drafts with no direct effect",
        "Errors easily corrected with no lasting impact",
        "Errors correctable with moderate effort or inconvenience",
        "Errors difficult to reverse; lasting financial or reputational impact",
        "Irreversible harm possible (safety, liberty, significant financial loss)"]),
    ("T6", "Vulnerable groups affected", 2, [
        "No individuals affected",
        "Staff only, with no vulnerability indicators",
        "General public / customers",
        "Groups with some vulnerability (e.g., elderly, low digital literacy, migrant workers)",
        "Children, patients, persons in financial distress or other highly vulnerable persons"]),
    ("T7", "Regulatory classification & sector criticality", 2, [
        "Unregulated, non-critical activity",
        "General regulatory obligations only (e.g., data protection)",
        "Regulated sector activity (e.g., banking, insurance, telecom, healthcare)",
        "Specific AI transparency or sector AI obligations apply (e.g., EU AI Act Art. 50, CBB requirements)",
        "High-risk AI classification (e.g., EU AI Act Annex III) or critical infrastructure"]),
    ("T8", "Technology novelty & opacity", 1, [
        "Rules-based / deterministic automation",
        "Traditional interpretable ML (e.g., regression, decision trees)",
        "Complex ML (e.g., gradient boosting, deep learning) with explainability tooling",
        "Generative AI / LLM (third-party foundation model)",
        "Agentic or multi-model AI ecosystem, self-learning in production"]),
    ("T9", "External exposure", 1, [
        "Offline / isolated system",
        "Internal users only, internal data sources",
        "Consumes external content (web, email, documents) or external APIs",
        "Directly customer- or public-facing interface",
        "Public-facing and able to act on external systems / third parties"]),
    ("T10", "Third-party dependence", 1, [
        "Fully in-house model and infrastructure",
        "In-house model on third-party cloud",
        "Vendor model customised for the organisation",
        "Commercial SaaS AI with limited visibility of model",
        "Critical dependence on a single opaque third-party AI provider"]),
]

TIERS = [
    ("Low", 0, "Self-managed by use-case owner; lightweight assessment; register entry.", "Use-case owner", "Every 24 months or on material change"),
    ("Medium", 25, "Standard risk assessment; SME review (privacy, security, data); documented controls.", "Head of function + Risk (2nd line) review", "Every 12 months or on material change"),
    ("High", 50, "Full risk & impact assessment (ISO/IEC 42005 / FRIA where required); independent validation; AI Committee approval.", "AI Committee / CoE", "Every 6 months + continuous KRI monitoring"),
    ("Critical", 75, "As High plus legal review, executive sign-off, board notification, pre-deployment red teaming and staged rollout.", "Executive Committee / Board Risk Committee", "Quarterly + continuous KRI monitoring"),
]

LIKELIHOOD = [
    (1, "Rare", "May occur only in exceptional circumstances; < 5% chance within 12 months; no known occurrences in similar systems."),
    (2, "Unlikely", "Could occur but not expected; 5-20% within 12 months; isolated occurrences in industry."),
    (3, "Possible", "Might occur; 20-50% within 12 months; has occurred in comparable organisations."),
    (4, "Likely", "Will probably occur; 50-80% within 12 months; has occurred previously in the organisation."),
    (5, "Almost certain", "Expected to occur; > 80% within 12 months or recurring."),
]

# impact scale per dimension (level: [individuals/rights, financial, regulatory/legal, operational, reputational, safety/societal/environment])
IMPACT_DIMS = ["Individuals & rights", "Financial", "Regulatory & legal", "Operational", "Reputational & client trust", "Safety, societal & environmental"]
IMPACT = [
    (1, "Insignificant", ["No discernible effect on individuals", "< BHD 5,000", "No breach; no regulatory interest", "Negligible disruption (< 1 hour)", "No external awareness", "No harm"]),
    (2, "Minor", ["Minor inconvenience to a few individuals, easily remedied", "BHD 5,000 - 50,000", "Minor breach; self-reported; no sanction", "Minor disruption (< 1 day); workaround available", "Limited local / client awareness", "Minor, temporary, contained harm"]),
    (3, "Moderate", ["Adverse effect on a group; rights affected but remediable", "BHD 50,000 - 250,000", "Breach resulting in regulatory enquiry or warning", "Disruption of key process (1-3 days)", "Negative media / client complaints; short-term", "Moderate harm; limited environmental impact"]),
    (4, "Major", ["Significant harm or discrimination affecting many individuals", "BHD 250,000 - 1,000,000", "Regulatory sanction, fine or enforcement action", "Critical service outage (> 3 days)", "Sustained national media coverage; loss of key clients", "Serious injury / significant societal harm"]),
    (5, "Severe", ["Irreversible harm to individuals' rights, safety or livelihood at scale", "> BHD 1,000,000", "Licence suspension / criminal liability / prohibited practice", "Prolonged failure of critical operations", "Long-term, international reputational damage", "Loss of life / widespread irreversible harm"]),
]

CONTROL_EFFECTIVENESS = [
    ("Strong", 0.4, "Controls designed and operating effectively; tested with no material exceptions; automated where appropriate."),
    ("Satisfactory", 0.6, "Controls largely designed and operating; minor exceptions; improvement opportunities identified."),
    ("Needs improvement", 0.8, "Controls partly designed or inconsistently operated; material exceptions; reliance limited."),
    ("Weak / None", 1.0, "Controls absent, not designed to address the risk, or not operating."),
]

RATING_BANDS = [("Low", 1, 4, "Accept; monitor through normal governance."),
                ("Medium", 5, 9, "Treat within normal planning cycle; owner accountable; monitor KRIs."),
                ("High", 10, 16, "Treatment plan within 90 days; AI Committee oversight; consider restricting use."),
                ("Critical", 17, 25, "Outside appetite: immediate action; escalate to Executive / Board; suspend or do not deploy until reduced.")]
