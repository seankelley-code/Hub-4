'use strict';
process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';

const express           = require('express');
const Anthropic         = require('@anthropic-ai/sdk');
const PptxGenJS         = require('pptxgenjs');
const buildEditablePptx = require('./buildPresentation');
const fs                = require('fs');
const path              = require('path');

const CLASP_RC = path.join(process.env.USERPROFILE || process.env.HOME || '', '.clasprc.json');
function getClaspCredentials() {
  try {
    const rc = JSON.parse(fs.readFileSync(CLASP_RC, 'utf8'));
    if (rc.tokens?.default?.access_token) {
      return { clientId: rc.oauth2ClientSettings?.clientId, clientSecret: rc.oauth2ClientSettings?.clientSecret, token: rc.tokens.default };
    }
    if (rc.token?.access_token) {
      return { clientId: rc.oauth2ClientSettings?.clientId, clientSecret: rc.oauth2ClientSettings?.clientSecret, token: rc.token };
    }
  } catch {}
  return null;
}

require('dotenv').config({
  path: path.join(process.env.USERPROFILE || '', 'Desktop', 'audit-app', '.env'),
});

const app    = express();
const client = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY });
const PORT   = 3003;

app.use(express.json({ limit: '10mb' }));
app.use(express.static(path.join(__dirname, 'public')));

const AUDITS_FILE = path.join(__dirname, 'data', 'audits.json');

function readAudits() {
  try { return JSON.parse(fs.readFileSync(AUDITS_FILE, 'utf8')); }
  catch { return []; }
}

function writeAudits(data) {
  fs.writeFileSync(AUDITS_FILE, JSON.stringify(data, null, 2));
}

// ── CRUD ──────────────────────────────────────────────────────────────────────

app.get('/api/audits', (req, res) => {
  const audits = readAudits().map(a => ({
    id: a.id, title: a.title, currentPhase: a.currentPhase,
    entity: a.discovery?.['disc-entity'] || '',
    auditType: a.discovery?.['disc-audit-type'] || '',
    createdAt: a.createdAt, updatedAt: a.updatedAt,
    chatCount: (a.chatHistory || []).length,
  }));
  res.json(audits);
});

app.post('/api/audits', (req, res) => {
  const audits = readAudits();
  const audit = {
    id: `audit-${Date.now()}`,
    title: req.body.title || 'New Audit',
    status: 'active',
    createdAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    currentPhase: 'discovery',
    chatHistory: [],
    discovery: {
      'disc-title': '', 'disc-entity': '', 'disc-business-unit': '',
      'disc-audit-type': '', 'disc-risk-category': '',
      'disc-period-start': '', 'disc-period-end': '',
      'disc-lead-auditor': '', 'disc-manager': '',
      'disc-team-members': '', 'disc-stakeholders': '',
      'disc-background': '', 'disc-risk-areas': '', 'disc-scope': '',
    },
    planning: {
      'plan-objectives': '', 'plan-scope': '', 'plan-out-of-scope': '',
      'plan-testing-approach': '', 'plan-key-systems': '',
      'plan-resource-hours': '', 'plan-notes': '', 'plan-risk-register': [],
    },
    fieldwork: {
      'fw-evidence': '', 'fw-tech-areas': '', 'fw-issues': '',
      'fw-contacts': '', 'fw-status': '', 'fw-worksteps': [],
    },
    reporting: {
      'rep-exec-summary': '', 'rep-opinion': '', 'rep-distribution': '',
      'rep-notes': '', 'rep-vetting': '', 'rep-issues-log': [],
    },
    wrapup: {
      'wu-final-opinion': '', 'wu-mgmt-responses': '',
      'wu-retrospective': '', 'wu-lessons': '', 'wu-maps': [],
      'wu-archive-checklist': {
        'Audit report issued and distributed': false,
        'Management responses received': false,
        'All MAPs documented in system': false,
        'Evidence workpapers archived': false,
        'Engagement closed in system': false,
        'Post-audit retrospective completed': false,
        'Post-audit survey sent': false,
      },
    },
    followup: {
      'fu-map-updates': '', 'fu-remediation': '', 'fu-escalations': '',
      'fu-conclusion': '', 'fu-open-issues': [],
    },
  };
  audits.push(audit);
  writeAudits(audits);
  res.json(audit);
});

app.get('/api/audits/download', (req, res) => {
  res.setHeader('Content-Disposition', 'attachment; filename="via-audits.json"');
  res.setHeader('Content-Type', 'application/json');
  res.sendFile(AUDITS_FILE);
});

app.get('/api/audits/:id', (req, res) => {
  const audit = readAudits().find(a => a.id === req.params.id);
  if (!audit) return res.status(404).json({ error: 'Not found' });
  res.json(audit);
});

app.put('/api/audits/:id', (req, res) => {
  const audits = readAudits();
  const idx = audits.findIndex(a => a.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'Not found' });
  audits[idx] = { ...audits[idx], ...req.body, id: req.params.id, updatedAt: new Date().toISOString() };
  writeAudits(audits);
  res.json(audits[idx]);
});

app.delete('/api/audits/:id', (req, res) => {
  const audits = readAudits().filter(a => a.id !== req.params.id);
  writeAudits(audits);
  res.json({ ok: true });
});

app.delete('/api/audits/:id/chat', (req, res) => {
  const audits = readAudits();
  const idx = audits.findIndex(a => a.id === req.params.id);
  if (idx !== -1) { audits[idx].chatHistory = []; writeAudits(audits); }
  res.json({ ok: true });
});

// ── AI System Prompt ──────────────────────────────────────────────────────────

const PHASE_LABELS = {
  discovery: 'Discovery', planning: 'Planning', fieldwork: 'Fieldwork',
  reporting: 'Reporting', wrapup: 'Wrap-Up', followup: 'Follow-Up',
};
const PHASES = ['discovery', 'planning', 'fieldwork', 'reporting', 'wrapup', 'followup'];

const VIA_DOMAIN_KNOWLEDGE = `
════════════════════════════════════════
VERIZON INTERNAL AUDIT (VIA) — METHODOLOGY REFERENCE
════════════════════════════════════════

ORGANIZATION
- VIA reports directly to the Audit Committee. SVP-IA confirms independence to Board annually.
- Teams: Operations (BU audits), FAAST (Forensic & Analytics Advisory Services — data analytics), Technology (IT/network/security), PPG (Professional Practices Group — methodology, QA, reporting)
- System of record: Optro (worksteps, issues, MAPs, sign-offs, audit opinions, UIDs)
- All workpapers stored in VIA Current Google Drive; PPG transfers to Historic Drive 10 days after issuance.
- AI usage policy: Do NOT input PII, Highly Confidential data, or sensitive info into AI tools. Verify all AI-generated output independently before using.

AUDIT LIFECYCLE — 5 PHASES + FOLLOW-UP
Phase 1 — Discovery
  • Announcement Email: Notify KPOs (Key Process Owners) using bottom-up approach. AD or Director sends via standard template. ACP audits use Engagement Letter instead.
  • Internal Kickoff Meeting: Full team (FinOps, FAAST, Tech, SME). Review objective, rationale, KPOs, systems/data, audit history, preliminary test plan, data needs, access timeline.
  • SME Engagement (optional): Business SME taskforce for targeted scoping. Risk Register is internal-only — NEVER share with SMEs.
  • Evidence of Conformance: Link announcement email in Optro; Senior Manager or AD sign-off as "Reviewed."

Phase 2 — Planning
  • Walkthroughs: Document end-to-end process; evaluate control design. Perform "test of one" per control. If design gap found → STOP testing that control, record as "missing control" in Risk Register.
  • Risk Register: Document ALL in-scope risks, risk levels, controls, test procedures. Include IIA Topical Requirements mapping. Google template → export to Optro. Sr. Director sign-off required. Updates during fieldwork marked "UPDATE" with explanation.
  • Audit Opening Meeting: Share in-scope/out-of-scope, rationale, objective, and design gaps with KPOs. Take attendance.
  • QA Confirmation (Disc & Planning): AD or Director completes before Audit Opening Meeting.

Phase 3 — Fieldwork & Vetting
  • Fieldwork Testing: Conclude on operating effectiveness of controls. Documentation must allow third-party re-performance. Minimum evidence: one passing example + all exceptions. Source docs linked in workpapers, saved to Google Drive. Redact all sensitive data.
  • Data Validation (C&A): REQUIRED before sampling. Reconcile count/dollar to system of record. Non-FAAST: Data Pull & Report Generation tab. FAAST: FAAST Masterfile. IPE (Info Provided by Entity) requires C&A validation.
  • Sampling — Test of Controls sample sizes:
    Daily/frequent: 25 (standard) / 10 (lower risk)
    Weekly: 5 / 2 | Monthly: 2 / 1 | Quarterly: 2 / 1 | Annual: 1 / 1
    Population 50–250: use 10% of occurrences. Population <50: use 5 items (or 100% if <5).
  • Sampling — Accept-Reject (populations ≥200):
    High evidence: 55 items | Moderate: 30 items | Low: 16 items
    Populations <200: 100–199→10, 50–99→5, 20–49→3, 0–19→2
  • Statistical sampling parameters: 90% confidence, ±10% margin, 90% response distribution. Use Raosoft.
  • Evidence Level by Risk + Reliance: High risk + low reliance → Moderate-High. High risk + high reliance → Low-Moderate. Low risk + any reliance → Low.
  • Communication of Audit Status: Keep KPOs informed throughout. Standard engagement status deck for formal meetings.
  • Vetting & Approvals: Vet all PIs and IOs with all impacted stakeholders (not just primary ones). MAP Owner AND Executive Owner approval required before issuance. PI: BU CFO must be notified.
  • Writing Observations: Factual, verifiable, specific/measurable. Include criteria vs. condition, risk/impact, and observation theme (root cause, up to 3 per observation). FROST describes the issue; observation theme describes WHY it happened.
  • MAPs: Specific remediation approach + due date (Target vs. Estimated) + Owner + Executive Owner.
    Target Date: confirmed timeline. Estimated Target Date: when dependent on another MAP.
  • Action Plan Owners: PI/IO Executive Owner = VP or above (NOT SVP). Owner = direct report of Executive Owner.
  • Fieldwork Closing Meeting: Present audit opinion, review draft report/closing deck, get MAP Owner agreement. Take attendance.
  • QA Confirmation (Fieldwork): AD or Director completes before issuing report.

Phase 4 — Reporting
  • Audit Report Quality Assurance:
    (1) Team self-review + Self-Review Checklist
    (2) Peer review — mandatory for all. PI or Inadequate Opinion: requires TWO peer reviews (Director first, then Sr. Director). ACP without PI: AD on Sr. Director's team. All others: AD or above.
    (3) SVP review — required; provide 4 business days before issuance; SVP APPROVAL required for PI/Inadequate reports.
    (4) Final issuance review — formatting corrections only, no substance changes post-SVP review.
  • Final Report Review: Share draft with Owners and Executive Owners. PI/Inadequate: SVP of Executive Owner must receive draft 48 hours before issuance.
  • Report Issuance: File name format "20XX.XXXX [Audit Title]". All docs signed off, issues in Optro, all PIs have SOR action plan. End Date = Final Report Date.
  • QA Confirmation (Reporting): AD or Director completes before issuing report.

Phase 5 — Wrap-Up
  • Post-Engagement Appraisal (PEA): For each staff and manager within 10 business days of report issuance. Stored in VIA Assoc Dir Google Drive. NOT linked in Optro.
  • Post-Audit Retrospective: Within 10 days of report issuance. Team covers: what worked, what didn't, adjustments, timeline, scope, lessons for IA team, AI tool use effectiveness. Recommend 1–3 actions. AD logs in Google Form.
  • QA Confirmation (Wrap-Up): Finalized QA Checklist linked in Optro.

Follow-Up — PIFU (Priority Issue Follow-Up)
  • For all PIs closed after report issuance. 6–9 months after SOR received.
  • Simplified, Standard, or Combined approach based on complexity, changes, team continuity, time elapsed, testing depth.
  • No observations → email (not formal report). Observations found → formal report.

ISSUE TYPES
• Priority Issue (PI): Significant control gap with major adverse impact, regulatory violations, repeat issues, fraud, or systemic data security implications. Requires MAPs + SOR + PIFU. CAE/AVP/SDs align before reporting. BU CFO must be notified.
• Important Observation (IO): Moderate control gap; does not meet PI threshold. Requires MAPs. Included in formal Audit Report.
• Recommended Improvement (RI): Minor gap or enhancement; low residual risk. No MAPs, NOT in report. Communicated verbally or via email. Used for all non-assurance engagements.

AUDIT OPINIONS (non-ACP)
• Effective: No reportable observations; mature controls.
• Opportunities for Improvement: ≥1 reportable observation; aggregate low risk; typically <7 reported observations; maturing controls.
• Needs Significant Improvement: 1+ PI (lower risk) or multiple IOs; high aggregate risk; typically 7+ reported observations; complex MAPs. Reviewed by SD and above.
• Inadequate: Highest risk; repeat high-risk observations; immature high-risk processes; significant scope limitations. SVP approval required before issuance.

PIFU Opinions: Resolved / Resolved with IO / Resolved with PI / Not Resolved / No Opinion

FROST RISK FRAMEWORK (Entity Risk Scoring)
F = Financial (impact if something goes wrong): <$5M=1, $5–30M=2, $31–65M=3, $66–100M=4, >$100M=5
R = Regulatory/Legal: complexity of regulations and risk of fines/lawsuits (1–5)
O = Operational: complexity, process maturity, instability from changes (1–5)
S = Strategic/Reputational: negative media, customer trust, velocity of damage (1–5)
T = Technology: from minimal IT dependence (1) to critical cyber/access control failures (5)
Formula: (F+R+O+S+T) × Mitigating Control Factor = Residual Risk Score
Control Environment Score 1–5: each 1-point improvement reduces Inherent Risk by 20%.

ENGAGEMENT TYPES
• Major Process Audit: End-to-end, comprehensive, formal report, all issue types.
• Focused Audit: Targeted scope (specific risk/control/regulation), formal report.
• Limited Assurance Review: Inquiry + analytics + limited testing; directional insight; report or memo; typically IOs (not PIs).
• PIFU: Follow-up on resolved PIs.
• Advisory Review: Non-assurance; Advisory Memo; RIs only; 12-month cool-off before assurance in same space.
• Special Project: Flexible, non-standard; no IIA standards; 12-month cool-off.

ROLES & HIERARCHY
CAE → SVP-IA → Senior Director → Associate Director/Director → Manager/Team Lead → Senior Manager → Staff
• MAP Executive Owner: VP or above (NOT SVP)
• PI reports: must be addressed to VP or above in the "To:" field
• PI notification: BU CFO must be informed
• PI/Inadequate report: SVP approval required before issuance
• Risk reevaluation/MAP closure: Sr. Director approval (PI closures: SVP)

KEY TOOLS
• Optro: Audit management system — worksteps, issues, MAPs, sign-offs, QA confirmations, UIDs
• Google Drive: Official workpaper storage
• Risk360 Exec: Bi-annual risk survey (April mid-year, September year-end) from VPs+; informs audit plan
• FAAST Masterfile: FAAST data extraction and validation docs
• Raosoft: Statistical sample size calculator

OBSERVATION THEMES (Root Cause Categories — up to 3 per observation)
These are the ONLY official VIA observation themes used in Optro:
1. Control design gap — a required control does not exist or was not properly designed
2. Insufficient training and awareness — personnel lacked knowledge or skills to perform the control
3. Inadequate documentation and record keeping — records, evidence, or audit trails are missing or incomplete
4. Inadequate governance and oversight — oversight structure, escalation, or accountability is insufficient
5. Ineffective communication and coordination — information was not shared between the right people/teams
6. System gaps — the technology or system did not support or enforce the required control
7. Human error — control existed but was not followed correctly due to individual mistake
8. Insufficient resource allocation and prioritization — insufficient people, time, or funding to execute controls
9. Other — does not fit above categories
10. Not Applicable
Usage rule: Theme describes WHY the issue happened (root cause). FROST describes WHAT risk domain it affects.

FROST RISK SUB-CATEGORIES (101 official categories for Risk Register and Issues)
Financial (F):
  F - Billing & Collections | F - Budgeting | F - Cash Management | F - Compensation | F - Credit
  F - Debt | F - Employee Expenses | F - Financial Reporting | F - Fourth Party Risk (Financial)
  F - Fraud | F - Liquidity | F - Market Disruption | F - Market Risk | F - Other Financial Risks
  F - Pension/Benefit Obligation | F - Purchasing & Payment | F - Segregation of Duties (Financial)
  F - Stock | F - Tax | F - Training (Financial)

Regulatory & Legal (R):
  R - Business Compliance | R - Certifications | R - Charity | R - Environmental Health and Safety (EH&S)
  R - Ethics | R - Fourth Party Risk (Regulatory & Legal) | R - Labor and Employment | R - Legal Management
  R - Licenses and Permits | R - Miscellaneous Laws and Regulations | R - Privacy | R - Sensitive Data
  R - Segregation of Duties (Regulatory & Legal) | R - Tax | R - Telecommunications Compliance
  R - Training (Regulatory & Legal)

Operational (O):
  O - Business Continuity & Disaster Recovery (BCDR) | O - Construction | O - Content | O - Customer Service
  O - Fourth Party Risk (Operational) | O - Insurance | O - Inventory | O - Key Performance Indicators (KPIs)
  O - Marketing & Advertising | O - Other Operational Risks | O - People Management | O - Performance
  O - Physical Security | O - Process Management & Execution | O - Quality Control | O - Real Estate
  O - Safety | O - Sales | O - Security | O - Segregation of Duties (Operational) | O - Supplier Management
  O - Training (Operational) | O - Transportation | O - Transportation/Logistics

Strategic & Reputational (S):
  S - Assurance Functions | S - Corporate Culture | S - Corporate Governance | S - Crisis Management
  S - Customer Experience | S - Divestitures | S - Economy | S - External Competition | S - Goals & Priorities
  S - Government & Politics | S - Investor Relations | S - Leadership Style | S - Mergers & Acquisition
  S - Organization Structure | S - Other Strategic & Reputational Risks | S - Outsourcing | S - Partnerships
  S - Pricing | S - Product/Service Line | S - Publicity | S - Research & Development | S - Social Media
  S - Spectrum | S - Training (Strategic & Reputational)

Technology & Information Security (T):
  T - Access Control (Logical) | T - Artificial Intelligence/Machine Learning | T - Asset Management (Hardware, Software & Network Infrastructure)
  T - Availability | T - Change Management (IT) | T - Data Management | T - Fourth Party Risk (Technology & Information Security)
  T - Network Infrastructure | T - Network Security | T - Network/System Infrastructure
  T - Other Technology & Information Security Risks | T - Security Incident Response | T - Software Development Management
  T - Technology and Information Security Policy and Procedures | T - Technology and Information Security Processes and Procedures
  T - Training (Technology & Information Security) | T - Vulnerability Management

OPTRO ISSUE FIELDS (required when documenting findings in the system)
Core fields: Issue Title | Observation Rating (PI/IO/RI) | Observation Theme(s) (up to 3) | FROST Risk Sub-category
Narrative fields: Risk & Impact | Background/Overview | Criteria | Condition | Cause | Effect | Recommendation
System fields: System(s) Impacted | BU Presidents Direct Report | Associate Director | Senior Director
MAP fields: MAP Owner | MAP Executive Owner (VP+) | Current Target Due Date | Status of MAP | MAP narrative
Tracking: Closed By | Closed On | Status | Restriction Level | Audit UID

CONTROL TYPES (used in Risk Register worksteps)
Preventive | Detective | Corrective | Manual | Automated | Compensating | Application Control

RISK LIKELIHOOD AND IMPACT LEVELS
For Risk Register: High / Medium / Low (assessed as combination of likelihood × impact)
Residual Risk = Inherent Risk × Mitigating Control Factor (based on FROST scoring)

GLOSSARY
VIA=Verizon Internal Audit | FAAST=Forensic & Analytics Advisory Services | PPG=Professional Practices Group | CAE=Chief Audit Executive | KPO=Key Process Owner | ACP=Attorney-Client Privilege | PIFU=Priority Issue Follow-Up | MAP=Management Action Plan | SOR=Statement of Resolution | PI=Priority Issue | IO=Important Observation | RI=Recommended Improvement | FROST=risk framework | IPE=Information Provided by Entity | GSAM=Global Sensitivity Account Management | QAIP=Quality Assurance & Improvement Program | VMAC=Verizon Management Audit Committee | UID=Unique ID (format: 202X:0XXX; PIFUs: 202X:05XX)
`;

function getPhaseGuidance(phase) {
  return {
    discovery: `VIA Discovery phase goals: (1) Send Announcement Email to KPOs using bottom-up approach — identify KPOs first, then executive owners. AD or Director sends via standard template. (2) Schedule Internal Kickoff Meeting with full team (FinOps, FAAST, Technology, SME) to cover: audit objective and rationale, potential KPOs, systems/data, audit history, preliminary test plan, data needs, system access timeline. (3) Optionally engage an SME — Risk Register is internal only, never share with SME. Help the auditor: identify FROST risks for the entity type, draft background narrative, define scope, identify stakeholders, and set up the Internal Kickoff Meeting agenda.`,

    planning:  `VIA Planning phase goals: (1) Perform Walkthroughs — document end-to-end process, evaluate control design, perform "test of one." If design gap found: STOP testing that control, record as "missing control." (2) Complete Risk Register — all in-scope risks, risk levels (using FROST), controls, and planned test procedures. Include IIA Topical Requirements mapping. (3) Complete QA Confirmation (Discovery & Planning) before Audit Opening Meeting. (4) Hold Audit Opening Meeting with KPOs — share in-scope/out-of-scope areas, rationale, objective, design gaps. Help the auditor: write SMART audit objectives, design test procedures for each risk in the risk register, assess the control environment, estimate resource hours, and draft opening meeting materials.`,

    fieldwork: `VIA Fieldwork & Vetting phase goals: (1) Execute test procedures — document approach (time period, sample parameters, methodology, data source, C&A validation, sample size, sampling expansion). (2) Data Validation (C&A) is REQUIRED before sampling — reconcile population to system of record. (3) Apply correct sampling: Test of Controls (25/10/5/2/1 by frequency), Accept-Reject (55/30/16 by evidence level for pop ≥200). (4) Communicate audit status to KPOs throughout. (5) Document exceptions with root cause and impact. (6) Vet findings with ALL impacted stakeholders — no surprises. MAP Owner + Executive Owner (VP+) approval required. PI: notify BU CFO. (7) Write observations with: criteria, condition, cause, effect, risk/impact, recommendation, and observation theme (root cause: why did this happen?). When assigning observation themes, choose from the 10 official VIA themes in the knowledge base — theme describes WHY, FROST sub-category describes WHAT domain. When assigning FROST sub-categories, use the exact names from the 101 official sub-categories listed. Help the auditor: document test steps, calculate correct sample sizes, draft finding narratives in VIA format, write MAPs with correct ownership levels (VP+ Executive Owner, direct report Owner).`,

    reporting: `VIA Reporting phase goals: (1) Audit Report Quality Assurance — Self-Review Checklist, then peer review (PI/Inadequate: Director first then Sr. Director), then SVP review (SVP APPROVAL required for PI/Inadequate), then final issuance review (formatting only). (2) Final Report Review with Owners and Executive Owners — PI/Inadequate: SVP of Executive Owner must receive draft 48 hours before issuance. (3) Complete Audit Report Distribution List and draft issuance email. (4) Issue report — file name format "20XX.XXXX [Audit Title]." All docs signed off, issues and action plans in Optro. (5) Select correct audit opinion: Effective (no observations), Opportunities for Improvement (<7 observations, low aggregate risk), Needs Significant Improvement (7+ or high risk), Inadequate (repeat high-risk, immature processes — SVP approval required). Help the auditor: draft executive summary, write finding narratives, select the correct opinion, prepare distribution list, write the issuance email.`,

    wrapup:    `VIA Wrap-Up phase goals: (1) Post-Engagement Appraisals (PEAs) for each staff member and manager within 10 business days of issuance — stored in VIA Assoc Dir Google Drive, NOT linked in Optro. (2) Post-Audit Retrospective within 10 days of issuance — what worked, what didn't, adjustments, timeline/scope performance, lessons for IA team, AI tool utilization effectiveness. Recommend 1–3 actions. AD logs in Google Form. (3) Finalize QA Checklist for Wrap-Up and link in Optro. Help the auditor: facilitate retrospective discussion, draft PEA talking points, capture lessons learned, complete archive checklist, ensure MAPs are properly documented in Optro.`,

    followup:  `VIA PIFU (Priority Issue Follow-Up) goals: Verify management resolved the Priority Issue exposure through the original action plan or approved alternative. Performed 6–9 months after signed SOR (Statement of Resolution). (1) Before PIFU starts: determine approach — Simplified (no significant process/control changes), Standard (complex, significant changes, broad stakeholders), or Combined (multiple PIFUs merged). (2) Perform walkthroughs (roll-forward or new) + test of one for each control. (3) If resolved with no new observations: issue email (not formal report). If observations found: issue formal audit report. (4) Track MAP remediation — reach out 30 days before due dates, obtain sufficient evidence (not just confirmation from business). Help the auditor: determine correct PIFU approach, draft close-out email or formal report, assess whether evidence is sufficient to close MAPs, draft SOR request, handle escalation language for overdue items.`,
  }[phase] || 'Provide general VIA audit guidance following IIA standards.';
}

function getTabGuidance(tab) {
  return {
    pre: `You are helping with the PRE-DISCOVERY tab. Your job is to draft the announcement email and fill it directly using fill_fields.
FIELDS TO FILL:
  email-subject  — professional subject line: "VIA Audit Notification — [Entity] [Audit Type]"
  email-body     — full announcement email body (Dear [KPO], VIA is commencing [Audit Name] covering [period]. Objective: [brief]. Your cooperation requested. Contact [Lead].)
  email-sent-date — date sent (YYYY-MM-DD)
Also use fill_fields for: audit-title-input (audit title), disc-lead, disc-manager, disc-ad, disc-sd if the user provides team info.`,

    disc: `You are helping with the DISCOVERY tab. Use fill_fields to populate discovery fields. ALWAYS call fill_fields — never just describe what should go in a field.
KEY FIELDS:
  disc-background    — Business Overview: 2–3 paragraphs on what the entity does, how it fits in Verizon, recent changes
  disc-rationale     — Why selected: Risk360 signal, regulatory trigger, management request, repeat finding, data anomaly
  disc-prior-history — Prior audits: title, year, opinion, open MAPs, repeat findings
  disc-systems       — Key systems: SAP, Salesforce, billing systems, custom tools
  disc-data-needs    — Data needs: what extracts are needed, source systems, volumes
  disc-scope-in      — In-scope: specific processes, systems, time period
  disc-scope-out     — Out of scope: explicit exclusions
  disc-notes         — Open questions and follow-up items
SETUP FIELDS (fill when user provides info):
  audit-title-input, disc-entity, disc-bu, disc-audit-type, disc-frost, disc-period-start, disc-period-end, disc-target-report, disc-uid, disc-plan-year
TEAM FIELDS: disc-lead, disc-manager, disc-ad, disc-sd, disc-faast, disc-tech, disc-team
DATES: disc-announce-date, disc-kickoff-date
When user says "fill in background" or "complete discovery" — call fill_fields with all relevant fields at once.`,

    wt: `You are helping with the WALKTHROUGHS tab. Use fill_walkthroughs to add Walkthrough Log rows. Use fill_fields for the text fields below.
WALKTHROUGH LOG ROWS (call fill_walkthroughs):
  Each row = one process/control area. Fields: process, owner, systems, date, design (Effective|Partial Gap|Gap|Not Tested), gap (Yes|No).
  If design = Gap or Partial Gap → set gap = Yes. Stop testing that control and flag for Risk Register.
TEXT FIELDS (call fill_fields):
  plan-wt-notes     — Overall walkthrough observations and themes
  plan-sme-engaged  — Yes | No | In Progress
  plan-sme-name     — SME contact name(s)
  plan-sme-date     — SME engagement date (YYYY-MM-DD)
  plan-sme-scope    — What the SME was engaged to support`,

    rr: `You are helping with the RISK REGISTER tab. Use fill_risk_register to add risk rows (Risk IDs like A.R1, controls A.R1.C1, FROST sub-categories, likelihood/impact). Use fill_fields for sign-off and notes fields:
  plan-rr-signoff-name  — Sr. Director name
  plan-rr-signoff-date  — Sign-off date (YYYY-MM-DD)
  plan-rr-signoff-notes — Sign-off notes
  plan-rr-notes         — General Risk Register notes
Risk Register is INTERNAL ONLY — never share with KPOs.`,

    dv: `You are helping with the DATA VALIDATION (C&A) tab. Two fills needed:
1. META FIELDS — call fill_fields: dv-who-generated (who pulled the data), dv-date-generated (YYYY-MM-DD), dv-source (source system name), dv-report-type (select), dv-dataset-link (URL), dv-code-text (SQL/code used)
2. TABLE ROWS — call fill_data_validation: one row per data extract. Approach options: Profiling Data | Front End Confirmation | Reconciliations | Reasonableness Check | Data Owner Backend Walkthrough | Independent Generation | Control Totals | IPE Procedures
When user asks to fill or generate DV, do BOTH.`,

    drl: `You are helping with the DOCUMENT REQUEST LIST (DRL) tab. Call fill_drl to populate the table. One row per document/artifact needed. Cover: Walkthrough, Occurrence testing, Content Review, Recalibration. Name specific documents — never generic placeholders.`,

    om: `You are helping with the OPENING MEETING tab. Use fill_fields to populate ALL meeting fields:
  plan-qa-name        — AD/Director confirming QA
  plan-qa-date        — QA confirmation date (YYYY-MM-DD)
  plan-qa-notes       — QA notes
  plan-om-date        — Meeting date (YYYY-MM-DD)
  plan-om-time        — Meeting time (HH:MM)
  plan-om-location    — Room / virtual link
  plan-om-scope-in    — In-scope areas to present
  plan-om-scope-out   — Out-of-scope areas to present
  plan-om-design-gaps — Design gaps to discuss
  plan-om-notes       — Meeting notes and action items
Opening meeting shares scope with KPOs — does NOT share Risk Register.`,

    ap: `You are helping with the AUDIT PROGRAM (AuditBoard) section in Fieldwork. Call fill_audit_programs to generate one card per DRL test step. Each card: name (A.R1.C1 — Title), risk, control, controlFreq, timePeriod, testObjective, auditProcedure (3–5 steps), workPerformed (templates), testingAttributes (checkpoints). Use mode "replace".`,

    ds: `You are helping with the DATA SAMPLING tab.
- User asks to fill/complete the CURRENT OPEN card → call fill_sampling_card (fills the card already on screen in-place)
- User asks to generate ALL sampling from Risk Register → call fill_data_sampling (creates all cards, one per in-scope control)
fill_sampling_card fields: name (A.R1.C1 — Title), risk, control, freq, period, objective, auditProcedure (3–5 steps, step 1 = pull population listing), workPerformed (3–5 past-tense templates), testingAttributes (3–5 binary checkpoints).
VIA table: 1–249 → 25 samples; 250–999 → 40; 1,000+ → 60.`,

    fw: `You are helping with the FIELDWORK phase.
WORKSTEP LOG ROWS — call fill_worksteps: stepId (A.R1.C1), control, procedures, sampleSize, evidence, status (Not Started|In Progress|Complete|Exception Found|Skipped)
SCALAR FIELDS — call fill_fields: fw-status (progress narrative), fw-tech-areas, fw-contacts, fw-issues, fw-evidence
AUDIT PROGRAM — call fill_audit_programs. DRL — call fill_drl.`,

    rep: `You are helping with the REPORTING phase.
ISSUES/FINDINGS LOG — call fill_issues: id (PI-01/IO-01/RI-01), type (PI|IO|RI), title, recommendation, riskLevel (High|Medium|Low), mgmtResponse, status (Open|In Draft|Agreed|Disputed|Closed)
SCALAR FIELDS — call fill_fields: rep-exec-summary (Executive Summary), rep-opinion (Effective|Opportunities for Improvement|Needs Significant Improvement|Inadequate), rep-notes, rep-vetting, rep-distribution
When drafting the report, fill BOTH the issues rows AND the scalar fields.`,

    wu: `You are helping with the WRAP-UP phase.
MAPs TABLE — call fill_maps: mapId (MAP-01), issue, action, owner, dueDate (YYYY-MM-DD), status (Open|In Progress|Completed|Overdue|Escalated), notes
SCALAR FIELDS — call fill_fields: wu-final-opinion, wu-mgmt-responses, wu-retrospective, wu-lessons`,

    fu: `You are helping with the FOLLOW-UP phase.
TRACKING TABLE — call fill_fu_issues: mapId (MAP-01), issue, owner, dueDate (YYYY-MM-DD), lastUpdate, escalated (Yes|No), status (Open|In Progress|Validated|Overdue|Closed)
SCALAR FIELDS — call fill_fields: fu-map-updates, fu-remediation, fu-escalations, fu-conclusion`,
  }[tab] || '';
}

function buildSystemPrompt(audit, focusPhase, activeTab) {
  const resolvedPhase = focusPhase || audit.currentPhase || 'discovery';
  const phaseIdx = PHASES.indexOf(resolvedPhase);
  const disc = audit.discovery || {};
  const plan = audit.planning  || {};
  const fw   = audit.fieldwork || {};
  const rep  = audit.reporting || {};
  const wu   = audit.wrapup    || {};
  const fu   = audit.followup  || {};

  const exceptions  = (fw['fw-worksteps']    || []).filter(w => w.result === 'exception');
  const issues      = (rep['rep-issues-log'] || []);
  const openMAPs    = (wu['wu-maps']         || []).filter(m => m.status !== 'closed');
  const openIssues  = (fu['fu-open-issues']  || []);
  const riskReg     = (plan['plan-risk-register'] || []);

  const lines = [
    `You are an expert AI assistant for Verizon Internal Audit (VIA).`,
    `You know VIA's exact methodology from the March 2026 Audit Manual — phases, issue types (PI/IO/RI),`,
    `sampling tables, FROST risk framework, observation writing standards, MAP ownership rules, audit opinions,`,
    `Optro workflows, and all VIA-specific terminology. Apply this knowledge precisely.`,
    VIA_DOMAIN_KNOWLEDGE,
    ``,
    `════════════════════════════════════════`,
    `ACTIVE ENGAGEMENT`,
    `════════════════════════════════════════`,
    `Title:           ${audit.title || 'Untitled Audit'}`,
    `Entity / Auditee:${disc['disc-entity']        || '(not set)'}`,
    `Business Unit:   ${disc['disc-business-unit'] || '(not set)'}`,
    `Audit Type:      ${disc['disc-audit-type']    || '(not set)'}`,
    `Risk Category:   ${disc['disc-risk-category'] || '(not set)'}`,
    `Audit Period:    ${disc['disc-period-start']  || '?'} – ${disc['disc-period-end'] || '?'}`,
    `Lead Auditor:    ${disc['disc-lead-auditor']  || '(not set)'}`,
    `Audit Manager:   ${disc['disc-manager']       || '(not set)'}`,
    `Team Members:    ${disc['disc-team-members']  || '(not set)'}`,
    `Stakeholders:    ${disc['disc-stakeholders']  || '(not set)'}`,
    ``,
    `════════════════════════════════════════`,
    `LIFECYCLE STATUS`,
    `════════════════════════════════════════`,
    `Current Phase:   ${PHASE_LABELS[audit.currentPhase]} (${phaseIdx + 1} of 6)`,
    `Phases complete: ${PHASES.slice(0, phaseIdx).map(p => PHASE_LABELS[p]).join(', ') || 'None'}`,
  ];

  if (disc['disc-background'])    lines.push(``, `Background:`, disc['disc-background']);
  if (disc['disc-rationale'])     lines.push(``, `Audit Rationale:`, disc['disc-rationale']);
  if (disc['disc-scope'])         lines.push(``, `Scope (Discovery):`, disc['disc-scope']);
  if (disc['disc-scope-in'])      lines.push(``, `Scope-In:`, disc['disc-scope-in']);
  if (disc['disc-scope-out'])     lines.push(``, `Scope-Out:`, disc['disc-scope-out']);
  if (disc['disc-risk-areas'])    lines.push(``, `Risk Areas Identified:`, disc['disc-risk-areas']);
  if (disc['disc-prior-history']) lines.push(``, `Prior Audit History:`, disc['disc-prior-history']);
  if (disc['disc-systems'])       lines.push(``, `Key Systems (Discovery):`, disc['disc-systems']);

  if (plan['plan-objectives'])       lines.push(``, `Audit Objectives:`, plan['plan-objectives']);
  if (plan['plan-scope'])            lines.push(`Planned Scope:`, plan['plan-scope']);
  if (plan['plan-out-of-scope'])     lines.push(`Out of Scope:`, plan['plan-out-of-scope']);
  if (plan['plan-testing-approach']) lines.push(`Testing Approach:`, plan['plan-testing-approach']);
  if (plan['plan-key-systems'])      lines.push(`Key Systems:`, plan['plan-key-systems']);

  // When focused on Risk Register tab, emit a dedicated context block so the LLM produces
  // specific, engagement-relevant risks rather than generic ones.
  if (activeTab === 'rr') {
    const rrContext = [
      disc['disc-entity']        && `Entity: ${disc['disc-entity']}`,
      disc['disc-audit-type']    && `Audit Type: ${disc['disc-audit-type']}`,
      disc['disc-rationale']     && `Why we are auditing this: ${disc['disc-rationale']}`,
      disc['disc-background']    && `Background: ${disc['disc-background']}`,
      disc['disc-scope-in']      && `In scope: ${disc['disc-scope-in']}`,
      disc['disc-scope-out']     && `Out of scope: ${disc['disc-scope-out']}`,
      disc['disc-risk-areas']    && `Known risk areas: ${disc['disc-risk-areas']}`,
      disc['disc-prior-history'] && `Prior audit history: ${disc['disc-prior-history']}`,
      disc['disc-systems']       && `Key systems involved: ${disc['disc-systems']}`,
      plan['plan-objectives']    && `Audit objectives: ${plan['plan-objectives']}`,
      plan['plan-testing-approach'] && `Testing approach: ${plan['plan-testing-approach']}`,
    ].filter(Boolean).join('\n');
    if (rrContext) {
      lines.push(
        ``, `════════════════════════════════════════`,
        `RISK REGISTER GENERATION CONTEXT`,
        `════════════════════════════════════════`,
        `Use ALL of the following engagement-specific context to generate risks that are precise`,
        `and relevant to this exact audit — not generic boilerplate:`,
        ``, rrContext,
        ``,
        `IMPORTANT — TEST PROCEDURES: Every in-scope risk row MUST include a procedure field`,
        `following this exact 4-step pattern (derived from VIA fieldwork methodology):`,
        `  Step 1 — Walkthrough: Conduct a walkthrough with [specific control owner/role] to understand`,
        `    the end-to-end process and confirm control design. Perform a test of one.`,
        `  Step 2 — Occurrence: Obtain [specific evidence type — meeting logs, system records, training`,
        `    attendance sheets, approval emails, etc.] to confirm the control occurred during the audit period.`,
        `  Step 3 — Content Review: Review [specific documentation — policy, training deck, checklist, etc.]`,
        `    to confirm content is adequate, current, and addresses the identified risk.`,
        `  Step 4 — Recalibration: Review the process used to update/recalibrate [materials/controls]`,
        `    to ensure it reflects current risk conditions and any prior audit findings.`,
      );
    }
  }

  if (riskReg.length) {
    lines.push(``, `Risk Register (${riskReg.length}):`);
    riskReg.forEach(r => lines.push(`  • ${r.risk || r} [${r.rating || '?'}] — ${r.control || 'No control noted'}`));
  }

  if (fw['fw-status'])   lines.push(``, `Fieldwork Status:`, fw['fw-status']);
  if (fw['fw-evidence']) lines.push(`Evidence Summary:`, fw['fw-evidence']);
  if (exceptions.length) {
    lines.push(``, `Exceptions Found (${exceptions.length}):`);
    exceptions.forEach(w => lines.push(`  • ${w.step || w.title}: ${w.notes || ''}`));
  }

  if (issues.length) {
    lines.push(``, `Issues Log (${issues.length}):`);
    issues.forEach(i => lines.push(`  • [${i.risk || '?'}] ${i.title || i.issue} — ${i.recommendation || 'no rec yet'}`));
  }

  if (rep['rep-exec-summary']) {
    const s = rep['rep-exec-summary'];
    lines.push(``, `Draft Exec Summary (excerpt):`, s.length > 400 ? s.slice(0, 400) + '…' : s);
  }
  if (rep['rep-opinion']) lines.push(`Audit Opinion:`, rep['rep-opinion']);

  if (openMAPs.length) {
    lines.push(``, `Open MAPs (${openMAPs.length}):`);
    openMAPs.forEach(m => lines.push(`  • ${m.title || m}: due ${m.dueDate || '?'}, owner ${m.owner || '?'}`));
  }
  if (fu['fu-conclusion']) lines.push(``, `Follow-Up Conclusion:`, fu['fu-conclusion']);

  lines.push(
    ``,
    `════════════════════════════════════════`,
    `YOUR CURRENT FOCUS — ${(PHASE_LABELS[resolvedPhase] || resolvedPhase).toUpperCase()}`,
    `════════════════════════════════════════`,
    getPhaseGuidance(resolvedPhase),
  );

  if (activeTab) {
    const tabGuidance = getTabGuidance(activeTab);
    if (tabGuidance) {
      lines.push(
        ``,
        `────────────────────────────────────────`,
        `ACTIVE TAB CONTEXT`,
        `────────────────────────────────────────`,
        tabGuidance,
      );
    }
  }

  lines.push(
    ``,
    `════════════════════════════════════════`,
    `FORM-FILLING TOOLS`,
    `════════════════════════════════════════`,
    `You have twelve tools: fill_fields, fill_sampling_card, fill_walkthroughs, fill_risk_register, fill_drl, fill_data_validation, fill_audit_programs, fill_data_sampling, fill_worksteps, fill_issues, fill_maps, and fill_fu_issues.`,
    `CRITICAL: ALWAYS call fill_fields (or the appropriate specialized tool) when the user asks you to`,
    `"fill in", "write", "draft", "populate", "build", "complete", "generate", or "create" ANY section.`,
    `Do NOT just write text — actually call the tool. The user wants to SEE fields populated on screen.`,
    `fill_fields: fills any individual textarea/input/select field by exact ID (see tool description for full ID list).`,
    `fill_walkthroughs: adds rows to the Walkthrough Log table.`,
    `fill_risk_register: adds rows to the Risk Register table.`,
    `fill_drl: populates the Document Request List table.`,
    `fill_data_validation: populates the Data Validation table.`,
    `fill_audit_programs: generates AuditBoard Program Procedure cards in Fieldwork.`,
    `fill_data_sampling: generates Data Sampling procedure and documentation cards.`,
    `fill_worksteps: adds rows to the Fieldwork Workstep Log.`,
    `fill_issues: adds rows to the Reporting Issues/Findings log.`,
    `fill_maps: adds rows to the Wrap-Up MAPs table.`,
    `fill_fu_issues: adds rows to the Follow-Up tracking table.`,
    `When filling multiple fields at once, pass ALL of them in a single fill_fields call.`,
    `ALWAYS also include a brief explanation of what you filled — never use a tool silently.`,
    `Do NOT use tools when the user asks for pure analysis, advice, or says "just tell me."`,
    ``,
    `ALWAYS: reference the engagement details above when relevant. Use IIA terminology.`,
    `Be direct and actionable. Format with bullets or numbered steps when helpful.`,
    `If fields are empty, offer to help draft content. Flag any methodology concerns.`,
  );

  return lines.join('\n');
}

// ── AI Tool Definitions ───────────────────────────────────────────────────────

const AUDIT_TOOLS = [
  {
    name: 'fill_fields',
    description: `Fill one or more form fields directly on the audit page. Call this tool whenever the user asks you to write, fill in, draft, populate, or complete ANY text section. ALWAYS use the exact field IDs listed below — wrong IDs silently do nothing.

═══ PRE-DISCOVERY (Announcement Email) ═══
  email-subject       — Announcement email subject line
  email-body          — Full announcement email body text
  email-sent-date     — Date email was sent (YYYY-MM-DD)

═══ DISCOVERY ═══
  audit-title-input   — Audit title (top of page header)

  Engagement Setup:
  disc-entity         — Entity / Auditee name
  disc-bu             — Business Unit (select: Comms | Media | Technology | Corporate | Enterprise | Consumer | Business | Other)
  disc-audit-type     — Audit Type (select: Major Process Audit | Controls Advisory | Spot Check | Integrated Audit | Follow-Up | IT Audit | Data Analytics)
  disc-frost          — Primary FROST Risk (select: F - Financial | R - Regulatory | O - Operational | S - Strategic)
  disc-period-start   — Audit period start date (YYYY-MM-DD)
  disc-period-end     — Audit period end date (YYYY-MM-DD)
  disc-target-report  — Target report date (YYYY-MM-DD)
  disc-restriction    — Restriction Level (select: Unrestricted | Restricted | Highly Restricted)
  disc-uid            — Optro Audit UID
  disc-plan-year      — Audit Plan Year (select: 2024 | 2025 | 2026 | 2027)

  Audit Team:
  disc-lead           — Lead Auditor name
  disc-manager        — Audit Manager / Team Lead name
  disc-ad             — Associate Director name
  disc-sd             — Senior Director name
  disc-faast          — FAAST Support contact
  disc-tech           — Technology Audit Support contact
  disc-team           — Additional team members / co-source staff

  Background & Rationale:
  disc-background     — Business Overview (entity context, operations, recent changes — 2–4 paragraphs)
  disc-rationale      — Audit Rationale — why was this selected (Risk360 signal, trigger, etc.)
  disc-prior-history  — Prior Audit History (past audits, opinions, open MAPs, repeat findings)
  disc-systems        — Key Systems & Applications (SAP, Salesforce, custom tools, etc.)
  disc-data-needs     — Data & Analytics Needs (extracts, sources, volumes)

  Preliminary Scope:
  disc-scope-in       — In Scope (processes, systems, locations, time periods included)
  disc-scope-out      — Out of Scope (explicit exclusions and rationale)

  Dates & Notes:
  disc-announce-date  — Announcement email sent date (YYYY-MM-DD)
  disc-kickoff-date   — Internal Kickoff Meeting date (YYYY-MM-DD)
  disc-notes          — Discovery Notes & Open Items

═══ PLANNING — WALKTHROUGHS TAB ═══
  plan-wt-notes       — Overall walkthrough notes and observations
  plan-sme-engaged    — SME Engaged? (select: Yes | No | In Progress)
  plan-sme-name       — SME Contact name(s)
  plan-sme-date       — SME engagement date (YYYY-MM-DD)
  plan-sme-scope      — Scope of SME engagement

═══ PLANNING — RISK REGISTER TAB ═══
  plan-rr-signoff-name  — Sr. Director sign-off name
  plan-rr-signoff-date  — Sign-off date (YYYY-MM-DD)
  plan-rr-signoff-notes — Sign-off notes
  plan-rr-notes         — Risk Register notes

═══ PLANNING — DATA VALIDATION TAB (meta fields) ═══
  dv-who-generated    — Who generated the report (name/team/system)
  dv-date-generated   — Date report was generated (YYYY-MM-DD)
  dv-source           — Source system / database name
  dv-report-type      — Report type (select: Standard Report | Ad-hoc Query | System Export | Manual Extract | FAAST Pull | Other)
  dv-dataset-link     — Link to dataset (Google Drive / SharePoint URL)
  dv-code-text        — SQL / Python / R / VBA code for the data pull

═══ PLANNING — OPENING MEETING TAB ═══
  plan-qa-name        — AD / Director name (QA confirmation)
  plan-qa-date        — QA confirmation date (YYYY-MM-DD)
  plan-qa-notes       — QA notes
  plan-om-date        — Opening meeting date (YYYY-MM-DD)
  plan-om-time        — Opening meeting time (HH:MM)
  plan-om-location    — Location / virtual link
  plan-om-scope-in    — In-scope areas presented at meeting
  plan-om-scope-out   — Out-of-scope areas presented at meeting
  plan-om-design-gaps — Control / design gaps discussed
  plan-om-notes       — Meeting notes & action items

═══ FIELDWORK ═══
  fw-status           — Current fieldwork status & progress narrative
  fw-tech-areas       — Technology / systems reviewed
  fw-contacts         — Key contacts engaged during fieldwork
  fw-issues           — Issues found during testing (pre-formal write-up)
  fw-evidence         — Evidence log

═══ REPORTING ═══
  rep-exec-summary    — Executive Summary (scope, methodology, key findings, overall conclusion — lead paragraph executives will read)
  rep-opinion         — Audit Opinion (select: Effective | Opportunities for Improvement | Needs Significant Improvement | Inadequate)
  rep-notes           — Report notes & draft status
  rep-vetting         — Closing meeting / vetting notes
  rep-distribution    — Report distribution list

═══ WRAP-UP ═══
  wu-final-opinion    — Final Audit Opinion (select: same values as rep-opinion)
  wu-mgmt-responses   — Management responses summary
  wu-retrospective    — Post-audit retrospective
  wu-lessons          — Lessons learned

═══ FOLLOW-UP ═══
  fu-map-updates      — MAP update summary
  fu-remediation      — Remediation evidence reviewed
  fu-escalations      — Escalations
  fu-conclusion       — Closure conclusion

IMPORTANT: For Walkthrough Log rows use fill_walkthroughs. For Risk Register rows use fill_risk_register. For DRL use fill_drl. For Data Validation table use fill_data_validation. For Audit Program cards use fill_audit_programs. For Data Sampling cards use fill_data_sampling. For Fieldwork workstep rows use fill_worksteps. For Reporting issues use fill_issues. For Wrap-Up MAPs use fill_maps. For Follow-Up tracking rows use fill_fu_issues. Use fill_fields for ALL other individual text/select fields.`,
    input_schema: {
      type: 'object',
      properties: {
        fields: {
          type: 'object',
          description: 'Object mapping exact field IDs (from the list above) to string values.',
          additionalProperties: { type: 'string' },
        },
      },
      required: ['fields'],
    },
  },
  {
    name: 'fill_data_validation',
    description: `Populate the Data Validation (C&A) table. Each row validates one data extract or report.

APPROACH valid values: Profiling Data | Front End Confirmation | Reconciliations | Reasonableness Check | Data Owner Backend Walkthrough | Independent Generation | Control Totals | IPE Procedures
CONCLUSION valid values: Complete & Accurate | Issues Noted | Pending
MODE: "replace" = clear existing rows first; "append" = add to existing`,
    input_schema: {
      type: 'object',
      properties: {
        rows: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              dataDesc:     { type: 'string', description: 'Specific name of the data extract or report' },
              sourceSystem: { type: 'string', description: 'System the data comes from (e.g. "vRepair", "SAP")' },
              approach:     { type: 'string', description: 'Validation approach' },
              conclusion:   { type: 'string', enum: ['Complete & Accurate', 'Issues Noted', 'Pending'] },
              notes:        { type: 'string', description: 'Notes on what to validate or any exceptions' },
            },
            required: ['dataDesc', 'sourceSystem', 'approach'],
          },
        },
        mode: { type: 'string', enum: ['replace', 'append'] },
      },
      required: ['rows'],
    },
  },
  {
    name: 'fill_drl',
    description: `Populate the Document Request List (DRL) table directly. Call this when the user asks to generate, build, or populate the DRL — especially from Risk Register test procedures. Always include explanatory text alongside every tool call.

Each DRL row represents one document or evidence item needed to execute a test procedure:
- testStep: Risk ID + Control ID reference (e.g. "A.R1 — A.R1.C1"). If one document serves multiple risks, list all: "A.R1 / B.R1 — A.R1.C1 / B.R1.C1"
- docDesc: Specific artifact name — name the actual report, system, policy, or log (e.g. "Monthly exception report from vRepair listing unresolved tickets > 30 days"). Never write generic placeholders.
- requestee: Role who provides this (e.g. "IT Operations Manager", "Control Owner", "Finance Team")
- notes: One sentence on what to verify in this document

Cover all 4 testing steps for each in-scope risk: Walkthrough (process docs, org charts, procedure guides), Occurrence (transaction logs, approvals, system reports), Content Review (policies, training decks, checklists), Recalibration (review sign-off records, update logs).
MODE: "replace" = clear existing rows first; "append" = add to existing (default)`,
    input_schema: {
      type: 'object',
      properties: {
        rows: {
          type: 'array',
          description: 'Array of DRL rows to add to the Document Request List.',
          items: {
            type: 'object',
            properties: {
              testStep:  { type: 'string', description: 'Risk ID + Control ID reference' },
              docDesc:   { type: 'string', description: 'Specific document or evidence description' },
              requestee: { type: 'string', description: 'Role or team who provides this document' },
              notes:     { type: 'string', description: 'What to verify in this document' },
            },
            required: ['testStep', 'docDesc'],
          },
        },
        mode: {
          type: 'string',
          enum: ['replace', 'append'],
          description: 'replace: clear existing rows first. append: add to existing. Default: replace when generating fresh from RR.',
        },
      },
      required: ['rows'],
    },
  },
  {
    name: 'fill_sampling_card',
    description: `Fill the fields of an existing open Sampling Plan card in-place. Use when the user wants the current card populated. If no card exists yet, one will be created.

FREQ valid values: Daily | Weekly | Monthly | Quarterly | Annual | Ad-hoc
auditProcedure: array of up to 3 steps. Step 1 = pull complete population listing.`,
    input_schema: {
      type: 'object',
      properties: {
        name:           { type: 'string', description: 'Procedure name' },
        risk:           { type: 'string', description: 'Associated risk statement' },
        control:        { type: 'string', description: 'Control description' },
        freq:           { type: 'string', enum: ['Daily','Weekly','Monthly','Quarterly','Annual','Ad-hoc'] },
        objective:      { type: 'string', description: 'Test objective' },
        auditProcedure: { type: 'array', items: { type: 'string' }, description: 'Up to 3 audit procedure steps' },
      },
      required: ['name', 'risk', 'control', 'objective', 'auditProcedure'],
    },
  },
  {
    name: 'fill_walkthroughs',
    description: `Add rows to the Walkthrough Log table. Call this when the user asks to populate, generate, or add walkthrough entries. Each row documents one process or control area walked through.

DESIGN valid values: Effective | Partial Gap | Gap | Not Tested
GAP DOCUMENTED valid values: Yes | No
MODE: "replace" = clear existing rows first; "append" = add to existing (default)

If design is "Partial Gap" or "Gap", set gap to "Yes" and note the gap in the notes field. A design gap means STOP testing that control and record it in the Risk Register as "Missing Control - Control Gap".`,
    input_schema: {
      type: 'object',
      properties: {
        rows: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              process: { type: 'string', description: 'Process or control area name, e.g. "Access Provisioning Review"' },
              owner:   { type: 'string', description: 'KPO/control owner name, e.g. "John Smith, IT Manager"' },
              systems: { type: 'string', description: 'Systems involved, e.g. "Active Directory, ServiceNow"' },
              date:    { type: 'string', description: 'Walkthrough date YYYY-MM-DD' },
              design:  { type: 'string', enum: ['Effective', 'Partial Gap', 'Gap', 'Not Tested'], description: 'Control design assessment' },
              gap:     { type: 'string', enum: ['Yes', 'No'], description: 'Was a design gap documented?' },
            },
            required: ['process'],
          },
        },
        mode: { type: 'string', enum: ['replace', 'append'] },
      },
      required: ['rows'],
    },
  },
  {
    name: 'fill_risk_register',
    description: `Add rows to the Risk Register table. Call this when the user asks to build, populate, or generate the risk register.

RISK DESC FORMAT (desc field):
  Line 1: "A.R1 - [Short Risk Title]"
  Line 2: "(DESCRIPTION) [what can fail], (ROOT CAUSE) due to [control gap], (IMPACT) resulting in [consequence]."

CONTROLS FORMAT:
  "(WHO) [who performs] (WHAT) [action] (WHEN) [frequency] (WHERE) [system/location] (WHY) to ensure [objective]"

TEST PROCEDURE: 4 steps — Walkthrough, Occurrence, Content Review, Recalibration.

CTYPE valid values: Preventative/Manual | Preventative/Automated | Preventative/Hybrid | Detective/Manual | Detective/Automated | Detective/Hybrid | Corrective/Manual | Corrective/Automated | Missing Control - Control Gap
FREQ valid values: Daily | Weekly | Monthly | Quarterly | Annual | Ad-hoc
INSCOPE valid values: Yes | No
MODE: "replace" = clear existing rows first; "append" = add to existing (default)`,
    input_schema: {
      type: 'object',
      properties: {
        rows: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              riskId:        { type: 'string', description: 'e.g. A.R1' },
              desc:          { type: 'string', description: '2-line risk description in VIA format' },
              frostType:     { type: 'string', enum: ['Technology & Information Security', 'Operational', 'Financial', 'Regulatory/Compliance', 'Strategic'] },
              likelihood:    { type: 'string', enum: ['Likely', 'Possible', 'Unlikely'] },
              impact:        { type: 'string', enum: ['High', 'Medium', 'Low'] },
              controls:      { type: 'string', description: 'Control description in WHO/WHAT/WHEN/WHERE/WHY format' },
              testProcedure: { type: 'string', description: '4-step test procedure' },
              ctype:         { type: 'string', description: 'Control type' },
              freq:          { type: 'string', description: 'Control frequency' },
              inscope:       { type: 'string', enum: ['Yes', 'No'] },
            },
            required: ['riskId', 'desc', 'likelihood', 'impact'],
          },
        },
        mode: { type: 'string', enum: ['replace', 'append'] },
      },
      required: ['rows'],
    },
  },
  {
    name: 'fill_audit_programs',
    description: `Generate AuditBoard Program Procedure cards for the Fieldwork section — one card per DRL test step or control. Call this when the user asks to generate, build, or populate AuditBoard audit program procedures from the DRL. Always include explanatory text.

Each card maps to one control/test step and has:
  - name: procedure name in AuditBoard format, e.g. "A.R1.C1 — Access Review: Approval Controls"
  - risk: full risk statement from the Risk Register (2-line VIA format)
  - control: full control description from the Risk Register
  - controlFreq: one of Daily | Weekly | Monthly | Quarterly | Annual | Ad-hoc
  - timePeriod: test coverage period, e.g. "Jan 1 – Dec 31, 2024"
  - testObjective: what the test verifies, e.g. "Confirm that access provisioning approvals are documented and completed by an authorized approver prior to access being granted."
  - auditProcedure: array of up to 5 steps — what the auditor will DO (obtain, review, inspect, compare, confirm)
  - workPerformed: array of up to 5 steps — what was actually done during testing (past tense: "Obtained...", "Reviewed...", "Confirmed...")
  - testingAttributes: array of up to 5 attributes — specific items to check on each sample (e.g. "Approval email present", "Date of approval precedes access grant date")

AUDIT PROCEDURE FORMAT: numbered steps describing what to do. Step 1 should always be a walkthrough step.
WORK PERFORMED FORMAT: leave as brief templates the auditor will fill in (e.g. "Obtained [X] from [system/person] covering [period].").
TESTING ATTRIBUTES FORMAT: specific binary checkpoints per sample (e.g. "Access was approved prior to provisioning date").`,
    input_schema: {
      type: 'object',
      properties: {
        rows: {
          type: 'array',
          description: 'Array of audit program procedure cards — one per DRL test step.',
          items: {
            type: 'object',
            properties: {
              name:              { type: 'string', description: 'AuditBoard Program Procedure Name, e.g. A.R1.C1 — Control Short Title' },
              risk:              { type: 'string', description: 'Risk statement from Risk Register' },
              control:           { type: 'string', description: 'Control description from Risk Register' },
              controlFreq:       { type: 'string', enum: ['Daily','Weekly','Monthly','Quarterly','Annual','Ad-hoc'] },
              timePeriod:        { type: 'string', description: 'Time period covered by the test, e.g. Jan 1 – Dec 31, 2024' },
              testObjective:     { type: 'string', description: 'What the test is designed to verify' },
              auditProcedure:    { type: 'array', items: { type: 'string' }, description: 'Up to 5 numbered audit procedure steps (what the auditor will do)' },
              workPerformed:     { type: 'array', items: { type: 'string' }, description: 'Up to 5 work performed steps (template text for what was done)' },
              testingAttributes: { type: 'array', items: { type: 'string' }, description: 'Up to 5 testing attributes (specific checkpoints per sample item)' },
            },
            required: ['name', 'risk', 'control', 'testObjective', 'auditProcedure', 'workPerformed', 'testingAttributes'],
          },
        },
        mode: {
          type: 'string',
          enum: ['replace', 'append'],
          description: 'replace: clear existing cards first (default). append: add to existing.',
        },
      },
      required: ['rows'],
    },
  },
  {
    name: 'fill_data_sampling',
    description: `Populate the Data Sampling tab with procedure cards. One card per in-scope control.

Apply VIA sampling table: 1–249 pop → 25; 250–999 → 40; 1,000+ → 60.
FREQ valid values: Daily | Weekly | Monthly | Quarterly | Annual | Ad-hoc
auditProcedure: up to 3 steps. Step 1 = pull complete population listing.
MODE: "replace" = clear existing; "append" = add to existing`,
    input_schema: {
      type: 'object',
      properties: {
        procedures: {
          type: 'array',
          description: 'Procedure cards — one per in-scope control',
          items: {
            type: 'object',
            properties: {
              name:           { type: 'string', description: 'Procedure name' },
              risk:           { type: 'string', description: 'Associated risk' },
              control:        { type: 'string', description: 'Control description' },
              freq:           { type: 'string', enum: ['Daily','Weekly','Monthly','Quarterly','Annual','Ad-hoc'] },
              objective:      { type: 'string', description: 'Test objective' },
              auditProcedure: { type: 'array', items: { type: 'string' }, description: 'Up to 3 procedure steps' },
            },
            required: ['name', 'risk', 'control', 'objective', 'auditProcedure'],
          },
        },
        mode: { type: 'string', enum: ['replace', 'append'] },
      },
      required: ['procedures'],
    },
  },
  {
    name: 'fill_worksteps',
    description: `Add rows to the Fieldwork Workstep Log. Call this when the user asks to populate, generate, or log fieldwork worksteps or testing results. One row per control/test step.

STATUS valid values: Not Started | In Progress | Complete | Exception Found | Skipped
MODE: replace = clear first; append = add to existing`,
    input_schema: {
      type: 'object',
      properties: {
        rows: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              stepId:     { type: 'string', description: 'Control ID reference, e.g. A.R1.C1' },
              control:    { type: 'string', description: 'Control or area being tested' },
              procedures: { type: 'string', description: 'Testing procedures performed' },
              sampleSize: { type: 'string', description: 'Number of samples tested, e.g. 25' },
              evidence:   { type: 'string', description: 'Evidence gathered and notes' },
              status:     { type: 'string', enum: ['Not Started', 'In Progress', 'Complete', 'Exception Found', 'Skipped'] },
            },
            required: ['stepId', 'control'],
          },
        },
        mode: { type: 'string', enum: ['replace', 'append'] },
      },
      required: ['rows'],
    },
  },
  {
    name: 'fill_issues',
    description: `Add rows to the Reporting Issues/Findings log. Call this when the user asks to draft, log, or populate audit findings or issues in the Reporting phase.

ISSUE TYPES: PI = Process Issue (reportable finding) | IO = Informational Observation | RI = Risk Item
RISK LEVEL: High | Medium | Low
STATUS: Open | In Draft | Agreed | Disputed | Closed
MODE: replace = clear first; append = add to existing`,
    input_schema: {
      type: 'object',
      properties: {
        rows: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              id:             { type: 'string', description: 'Issue ID e.g. PI-01, IO-01, RI-01' },
              type:           { type: 'string', enum: ['PI', 'IO', 'RI'] },
              title:          { type: 'string', description: 'Issue title and full description' },
              recommendation: { type: 'string', description: 'Recommended corrective action' },
              riskLevel:      { type: 'string', enum: ['High', 'Medium', 'Low'] },
              mgmtResponse:   { type: 'string', description: 'Management response and agreed action' },
              status:         { type: 'string', enum: ['Open', 'In Draft', 'Agreed', 'Disputed', 'Closed'] },
            },
            required: ['id', 'type', 'title'],
          },
        },
        mode: { type: 'string', enum: ['replace', 'append'] },
      },
      required: ['rows'],
    },
  },
  {
    name: 'fill_maps',
    description: `Add rows to the Wrap-Up Management Action Plans (MAPs) table. Call this when the user asks to create, generate, or populate MAPs for agreed audit findings.

STATUS: Open | In Progress | Completed | Overdue | Escalated
MODE: replace = clear first; append = add to existing`,
    input_schema: {
      type: 'object',
      properties: {
        rows: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              mapId:   { type: 'string', description: 'MAP ID e.g. MAP-01' },
              issue:   { type: 'string', description: 'Issue or finding reference' },
              action:  { type: 'string', description: 'Agreed remediation action' },
              owner:   { type: 'string', description: 'Action owner name and role' },
              dueDate: { type: 'string', description: 'Due date YYYY-MM-DD' },
              status:  { type: 'string', enum: ['Open', 'In Progress', 'Completed', 'Overdue', 'Escalated'] },
              notes:   { type: 'string', description: 'Additional notes' },
            },
            required: ['mapId', 'issue', 'action'],
          },
        },
        mode: { type: 'string', enum: ['replace', 'append'] },
      },
      required: ['rows'],
    },
  },
  {
    name: 'fill_fu_issues',
    description: `Add rows to the Follow-Up issue tracking table. Call this when the user asks to populate or update follow-up tracking for open MAPs or findings.

ESCALATED: Yes | No
STATUS: Open | In Progress | Validated | Overdue | Closed
MODE: replace = clear first; append = add to existing`,
    input_schema: {
      type: 'object',
      properties: {
        rows: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              mapId:      { type: 'string', description: 'MAP ID reference e.g. MAP-01' },
              issue:      { type: 'string', description: 'Issue or finding description' },
              owner:      { type: 'string', description: 'Owner name' },
              dueDate:    { type: 'string', description: 'Due date YYYY-MM-DD' },
              lastUpdate: { type: 'string', description: 'Latest status update or notes' },
              escalated:  { type: 'string', enum: ['Yes', 'No'] },
              status:     { type: 'string', enum: ['Open', 'In Progress', 'Validated', 'Overdue', 'Closed'] },
            },
            required: ['mapId', 'issue'],
          },
        },
        mode: { type: 'string', enum: ['replace', 'append'] },
      },
      required: ['rows'],
    },
  },
];

// ── AI Chat (SSE streaming with tool use) ────────────────────────────────────

app.post('/api/audits/:id/chat', async (req, res) => {
  const audits = readAudits();
  const idx    = audits.findIndex(a => a.id === req.params.id);
  if (idx === -1) return res.status(404).json({ error: 'Not found' });

  const audit = audits[idx];
  const userMessage = req.body.message;
  if (!userMessage) return res.status(400).json({ error: 'No message' });

  // Map client phase keys to server phase keys
  const phaseMap = { disc: 'discovery', plan: 'planning' };
  const rawPhase = req.body.activePhase;
  const focusPhase = phaseMap[rawPhase] || rawPhase || null;
  const activeTab  = req.body.activeTab  || null;

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');

  const history = (audit.chatHistory || []).slice(-30);
  const messages = [
    ...history.map(m => ({ role: m.role, content: m.content })),
    { role: 'user', content: userMessage },
  ];

  let fullResponse = '';
  const toolAccumulators = {}; // index → { name, id, json }
  try {
    const stream = await client.messages.stream({
      model: 'claude-sonnet-4-6',
      max_tokens: 16000,
      system: buildSystemPrompt(audit, focusPhase, activeTab),
      tools: AUDIT_TOOLS,
      tool_choice: { type: 'auto' },
      messages,
    });

    for await (const chunk of stream) {
      if (chunk.type === 'content_block_start') {
        if (chunk.content_block && chunk.content_block.type === 'tool_use') {
          toolAccumulators[chunk.index] = {
            name: chunk.content_block.name,
            id:   chunk.content_block.id,
            json: '',
          };
        }
      } else if (chunk.type === 'content_block_delta') {
        if (chunk.delta.type === 'text_delta') {
          fullResponse += chunk.delta.text;
          res.write(`data: ${JSON.stringify({ text: chunk.delta.text })}\n\n`);
        } else if (chunk.delta.type === 'input_json_delta') {
          if (toolAccumulators[chunk.index]) {
            toolAccumulators[chunk.index].json += chunk.delta.partial_json;
          }
        }
      } else if (chunk.type === 'content_block_stop') {
        const acc = toolAccumulators[chunk.index];
        if (acc) {
          try {
            const input = JSON.parse(acc.json);
            res.write(`data: ${JSON.stringify({ tool_call: { name: acc.name, input } })}\n\n`);
          } catch(e) { console.warn('Tool JSON parse failed for', acc.name, ':', e.message); }
          delete toolAccumulators[chunk.index];
        }
      }
    }
    res.write('data: [DONE]\n\n');
    res.end();

    // Persist conversation
    if (!audits[idx].chatHistory) audits[idx].chatHistory = [];
    audits[idx].chatHistory.push(
      { role: 'user',      content: userMessage,                        timestamp: new Date().toISOString() },
      { role: 'assistant', content: fullResponse || '[form fields filled]', timestamp: new Date().toISOString() },
    );
    audits[idx].updatedAt = new Date().toISOString();
    writeAudits(audits);
  } catch (e) {
    res.write(`data: ${JSON.stringify({ error: e.message })}\n\n`);
    res.end();
  }
});

// ── DRL Generator (no-tools, returns raw JSON array) ─────────────────────────
app.post('/api/audits/:id/gen-drl', async (req, res) => {
  const audits = readAudits();
  const audit  = audits.find(a => a.id === req.params.id);
  if (!audit) return res.status(404).json({ error: 'Not found' });

  const { rrRows } = req.body;
  if (!rrRows || !rrRows.length) return res.status(400).json({ error: 'No RR rows provided' });

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');

  const rrSummary = rrRows.map(r => {
    const parts = [];
    if (r.testProcedure) parts.push(`Test Procedure: ${r.testProcedure}`);
    if (r.finops)        parts.push(`FinOps Procedure: ${r.finops}`);
    if (r.tech)          parts.push(`Technology Procedure: ${r.tech}`);
    if (r.faast)         parts.push(`FAAST Procedure: ${r.faast}`);
    return `Risk ${r.riskId} | Control ${r.ctrlId} | Scope: ${r.scopeArea}\nDescription: ${r.desc}\n${parts.join('\n')}`;
  }).join('\n\n---\n\n');

  const systemPrompt = `You are an expert Verizon Internal Audit (VIA) document request specialist. You produce concise, specific Document Request Lists from test procedures. Return ONLY a valid JSON array — no prose, no markdown, no code fences.`;

  const userPrompt = `Based on these Risk Register test procedures, generate a Document Request List (DRL).

Return a JSON array where each element has exactly these keys:
- "testStep": Risk ID + Control ID reference, e.g. "A.R1 — A.R1.C1"
- "docDesc": specific name of the document or evidence (e.g. "Monthly exception report from vRepair showing unresolved tickets > 30 days")
- "requestee": role who provides it (e.g. "IT Operations Manager", "Control Owner", "Finance Team")
- "notes": one sentence on what to verify in this document

Rules:
- One row per distinct document type needed
- Be specific — name the actual artifact, system, or report, not generic "supporting documentation"
- Cover walkthroughs (process docs, org charts), occurrence testing (transaction logs, approvals), content review (policies, procedures), and recalibration evidence (review sign-offs)
- If one document satisfies multiple risks, list both Risk IDs in testStep
- Return ONLY the JSON array, nothing else

Risk Register Test Procedures:
${rrSummary}`;

  try {
    const stream = await client.messages.stream({
      model: 'claude-sonnet-4-6',
      max_tokens: 4096,
      system: systemPrompt,
      messages: [{ role: 'user', content: userPrompt }],
    });

    for await (const chunk of stream) {
      if (chunk.type === 'content_block_delta' && chunk.delta.type === 'text_delta') {
        res.write(`data: ${JSON.stringify({ text: chunk.delta.text })}\n\n`);
      }
    }
    res.write('data: [DONE]\n\n');
    res.end();
  } catch (e) {
    res.write(`data: ${JSON.stringify({ error: e.message })}\n\n`);
    res.end();
  }
});

// ── Agent War Room ────────────────────────────────────────────────────────────
// Three specialized agents each build a risk register, then a Synthesis agent
// selects the best rows and produces the final optimal register.

app.post('/api/audits/:id/war-room', async (req, res) => {
  const audits = readAudits();
  const audit  = audits.find(a => a.id === req.params.id);
  if (!audit) return res.status(404).json({ error: 'Not found' });

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');

  const write = (obj) => res.write(`data: ${JSON.stringify(obj)}\n\n`);

  const disc = audit.discovery || {};
  const plan = audit.planning  || {};

  const ctx = [
    `Audit Title: ${audit.title || 'Untitled'}`,
    disc['disc-entity']        && `Entity / Auditee: ${disc['disc-entity']}`,
    disc['disc-audit-type']    && `Audit Type: ${disc['disc-audit-type']}`,
    disc['disc-rationale']     && `Why we are auditing this: ${disc['disc-rationale']}`,
    disc['disc-background']    && `Background: ${disc['disc-background']}`,
    disc['disc-scope-in']      && `In Scope: ${disc['disc-scope-in']}`,
    disc['disc-scope-out']     && `Out of Scope: ${disc['disc-scope-out']}`,
    disc['disc-risk-areas']    && `Known risk areas: ${disc['disc-risk-areas']}`,
    disc['disc-prior-history'] && `Prior audit history: ${disc['disc-prior-history']}`,
    disc['disc-systems']       && `Key systems: ${disc['disc-systems']}`,
    plan['plan-objectives']    && `Audit objectives: ${plan['plan-objectives']}`,
    plan['plan-testing-approach'] && `Testing approach: ${plan['plan-testing-approach']}`,
  ].filter(Boolean).join('\n');

  const PROC_INSTRUCTIONS = `For every in-scope risk include a procedure field with this exact 4-step format:
"Step 1 — Walkthrough: Conduct a walkthrough with [role] to understand the process and confirm control design. Perform a test of one.
Step 2 — Occurrence: Obtain [specific evidence type] to confirm the control executed during the audit period. Sample per VIA table.
Step 3 — Content Review: Review [specific documents] to confirm content is adequate and current.
Step 4 — Recalibration: Review the process used to update/recalibrate [materials] to confirm it reflects current risk conditions."`;

  const ROW_TOOL = {
    name: 'submit_risk_register',
    description: 'Submit the risk register rows you generated.',
    input_schema: {
      type: 'object',
      properties: {
        rows: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              riskId:     { type: 'string' },
              desc:       { type: 'string' },
              frost:      { type: 'string' },
              likelihood: { type: 'string', enum: ['Likely','Possible','Unlikely'] },
              impact:     { type: 'string', enum: ['High','Medium','Low'] },
              controls:   { type: 'string' },
              ctype:      { type: 'string' },
              freq:       { type: 'string' },
              inscope:    { type: 'string', enum: ['Yes','No'] },
              topical:    { type: 'string' },
              procedure:  { type: 'string' },
            },
            required: ['riskId','desc','frost','likelihood','impact'],
          },
        },
      },
      required: ['rows'],
    },
  };

  const AGENTS = [
    {
      label: 'Agent Alpha · Financial & Operational',
      emoji: '🔴',
      focus: `You are Agent Alpha. Focus ONLY on FINANCIAL (F) and OPERATIONAL (O) FROST risks.
Identify 5–8 specific risks in: billing, financial reporting, fraud, purchasing, cash management, process management, people management, quality control, supplier management, segregation of duties, and inventory for this exact engagement.
${PROC_INSTRUCTIONS}`,
    },
    {
      label: 'Agent Beta · Regulatory & Technology',
      emoji: '🔵',
      focus: `You are Agent Beta. Focus ONLY on REGULATORY/LEGAL (R) and TECHNOLOGY/INFORMATION SECURITY (T) FROST risks.
Identify 5–8 specific risks in: compliance, privacy, sensitive data, access control, change management, vulnerability management, data management, cybersecurity, and software development for this exact engagement.
${PROC_INSTRUCTIONS}`,
    },
    {
      label: 'Agent Gamma · Strategic & Cross-Functional',
      emoji: '🟡',
      focus: `You are Agent Gamma. Focus ONLY on STRATEGIC/REPUTATIONAL (S) FROST risks and cross-cutting missing controls.
Identify 4–7 specific risks in: governance, corporate culture, third-party risk, business continuity, crisis management, performance, goals & priorities, and any missing control gaps that span multiple areas for this exact engagement.
${PROC_INSTRUCTIONS}`,
    },
  ];

  const versions = [];

  for (const agent of AGENTS) {
    write({ status: `${agent.emoji} ${agent.label} generating…` });
    try {
      const resp = await client.messages.create({
        model: 'claude-sonnet-4-6',
        max_tokens: 4096,
        tools: [ROW_TOOL],
        tool_choice: { type: 'any' },
        messages: [{
          role: 'user',
          content: `${agent.focus}\n\nENGAGEMENT CONTEXT:\n${ctx}\n\nGenerate your risk register rows now using the submit_risk_register tool. Use exact VIA format for risk IDs, FROST sub-categories, and risk descriptions.`,
        }],
      });
      const block = resp.content.find(b => b.type === 'tool_use');
      const rows  = block?.input?.rows || [];
      versions.push({ label: agent.label, rows });
      write({ status: `${agent.emoji} ${agent.label} — ${rows.length} risks identified ✓` });
    } catch(e) {
      write({ status: `${agent.emoji} ${agent.label} error: ${e.message}`, error: true });
      versions.push({ label: agent.label, rows: [] });
    }
  }

  write({ status: `⚖️ Synthesis agent analyzing all three versions…` });

  try {
    const allJson = versions.map((v, i) =>
      `=== ${v.label} ===\n${JSON.stringify(v.rows, null, 2)}`
    ).join('\n\n');

    const synthResp = await client.messages.create({
      model: 'claude-sonnet-4-6',
      max_tokens: 4096,
      tools: [{ ...ROW_TOOL, name: 'submit_best_register', description: 'Submit the final synthesized risk register.' }],
      tool_choice: { type: 'any' },
      messages: [{
        role: 'user',
        content: `You are the Synthesis Agent. Three agents generated a risk register for this audit. Analyze all versions, eliminate duplicates, select the strongest entries, and produce the optimal final register.\n\nENGAGEMENT CONTEXT:\n${ctx}\n\nAGENT OUTPUTS:\n${allJson}\n\nInstructions:\n1. Keep unique, high-quality entries from each agent — don't discard good risks just because they overlap in category\n2. For true duplicates, merge them keeping the best-written version\n3. Re-assign Risk IDs sequentially: A.R1, A.R2, A.R3, etc.\n4. Ensure every in-scope row has a complete 4-step procedure (Walkthrough → Occurrence → Content Review → Recalibration)\n5. Ensure all risk descriptions follow VIA (DESCRIPTION)/(ROOT CAUSE)/(IMPACT) format\n\nSubmit the final register using submit_best_register.`,
      }],
    });
    const synthBlock = synthResp.content.find(b => b.type === 'tool_use');
    const finalRows  = synthBlock?.input?.rows || versions.flatMap(v => v.rows);
    write({ status: `✅ War Room complete — ${finalRows.length} optimal risks selected`, result: finalRows });
  } catch(e) {
    const fallback = versions.flatMap(v => v.rows);
    write({ status: `⚠️ Synthesis error — returning ${fallback.length} raw risks`, result: fallback, error: true });
  }

  res.write('data: [DONE]\n\n');
  res.end();
});

// ── Document Validator ────────────────────────────────────────────────────────

const multer = require('multer');
const upload = multer({ storage: multer.memoryStorage(), limits: { fileSize: 25 * 1024 * 1024 } });

async function extractDocumentText(buffer, mimetype, originalname) {
  const ext = path.extname(originalname).toLowerCase();
  if (mimetype.startsWith('text/') || ['.csv','.txt','.tsv','.log'].includes(ext)) {
    return buffer.toString('utf8');
  }
  if (mimetype === 'application/pdf' || ext === '.pdf') {
    const pdfParse = require('pdf-parse');
    const data = await pdfParse(buffer);
    return data.text;
  }
  if (mimetype.includes('wordprocessingml') || mimetype.includes('msword') || ext === '.docx' || ext === '.doc') {
    const mammoth = require('mammoth');
    const result  = await mammoth.extractRawText({ buffer });
    return result.value;
  }
  if (mimetype.includes('spreadsheetml') || mimetype.includes('excel') || ['.xlsx','.xls'].includes(ext)) {
    const XLSX = require('xlsx');
    const wb   = XLSX.read(buffer, { type: 'buffer' });
    return wb.SheetNames.map(n => `=== Sheet: ${n} ===\n` + XLSX.utils.sheet_to_csv(wb.Sheets[n])).join('\n\n');
  }
  throw new Error(`Unsupported file type: ${ext || mimetype}`);
}

const VIA_SYSTEM_PROMPT = `You are a senior Verizon Internal Audit (VIA) specialist. Your role is to evaluate documents against VIA audit methodology to assess their value as audit evidence or support materials.

VIA METHODOLOGY CRITERIA:
• 6 Audit Phases: Pre-Discovery → Discovery → Planning → Fieldwork → Reporting → Wrap-Up → Follow-Up
• FROST Risk Framework: Financial (F), Regulatory/Compliance (R), Operational (O), Strategic (S), Technology & Information Security (T)
• Issue Types: PI (Priority Issue — high impact, requires MAP), IO (Informational Observation — low risk, advisory), RI (Risk Item — watch item, not yet an issue)
• Control Evidence Standard (5 Ws): WHO performed the control, WHAT action was taken, WHEN it occurred, WHERE it applied, WHY it ensures the control objective
• Sampling Standards (IIA): Evidence of population completeness and accuracy (C&A) required before sampling; sample must be representative
• Risk Register: Each control should have Likelihood (Likely/Possible/Unlikely), Impact (High/Medium/Low), Control Type (Preventative/Detective/Corrective × Manual/Automated/Hybrid), Frequency
• Topical Requirements: IIA standards mapped as Cyber.Gov, Cyber.Risk, Cyber.Ctrl, TP.Gov, TP.Risk, TP.Ctrl series
• IPE (Information Provided by Entity): All data from the auditee must be validated for completeness and accuracy before use in testing
• Management Action Plans (MAPs): Required for all PI and IO findings; must include owner, due date, and remediation steps
• PIFU: Post-Implementation Follow-Up for completed actions
• SOR: Statement of Responsibilities — annual certification

DOCUMENT ASSESSMENT CRITERIA:
1. Evidence Sufficiency — enough detail to support conclusions and withstand peer review
2. Completeness (C&A) — covers the full control population and time period
3. Accuracy & Reliability — data from credible, traceable source; IPE validated
4. Relevance — directly addresses a control objective, risk, or FROST category
5. Timeliness — current for the audit period being tested
6. Authorization Chain — proper approvals, sign-offs, and management ownership visible
7. Auditability / IPE — traceable to source system; auditor can independently validate
8. Control Effectiveness — demonstrates the control actually operated as designed

Always be specific. Reference exact VIA criteria. Identify which FROST category and audit phase the document best supports.`;

app.post('/api/validate-document', upload.single('file'), async (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file provided' });

  res.setHeader('Content-Type', 'text/event-stream');
  res.setHeader('Cache-Control', 'no-cache');
  res.setHeader('Connection', 'keep-alive');

  const emit = (obj) => res.write(`data: ${JSON.stringify(obj)}\n\n`);

  try {
    emit({ type: 'status', message: `Extracting content from ${req.file.originalname}…` });

    let docText;
    try {
      docText = await extractDocumentText(req.file.buffer, req.file.mimetype, req.file.originalname);
    } catch (e) {
      emit({ type: 'error', message: `Could not read file: ${e.message}` });
      res.write('data: [DONE]\n\n'); return res.end();
    }

    const MAX = 80000;
    const wasTruncated = docText.length > MAX;
    if (wasTruncated) docText = docText.slice(0, MAX) + '\n\n[… document truncated for analysis …]';

    emit({ type: 'status', message: 'Analyzing against VIA audit criteria…' });

    const { auditPhase = '', auditType = '', controlArea = '' } = req.body;
    const ctx = [
      auditPhase   && `Audit Phase: ${auditPhase}`,
      auditType    && `Audit Type: ${auditType}`,
      controlArea  && `Control / Risk Area: ${controlArea}`,
    ].filter(Boolean).join('\n');

    const userPrompt = `Evaluate this document as potential audit evidence or support material for a VIA audit.
${ctx ? `\nAudit Context Provided:\n${ctx}\n` : ''}
File: ${req.file.originalname} (${(req.file.size / 1024).toFixed(1)} KB)${wasTruncated ? ' — truncated for analysis' : ''}

DOCUMENT CONTENT:
${docText}

Provide a structured assessment using EXACTLY these section headers:

## Overall Assessment
State one of: ✅ SUITABLE FOR AUDIT USE | ⚠️ PARTIALLY SUITABLE | ❌ NOT SUITABLE — then 2–3 sentences summarizing why.

## What This Document Is
1–2 sentences describing the document type, apparent source system, and purpose.

## Audit Utility
Which audit phase(s) this supports (Discovery / Planning / Fieldwork / Reporting / Wrap-Up / Follow-Up), which FROST category it falls under, and what specific controls or risk areas it evidences.

## ✅ Strengths
Bullet list — what makes this document strong for audit use. Reference specific VIA criteria (completeness, accuracy, authorization, timeliness, IPE, control effectiveness, etc.).

## ⚠️ Gaps & Limitations
Bullet list — what is missing, weak, or potentially problematic. Be specific about which VIA criteria are not met and what an auditor should watch for.

## VIA Criteria Scorecard
Rate each item with ✅ Met | ⚠️ Partial | ❌ Not Met and a brief note:
- **Evidence Sufficiency:**
- **Completeness (C&A):**
- **Accuracy & Reliability:**
- **Relevance to Control Objective:**
- **Timeliness:**
- **Authorization / Approval Chain:**
- **Auditability / IPE:**
- **Control Effectiveness Demonstrated:**

## Recommendations
Numbered list of specific actions the auditor should take — what to request additionally, how to use this document in testing, or what caveats to document.`;

    const stream = await client.messages.stream({
      model: 'claude-sonnet-4-6',
      max_tokens: 2500,
      system: VIA_SYSTEM_PROMPT,
      messages: [{ role: 'user', content: userPrompt }],
    });

    for await (const chunk of stream) {
      if (chunk.type === 'content_block_delta' && chunk.delta?.text) {
        emit({ type: 'delta', text: chunk.delta.text });
      }
    }
  } catch (err) {
    emit({ type: 'error', message: err.message });
  }

  res.write('data: [DONE]\n\n');
  res.end();
});

// ── Presentation routes ───────────────────────────────────────────────────────

app.get('/api/clasp/status', (req, res) => {
  res.json({ connected: !!getClaspCredentials() });
});

app.post('/api/presentation/pptx', async (req, res) => {
  const { slides: imgs } = req.body;
  if (!Array.isArray(imgs) || !imgs.length)
    return res.status(400).json({ error: 'No slides provided' });
  try {
    const pptx = new PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';
    for (const imgData of imgs) {
      const slide = pptx.addSlide();
      slide.addImage({ data: imgData, x: 0, y: 0, w: '100%', h: '100%' });
    }
    const buffer = await pptx.write({ outputType: 'nodebuffer' });
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
    res.setHeader('Content-Disposition', 'attachment; filename="Verizon-Audit-AI-Presentation.pptx"');
    res.send(buffer);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.post('/api/presentation/gslides', async (req, res) => {
  const claspCreds = getClaspCredentials();
  if (!claspCreds)
    return res.status(401).json({ error: 'Not connected to Google', needsLogin: true });
  try {
    const { google } = require('googleapis');
    const { Readable } = require('stream');
    const pptx       = buildEditablePptx();
    const pptxBuffer = await pptx.write({ outputType: 'nodebuffer' });
    const oauthClient = new google.auth.OAuth2(claspCreds.clientId, claspCreds.clientSecret, 'http://localhost');
    oauthClient.setCredentials(claspCreds.token);
    const { credentials } = await oauthClient.refreshAccessToken();
    oauthClient.setCredentials(credentials);
    const drive    = google.drive({ version: 'v3', auth: oauthClient });
    const response = await drive.files.create({
      requestBody: { name: 'Verizon Audit AI Presentation', mimeType: 'application/vnd.google-apps.presentation' },
      media: { mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation', body: Readable.from(pptxBuffer) },
      fields: 'id,webViewLink',
    });
    res.json({ url: response.data.webViewLink, id: response.data.id });
  } catch (err) {
    const needsLogin = /invalid_grant|expired|revoked/i.test(err.message);
    res.status(500).json({ error: err.message, needsLogin });
  }
});

app.listen(PORT, () => {
  console.log(`\n  Audit AI  →  http://localhost:${PORT}\n`);
});
