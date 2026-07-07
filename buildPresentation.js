'use strict';
const PptxGenJS = require('pptxgenjs');

module.exports = function buildEditablePptx() {
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE'; // 13.33" × 7.5"

  // ── Palette ──────────────────────────────────────────────────────────────────
  const BG      = '0F0E17';
  const RED     = 'CD040B';
  const WHITE   = 'FFFFFF';
  const MUTED   = '8888AA';
  const DIM     = '444466';
  const CARD    = '16213E';
  const CBORDER = '252545';
  const FONT    = 'Calibri';
  const W = 13.33, H = 7.5;

  // ── Shared helpers ────────────────────────────────────────────────────────────
  function accentBar(s) {
    s.addShape('rect', { x: 0, y: 0, w: 0.07, h: H, fill: { color: RED }, line: { type: 'none' } });
  }
  function eyebrow(s, text, y = 0.4) {
    s.addText(text.toUpperCase(), { x: 0.9, y, w: 11, h: 0.28, fontSize: 9, bold: true, color: RED, fontFace: FONT, charSpacing: 2 });
  }
  function heading(s, text, y = 0.72) {
    s.addText(text, { x: 0.9, y, w: 11.5, h: 0.62, fontSize: 26, bold: true, color: WHITE, fontFace: FONT });
  }
  function rule(s, y = 1.4) {
    s.addShape('rect', { x: 0.9, y, w: 1.1, h: 0.04, fill: { color: RED }, line: { type: 'none' } });
  }
  function sub(s, text, y = 1.52, maxW = 11.4) {
    s.addText(text, { x: 0.9, y, w: maxW, h: 0.52, fontSize: 11.5, color: MUTED, fontFace: FONT, wrap: true });
  }
  function card(s, x, y, w, h, opts = {}) {
    s.addShape('roundRect', {
      x, y, w, h,
      fill: { color: opts.fill || CARD },
      line: { color: opts.border || CBORDER, width: 0.5 },
      rectRadius: opts.r != null ? opts.r : 0.08,
    });
  }
  function pill(s, text, x, y, w, textColor = RED) {
    s.addShape('roundRect', { x, y, w, h: 0.36, fill: { color: RED, transparency: 80 }, line: { color: RED, width: 0.5 }, rectRadius: 0.18 });
    s.addText(text, { x, y, w, h: 0.36, fontSize: 9, bold: true, color: textColor, fontFace: FONT, align: 'center', valign: 'middle' });
  }

  // ──────────────────────────────────────────────────────────────────────────────
  // Slide 1 — Title
  // ──────────────────────────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.background = { color: BG };
    s.addShape('rect', { x: 0, y: 0,      w: W,    h: 0.06, fill: { color: RED }, line: { type: 'none' } });
    s.addShape('rect', { x: 0, y: H-0.06, w: W,    h: 0.06, fill: { color: RED }, line: { type: 'none' } });

    s.addText('VERIZON — INTERNAL AUDIT', {
      x: 0, y: 0.25, w: W, h: 0.28, fontSize: 9, bold: true, color: RED,
      fontFace: FONT, charSpacing: 3, align: 'center',
    });
    s.addText('CFO Presentation · June 2026', {
      x: 0, y: 1.6, w: W, h: 0.3, fontSize: 11, color: MUTED, fontFace: FONT, align: 'center',
    });
    s.addText('Automating the', {
      x: 0, y: 2.05, w: W, h: 0.7, fontSize: 46, bold: true, color: WHITE, fontFace: FONT, align: 'center',
    });
    s.addText('Internal Audit', {
      x: 0, y: 2.72, w: W, h: 0.7, fontSize: 46, bold: true, color: RED, fontFace: FONT, align: 'center',
    });
    s.addText('Process', {
      x: 0, y: 3.38, w: W, h: 0.7, fontSize: 46, bold: true, color: WHITE, fontFace: FONT, align: 'center',
    });
    s.addText('Leveraging artificial intelligence to transform how Verizon\'s audit team documents, executes, and delivers audit engagements', {
      x: 2.2, y: 4.3, w: 8.93, h: 0.65, fontSize: 13, color: MUTED, fontFace: FONT, align: 'center', wrap: true,
    });
    s.addText('Presented by  Verizon Audit Intern     |     Audience  Chief Financial Officer     |     Date  June 2026', {
      x: 0, y: 6.75, w: W, h: 0.32, fontSize: 10, color: DIM, fontFace: FONT, align: 'center',
    });
  }

  // ──────────────────────────────────────────────────────────────────────────────
  // Slide 2 — Agenda
  // ──────────────────────────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.background = { color: BG };
    accentBar(s);
    eyebrow(s, 'Overview', 0.4);
    heading(s, 'What We\'ll Cover Today', 0.72);
    rule(s, 1.4);

    const items = [
      ['1','The Current State','How audits are managed manually today'],
      ['2','Key Inefficiencies','Time, cost, and risk of the manual process'],
      ['3','Our Solution','The Audit AI Platform we built'],
      ['4','Key Features','AI, workflows, exports, and tracking'],
      ['5','Business Value','ROI, time savings, and risk reduction'],
      ['6','Next Steps','Roadmap for full deployment'],
    ];
    items.forEach(([n, title, desc], i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = col === 0 ? 0.9 : 7.05;
      const y = 1.62 + row * 1.68;
      const cw = 5.78;
      card(s, x, y, cw, 1.48);
      s.addShape('ellipse', { x: x+0.18, y: y+0.3, w: 0.55, h: 0.55, fill: { color: RED }, line: { type: 'none' } });
      s.addText(n, { x: x+0.18, y: y+0.3, w: 0.55, h: 0.55, fontSize: 13, bold: true, color: WHITE, fontFace: FONT, align: 'center', valign: 'middle' });
      s.addText(title, { x: x+0.9, y: y+0.18, w: cw-1.05, h: 0.38, fontSize: 13, bold: true, color: WHITE, fontFace: FONT });
      s.addText(desc,  { x: x+0.9, y: y+0.6,  w: cw-1.05, h: 0.72, fontSize: 10.5, color: MUTED, fontFace: FONT, wrap: true });
    });
  }

  // ──────────────────────────────────────────────────────────────────────────────
  // Slide 3 — Current State
  // ──────────────────────────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.background = { color: BG };
    accentBar(s);
    eyebrow(s, 'Current State', 0.35);
    heading(s, 'The Scale of Verizon\'s Audit Function', 0.65);
    rule(s, 1.32);
    sub(s, 'Large enterprises like Verizon operate a high-volume audit program spanning finance, operations, compliance, and technology — all managed largely by hand.', 1.43);

    const stats = [
      ['6+',   'Distinct phases in every audit engagement, each requiring extensive documentation'],
      ['60%',  'Of an auditor\'s time is spent on documentation rather than actual analysis'],
      ['4–8',  'Weeks to complete a single audit from discovery through follow-up'],
      ['100s', 'Of audits conducted annually across Verizon\'s business units and geographies'],
    ];
    stats.forEach(([big, label], i) => {
      const x = 0.9 + i * 2.97;
      card(s, x, 2.1, 2.75, 5.05, { border: RED });
      s.addText(big,  { x, y: 2.45, w: 2.75, h: 0.95, fontSize: 40, bold: true, color: RED,   fontFace: FONT, align: 'center' });
      s.addText(label,{ x: x+0.15, y: 3.55, w: 2.45, h: 3.35, fontSize: 11, color: MUTED, fontFace: FONT, wrap: true, align: 'center' });
    });
  }

  // ──────────────────────────────────────────────────────────────────────────────
  // Slide 4 — Pain Points
  // ──────────────────────────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.background = { color: BG };
    accentBar(s);
    eyebrow(s, 'The Problem', 0.35);
    heading(s, 'Inefficiencies in the Manual Audit Process', 0.65);
    rule(s, 1.32);

    const pains = [
      ['Time-Intensive Documentation','Auditors manually write risk registers, work steps, and findings from scratch for every engagement, consuming hours that should be spent on analysis.'],
      ['No Standardized Templates','Each auditor formats deliverables differently — inconsistent quality, structure, and terminology makes peer review and cross-team comparison difficult.'],
      ['Fragmented Coordination','Draft reports circulate via email chains and shared drives. Version control breaks down and real-time audit status is impossible to see at a glance.'],
      ['Manual Entry Errors','Copy-pasting between Word, Excel, and email introduces errors into risk ratings, issue classifications, and Management Action Plans (MAPs).'],
      ['No Centralized Visibility','Leadership has no single dashboard showing which audits are in progress, complete, or where bottlenecks exist in the audit pipeline.'],
      ['Slow Report Distribution','Formatting final reports takes hours. Sending to stakeholders requires manual PDF conversion, email composition, and follow-up tracking.'],
    ];
    pains.forEach(([title, desc], i) => {
      const col = i % 2, row = Math.floor(i / 2);
      const x = col === 0 ? 0.9 : 7.05;
      const y = 1.5 + row * 1.95;
      const cw = 5.78;
      card(s, x, y, cw, 1.78);
      s.addShape('rect', { x, y, w: 0.22, h: 1.78, fill: { color: RED }, line: { type: 'none' }, rectRadius: 0 });
      s.addText(title, { x: x+0.35, y: y+0.2,  w: cw-0.5, h: 0.38, fontSize: 12.5, bold: true, color: WHITE, fontFace: FONT });
      s.addText(desc,  { x: x+0.35, y: y+0.62, w: cw-0.5, h: 1.0,  fontSize: 10.5, color: MUTED, fontFace: FONT, wrap: true });
    });
  }

  // ──────────────────────────────────────────────────────────────────────────────
  // Slide 5 — Cost
  // ──────────────────────────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.background = { color: BG };
    accentBar(s);
    eyebrow(s, 'Business Impact', 0.35);
    heading(s, 'What Manual Processes Cost the Organization', 0.65);
    rule(s, 1.32);
    sub(s, 'These inefficiencies aren\'t just operational friction — they translate directly into financial and compliance risk.', 1.43);

    const costs = [
      ['~40','hours / audit','Estimated time spent on documentation per engagement — time that could be redirected to higher-value risk analysis'],
      ['High','opportunity cost','Senior auditors spending billable hours on data entry and formatting rather than identifying controls gaps and advising stakeholders'],
      ['Elevated','compliance risk','Inconsistent documentation and missed follow-ups increase exposure to regulatory findings and repeat audit exceptions'],
      ['Delayed','findings delivery','Management doesn\'t receive actionable recommendations in time to remediate issues before the next audit cycle begins'],
    ];
    costs.forEach(([num, unit, desc], i) => {
      const x = 0.9 + i * 2.97;
      card(s, x, 2.1, 2.75, 5.05);
      s.addText(num,  { x, y: 2.45, w: 2.75, h: 0.75, fontSize: 30, bold: true, color: WHITE, fontFace: FONT, align: 'center' });
      s.addText(unit, { x, y: 3.22, w: 2.75, h: 0.35, fontSize: 11, bold: true, color: RED,   fontFace: FONT, align: 'center' });
      s.addShape('rect', { x: x+0.5, y: 3.65, w: 1.75, h: 0.03, fill: { color: CBORDER }, line: { type: 'none' } });
      s.addText(desc, { x: x+0.15, y: 3.75, w: 2.45, h: 3.15, fontSize: 10.5, color: MUTED, fontFace: FONT, wrap: true, align: 'center' });
    });
  }

  // ──────────────────────────────────────────────────────────────────────────────
  // Slide 6 — Solution Intro
  // ──────────────────────────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.background = { color: BG };

    pill(s, '★  Introducing the Solution', 4.67, 0.55, 4.0);
    s.addText('The Audit AI Platform', {
      x: 1.5, y: 1.1, w: 10.33, h: 0.78, fontSize: 36, bold: true, color: WHITE, fontFace: FONT, align: 'center',
    });
    s.addText('A purpose-built web application that automates documentation, standardizes workflows, and integrates AI-assisted content generation across every phase of the audit lifecycle.', {
      x: 1.9, y: 2.0, w: 9.53, h: 0.62, fontSize: 12, color: MUTED, fontFace: FONT, align: 'center', wrap: true,
    });

    const pillars = [
      ['AI-Powered',  'Claude AI generates risk registers, work steps, objectives, and findings — auditors review and refine rather than write from scratch'],
      ['Standardized','Every audit follows the same structured 6-phase workflow with consistent templates, terminology, and output format'],
      ['Centralized', 'All audits, statuses, and deliverables live in one platform — visible to the entire team in real time'],
      ['Fast',        'One-click PDF reports, Gmail integration, and Google Docs export eliminate manual formatting and distribution steps'],
    ];
    pillars.forEach(([title, desc], i) => {
      const x = 0.9 + i * 2.97;
      card(s, x, 2.82, 2.75, 4.35);
      s.addText(title, { x: x+0.15, y: 2.82+0.4,  w: 2.45, h: 0.45, fontSize: 14, bold: true, color: WHITE, fontFace: FONT, align: 'center' });
      s.addShape('rect', { x: x+0.5, y: 2.82+0.95, w: 1.75, h: 0.03, fill: { color: RED }, line: { type: 'none' } });
      s.addText(desc, { x: x+0.15, y: 2.82+1.1, w: 2.45, h: 3.0, fontSize: 10.5, color: MUTED, fontFace: FONT, wrap: true, align: 'center' });
    });
  }

  // ──────────────────────────────────────────────────────────────────────────────
  // Slide 7 — Workflow
  // ──────────────────────────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.background = { color: BG };
    accentBar(s);
    eyebrow(s, 'Feature · Structured Workflow', 0.35);
    heading(s, 'A Standardized 6-Phase Audit Lifecycle', 0.65);
    rule(s, 1.32);
    sub(s, 'Every audit progresses through the same phases with pre-built templates — eliminating the cold-start problem of building documents from scratch.', 1.43);

    const phases = [
      ['01','Discovery', 'Entity background, risk areas, scope, team assignment'],
      ['02','Planning',  'Objectives, risk register, testing approach, resource plan'],
      ['03','Fieldwork', 'Evidence collection, work steps, observations, contacts'],
      ['04','Reporting', 'Issues log, exec summary, audit opinion, vetting'],
      ['05','Wrap Up',   'MAPs, management responses, retrospective, archive'],
      ['06','Follow Up', 'MAP status, remediation tracking, open issues, conclusion'],
    ];
    const cw = 1.98;
    phases.forEach(([num, title, desc], i) => {
      const x = 0.72 + i * 2.02;
      if (i < 5) {
        s.addShape('rect', { x: x+cw, y: 3.1, w: 0.04, h: 0.03, fill: { color: RED }, line: { type: 'none' } });
      }
      card(s, x, 2.1, cw, 5.1);
      s.addText(num,   { x, y: 2.22, w: cw, h: 0.28, fontSize: 9,  bold: true, color: RED,   fontFace: FONT, align: 'center', charSpacing: 1 });
      s.addText(title, { x: x+0.08, y: 2.58, w: cw-0.16, h: 0.42, fontSize: 13, bold: true, color: WHITE, fontFace: FONT, align: 'center' });
      s.addShape('rect', { x: x+0.4, y: 3.07, w: cw-0.8, h: 0.03, fill: { color: RED }, line: { type: 'none' } });
      s.addText(desc, { x: x+0.1, y: 3.18, w: cw-0.2, h: 3.8, fontSize: 10, color: MUTED, fontFace: FONT, wrap: true, align: 'center' });
    });
  }

  // ──────────────────────────────────────────────────────────────────────────────
  // Slide 8 — AI Features
  // ──────────────────────────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.background = { color: BG };
    accentBar(s);
    eyebrow(s, 'Feature · Artificial Intelligence', 0.35);
    heading(s, 'AI-Assisted Content Generation', 0.65);
    rule(s, 1.32);

    const bullets = [
      'Auditors describe the engagement context and AI generates complete risk registers, audit objectives, work steps, and findings drafts',
      'Context-aware prompts built into each phase — one click fills an entire section with a relevant, editable draft',
      'Powered by Anthropic\'s Claude — the same enterprise AI trusted by Fortune 500 companies for sensitive document work',
      'Auditors remain in full control — AI generates suggestions, humans review, edit, and approve all content',
      'Dramatically reduces the most time-consuming part of every audit: populating structured documentation fields',
    ];
    bullets.forEach((text, i) => {
      const y = 1.52 + i * 1.06;
      s.addShape('ellipse', { x: 0.9, y: y+0.1, w: 0.18, h: 0.18, fill: { color: RED }, line: { type: 'none' } });
      s.addText(text, { x: 1.22, y, w: 5.8, h: 0.95, fontSize: 11.5, color: MUTED, fontFace: FONT, wrap: true, valign: 'top' });
    });

    // Chat demo panel
    card(s, 7.35, 1.45, 5.55, 5.72, { fill: '0D1B2A', border: '1E3A5F' });
    s.addText('DEMO', { x: 7.35, y: 1.55, w: 5.55, h: 0.28, fontSize: 8, bold: true, color: '3A6EA5', fontFace: FONT, align: 'center', charSpacing: 3 });

    s.addText('Auditor', { x: 7.55, y: 2.0, w: 2, h: 0.22, fontSize: 9, color: '3A6EA5', fontFace: FONT, bold: true });
    s.addShape('roundRect', { x: 7.55, y: 2.22, w: 4.95, h: 0.92, fill: { color: '1E3A5F' }, line: { type: 'none' }, rectRadius: 0.1 });
    s.addText('"Build a risk register for a livestock operations audit with 5 key risk areas"', {
      x: 7.68, y: 2.3, w: 4.72, h: 0.78, fontSize: 11, color: WHITE, fontFace: FONT, italic: true, wrap: true,
    });

    s.addText('Audit AI', { x: 7.55, y: 3.32, w: 2, h: 0.22, fontSize: 9, color: RED, fontFace: FONT, bold: true });
    s.addShape('roundRect', { x: 7.55, y: 3.55, w: 4.95, h: 1.85, fill: { color: '162032' }, line: { color: '1E3A5F', width: 0.5 }, rectRadius: 0.1 });
    s.addText('Generated 5 risks: Inventory Accuracy, Animal Health Compliance, Feed Procurement, Revenue Controls, Biosecurity — each with likelihood, impact, and recommended audit response ✓', {
      x: 7.68, y: 3.63, w: 4.72, h: 1.68, fontSize: 11, color: WHITE, fontFace: FONT, wrap: true,
    });

    s.addText('Auto-populated into the Planning phase template', {
      x: 7.35, y: 5.6, w: 5.55, h: 0.38, fontSize: 9, color: DIM, fontFace: FONT, align: 'center',
    });
  }

  // ──────────────────────────────────────────────────────────────────────────────
  // Slide 9 — Export
  // ──────────────────────────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.background = { color: BG };
    accentBar(s);
    eyebrow(s, 'Feature · Distribution', 0.35);
    heading(s, 'Instant Export & Stakeholder Distribution', 0.65);
    rule(s, 1.32);
    sub(s, 'What once took hours of manual formatting is now a single click. Reports are consistently branded and immediately shareable.', 1.43);

    const exports = [
      ['PDF Report',       'Professional, branded PDF generated instantly with all phases, tables, risk registers, issues logs, and MAPs formatted automatically.', 'One-click download'],
      ['Gmail Integration','Clicking "Gmail" downloads the PDF and opens a pre-composed Gmail draft with the audit summary in the body — ready to send in seconds.',  'Opens in browser'],
      ['Google Docs Export','The full audit is uploaded directly to Google Drive as a live Google Doc — shareable, commentable, and accessible to the entire team.', 'Saved to Drive'],
    ];
    exports.forEach(([title, desc, tag], i) => {
      const x = 0.9 + i * 4.12;
      const cw = 3.88;
      card(s, x, 2.1, cw, 5.05);
      s.addText(title, { x: x+0.15, y: 2.35, w: cw-0.3, h: 0.45, fontSize: 15, bold: true, color: WHITE, fontFace: FONT, align: 'center' });
      s.addShape('rect', { x: x+0.4, y: 2.88, w: cw-0.8, h: 0.04, fill: { color: RED }, line: { type: 'none' } });
      s.addText(desc, { x: x+0.2, y: 3.0, w: cw-0.4, h: 2.85, fontSize: 11.5, color: MUTED, fontFace: FONT, wrap: true, align: 'center' });
      pill(s, tag, x+0.55, 6.75-0.38, cw-1.1);
    });
  }

  // ──────────────────────────────────────────────────────────────────────────────
  // Slide 10 — Comparison Table
  // ──────────────────────────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.background = { color: BG };
    accentBar(s);
    eyebrow(s, 'Side-by-Side', 0.3);
    heading(s, 'Manual Process vs. Audit AI Platform', 0.58);
    rule(s, 1.26);

    const rows = [
      ['Audit Activity',               'Manual Process',                   'Audit AI Platform'],
      ['Document a single audit phase', '2–4 hours of writing',             'Minutes with AI assistance'],
      ['Build a risk register',         'Research + manual entry',           'AI-generated in seconds'],
      ['Generate final report',         'Hours of formatting in Word',       'Instant professional PDF'],
      ['Distribute to stakeholders',    'Email chains, manual attach',       'One-click Gmail + Google Docs'],
      ['Track audit pipeline status',   'Spreadsheet or email thread',       'Live dashboard with status flags'],
      ['Consistency across audits',     'Varies by auditor',                 'Enforced by standardized templates'],
      ['Follow-up / MAP tracking',      'Separate tracker spreadsheet',      'Built into the platform'],
    ];
    const tX = 0.9, tY = 1.42, tW = 11.5;
    const cols = [4.1, 3.7, 3.7];
    const rH = 0.63;

    rows.forEach(([act, manual, ai], r) => {
      const isHdr = r === 0;
      const y = tY + r * rH;
      const bg = isHdr ? '1A1A38' : r % 2 === 0 ? '0F0E17' : '111128';
      s.addShape('rect', { x: tX, y, w: tW, h: rH, fill: { color: bg }, line: { color: CBORDER, width: 0.5 } });
      s.addText(act,    { x: tX+0.12, y, w: cols[0]-0.2, h: rH, fontSize: 11, bold: isHdr, color: isHdr ? WHITE : MUTED, fontFace: FONT, valign: 'middle' });
      s.addText(isHdr ? '⚠  '+manual : manual, { x: tX+cols[0]+0.12, y, w: cols[1]-0.2, h: rH, fontSize: 11, bold: isHdr, color: 'FF9999', fontFace: FONT, valign: 'middle', align: isHdr ? 'center' : 'left' });
      s.addText(isHdr ? '✦  '+ai : ai,         { x: tX+cols[0]+cols[1]+0.12, y, w: cols[2]-0.2, h: rH, fontSize: 11, bold: isHdr, color: '86EFAC', fontFace: FONT, valign: 'middle', align: isHdr ? 'center' : 'left' });
    });
  }

  // ──────────────────────────────────────────────────────────────────────────────
  // Slide 11 — ROI
  // ──────────────────────────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.background = { color: BG };
    accentBar(s);
    eyebrow(s, 'Business Value', 0.35);
    heading(s, 'Return on Investment', 0.65);
    rule(s, 1.32);
    sub(s, 'Even conservative estimates point to significant recoverable capacity across the audit function.', 1.43);

    const roi = [
      ['~60%',        'Documentation time recovered',      'AI handles the first draft. Auditors spend time reviewing and refining — not writing from a blank page.'],
      ['Days→Hours',  'Report-to-distribution cycle',      'Formatting and distributing a final audit report drops from a full day\'s work to a single click.'],
      ['↓ Risk',      'Compliance and error exposure',     'Standardized templates and structured fields eliminate the class of errors caused by manual copy-paste.'],
      ['↑ Scale',     'More audits, same headcount',       'Recovered capacity means the same team can cover more of the audit universe — or go deeper on high-risk areas.'],
    ];
    roi.forEach(([big, label, desc], i) => {
      const x = 0.9 + i * 2.97;
      card(s, x, 2.1, 2.75, 5.05);
      s.addText(big,   { x: x+0.1, y: 2.38, w: 2.55, h: 0.78, fontSize: 26, bold: true, color: RED,   fontFace: FONT, align: 'center', wrap: true });
      s.addText(label, { x: x+0.15, y: 3.22, w: 2.45, h: 0.55, fontSize: 11, bold: true, color: WHITE, fontFace: FONT, align: 'center', wrap: true });
      s.addShape('rect', { x: x+0.5, y: 3.85, w: 1.75, h: 0.03, fill: { color: CBORDER }, line: { type: 'none' } });
      s.addText(desc, { x: x+0.15, y: 3.95, w: 2.45, h: 2.95, fontSize: 10.5, color: MUTED, fontFace: FONT, wrap: true, align: 'center' });
    });
  }

  // ──────────────────────────────────────────────────────────────────────────────
  // Slide 12 — Next Steps
  // ──────────────────────────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.background = { color: BG };
    accentBar(s);
    eyebrow(s, 'Roadmap', 0.35);
    heading(s, 'Next Steps', 0.65);
    rule(s, 1.32);

    const steps = [
      ['1','Prototype Built & Functional', 'Full 6-phase platform with AI, PDF, Gmail, Google Docs export, and audit management is live and tested', 'COMPLETE', RED,    'CD040B'],
      ['2','Pilot with Internal Audit Team','Run 2–3 real engagements through the platform with active auditors to validate time savings and identify gaps', 'UP NEXT',  '3B82F6','3B82F6'],
      ['3','Integrate with Existing Systems','Connect to Verizon\'s audit management, GRC, and ticketing systems for seamless data flow', 'PLANNED', '555577','555577'],
      ['4','Full Deployment & Training',    'Roll out to the full audit team with onboarding documentation, admin controls, and ongoing AI model tuning', 'PLANNED',  '555577','555577'],
    ];
    steps.forEach(([n, title, desc, badge, dotColor, badgeColor], i) => {
      const y = 1.52 + i * 1.42;
      card(s, 0.9, y, 11.5, 1.25);
      s.addShape('ellipse', { x: 1.1, y: y+0.3, w: 0.58, h: 0.58, fill: { color: dotColor }, line: { type: 'none' } });
      s.addText(n, { x: 1.1, y: y+0.3, w: 0.58, h: 0.58, fontSize: 14, bold: true, color: WHITE, fontFace: FONT, align: 'center', valign: 'middle' });
      s.addText(title, { x: 1.88, y: y+0.12, w: 7.7, h: 0.42, fontSize: 13.5, bold: true, color: WHITE, fontFace: FONT });
      s.addText(desc,  { x: 1.88, y: y+0.58, w: 7.7, h: 0.55, fontSize: 11, color: MUTED, fontFace: FONT, wrap: true });
      s.addShape('roundRect', { x: 10.05, y: y+0.33, w: 2.05, h: 0.42, fill: { color: badgeColor, transparency: 85 }, line: { color: badgeColor, width: 0.5 }, rectRadius: 0.21 });
      s.addText(badge, { x: 10.05, y: y+0.33, w: 2.05, h: 0.42, fontSize: 9.5, bold: true, color: badgeColor, fontFace: FONT, align: 'center', valign: 'middle', charSpacing: 1 });
    });
  }

  // ──────────────────────────────────────────────────────────────────────────────
  // Slide 13 — Thank You
  // ──────────────────────────────────────────────────────────────────────────────
  {
    const s = pptx.addSlide();
    s.background = { color: BG };
    s.addShape('rect', { x: 0, y: 0,      w: W, h: 0.06, fill: { color: RED }, line: { type: 'none' } });
    s.addShape('rect', { x: 0, y: H-0.06, w: W, h: 0.06, fill: { color: RED }, line: { type: 'none' } });
    s.addText('VERIZON — INTERNAL AUDIT', {
      x: 0, y: 0.25, w: W, h: 0.28, fontSize: 9, bold: true, color: RED, fontFace: FONT, charSpacing: 3, align: 'center',
    });
    s.addText('Thank You', {
      x: 0, y: 1.9, w: W, h: 1.15, fontSize: 62, bold: true, color: WHITE, fontFace: FONT, align: 'center',
    });
    s.addShape('rect', { x: 5.67, y: 3.25, w: 2.0, h: 0.06, fill: { color: RED }, line: { type: 'none' } });
    s.addText('The manual audit process is a solvable problem. With AI-assisted documentation, structured workflows, and one-click distribution — we can give the audit team back the time to do the work that actually matters.', {
      x: 1.5, y: 3.45, w: 10.33, h: 1.3, fontSize: 14, color: MUTED, fontFace: FONT, align: 'center', wrap: true,
    });
    s.addText('Questions & Discussion', {
      x: 0, y: 5.1, w: W, h: 0.38, fontSize: 13, color: DIM, fontFace: FONT, align: 'center', charSpacing: 2,
    });
  }

  return pptx;
};
