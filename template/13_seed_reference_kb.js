/**
 * 13_seed_reference_kb.js
 *
 * Fills two gaps found by auditing the live Agents/KB sheets (2026-09-09):
 *  1. 12 of 15 seeded agents (12_seed_all_risk_agents.js) had zero KB rows,
 *     even though their personas cite doc filenames (01-nist-csf.md,
 *     06-identity-basics.md, 08-its-policy-catalog.md, etc.) that never
 *     existed in the KB tab.
 *  2. Every agent's 'tone' column (added in the v1.2.0 migration) was blank.
 *
 * All three functions here are upserts keyed by a stable id / slug, safe to
 * re-run: seedReferenceKB_ updates-in-place by KB row id, and the tone/skill
 * backfills only touch cells that are still blank / still at the migration
 * default, so manual admin edits made after running this once are preserved.
 */

function _isAdminUserForSeed() {
  return _isAdminUser_();
}

function _upsertKbRow_(id, slug, title, body, tags) {
  const sheet = _getOrCreateTab_('KB', ['id', 'slug', 'title', 'body', 'tags']);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === id) {
      sheet.getRange(i + 1, 1, 1, 5).setValues([[id, slug, title, body, tags]]);
      return 'updated';
    }
  }
  sheet.appendRow([id, slug, title, body, tags]);
  return 'inserted';
}

// slug 'shared' = visible to every agent (see _getKbForSlug_ in 20_chat.js).
var REFERENCE_KB_DOCS = [
  ['ref-01-nist-csf', 'shared', '01-nist-csf', 'NIST Cybersecurity Framework 2.0 — six functions: GOVERN (set risk strategy, roles, policy — the newest 2.0 function); IDENTIFY (asset inventory, risk assessment, understand what needs protecting); PROTECT (access control, awareness training, data security, safeguards); DETECT (continuous monitoring, anomaly detection); RESPOND (incident response plan, communications, mitigation); RECOVER (restore capabilities, lessons learned, post-incident review). Use CSF functions to structure any security recommendation: name which function(s) the control belongs to.', 'nist,framework,csf'],
  ['ref-02-cis-controls', 'shared', '02-cis-controls', 'CIS Critical Security Controls v8 — 18 controls grouped by Implementation Group (IG1 = basic cyber hygiene for every org, IG2 = adds controls for orgs with more risk/resources, IG3 = mature/high-risk orgs). Key controls: 1 Inventory of Enterprise Assets, 2 Inventory of Software Assets, 3 Data Protection, 4 Secure Configuration, 5 Account Management, 6 Access Control Management, 7 Continuous Vulnerability Management, 8 Audit Log Management, 9 Email/Browser Protections, 10 Malware Defenses, 11 Data Recovery, 12 Network Infrastructure Management, 13 Network Monitoring/Defense, 14 Security Awareness Training, 15 Service Provider Management, 16 Application Software Security, 17 Incident Response Management, 18 Penetration Testing. When recommending a control, name the CIS Control number and Implementation Group.', 'cis,controls,framework'],
  ['ref-03-k12-data', 'shared', '03-k12-data', 'K-12 student data protection basics. FERPA: student education records are confidential; disclosure requires consent or a documented exception (e.g., school official with legitimate interest); parents/eligible students (18+) can inspect and request correction of records. COPPA: applies to online services collecting personal info from children under 13 — requires verifiable parental consent and data-minimization. CIPA: schools receiving E-Rate discounts must filter/monitor internet access for minors. LAUSD-specific: classify all student data per BUL-999.16 before storing in any new system; any new tool touching student data needs a documented risk/vendor review before deployment (see 04-vendor-arb-checklist).', 'ferpa,coppa,cipa,student-data,k12'],
  ['ref-04-vendor-arb-checklist', 'grc', '04-vendor-arb-checklist (full)', 'Architecture Review Board (ARB) / vendor risk checklist: 1) Data classification — what data will the vendor touch, at what sensitivity (per BUL-999.16)? 2) Contract — is there an executed agreement covering data ownership, breach notification, and termination/data-return terms? 3) Security posture — SOC 2 / ISO 27001 report or equivalent evidence on file? 4) Data residency/subprocessors — where is data stored/processed, and by whom downstream? 5) Access model — SSO/MFA support, least-privilege roles? 6) CISO sign-off — required for any vendor touching student data or classified-high data. Do not approve a vendor missing items 1-3 for anything above the lowest data-sensitivity tier.', 'vendor,arb,checklist'],
  ['ref-05-phishing-helpdesk', 'security-helpdesk', '05-phishing-helpdesk', 'Phishing triage for helpdesk: red flags = urgency/fear language, off-domain or lookalike sender/link, unexpected attachment, request for credentials or gift cards, generic greeting on a claimed "personal" message. If a user reports a suspected phish: (1) do not click any link/attachment, (2) have them forward it as an attachment to security@lausd.net or use the report-phish button if available, (3) if they already clicked or entered credentials, treat as a likely compromise — walk them through an immediate password reset via MyLogin and note it for IDM/CTU follow-up. Never ask a user to read back a password or MFA code to "verify" — that request itself is a common phishing/vishing pattern.', 'phish,report,helpdesk'],
  ['ref-06-identity-basics', 'idm', '06-identity-basics', 'Identity & access management fundamentals: least privilege — grant only the access needed for the current role, nothing "just in case". Separation of duties — no single person should both request and approve their own access change. Lifecycle — provision on verified start date/role, review access at role change, deprovision same-day on separation. MFA — required for all accounts with access to sensitive systems or student data; phishing-resistant methods (hardware key, platform authenticator) preferred over SMS where available. Access reviews — periodic recertification of who has access to what, with an owner attesting it is still needed. Never bypass an approval workflow to expedite access, even for urgent requests — escalate urgency instead.', 'identity,iam,mfa,access'],
  ['ref-07-ctu-sources', 'ctu', '07-ctu-sources (full)', 'Primary threat intel sources, in priority order: CISA (advisories + Known Exploited Vulnerabilities catalog — treat KEV listing as high urgency), MSRC (Microsoft security updates/advisories), NVD (CVE details, CVSS scoring), vendor security bulletins for affected products, sector ISACs (e.g., MS-ISAC for K-12/state-local govt) for K-12-specific threat context. Escalation matrix: Critical/KEV-listed + internet-facing = escalate to SOC Manager immediately with patch/block/monitor recommendation; High severity, not actively exploited = normal patch cycle with monitoring; Medium/Low = track, patch on standard schedule. Always give a confidence level (confirmed / likely / unconfirmed) and cite the specific source, not just "reports suggest".', 'intel,sources,ctu,escalation'],
  ['ref-08-its-policy-catalog', 'shared', '08-its-policy-catalog', 'How LAUSD IT security policy is organized (index, not a single document): Acceptable Use (what staff/students may do with district systems), Data Classification (BUL-999.16 — sensitivity tiers and required controls per tier), Access Control (identity lifecycle, MFA, least privilege — see 06-identity-basics), Incident Response (detection, escalation, communication — see 07-ctu-sources escalation matrix), Vendor/Third-Party Risk (ARB process — see 04-vendor-arb-checklist), Change Management (review/approval before production changes to systems handling sensitive data). When a question spans categories, name which policy area(s) apply rather than guessing at a single bulletin number — cite the exact BUL/RUP number only when you actually have it.', 'policy,index,its'],
  ['ref-09-rup-obligations', 'shared', '09-rup-obligations', 'Responsible Use Policy (RUP) — general obligations, not a substitute for the current signed district RUP: staff and students must use district systems for their intended educational/operational purpose, must not share credentials, must report suspected security incidents promptly, and must not attempt to bypass content filtering or access controls. System owners introducing a new tool must register it for RUP review before rollout if it touches student data or district accounts — this triggers the same data-classification and vendor-review steps as 04-vendor-arb-checklist. Do not tell a user their specific use case is RUP-compliant without the actual current RUP text — say you cannot confirm and point them to IT Governance.', 'rup,policy,obligations'],
  ['ref-10-cybersafety-public', 'shared', '10-cybersafety-public', 'Public-facing cyber safety basics for staff/students: use a unique, long passphrase per account and a password manager rather than reuse; enable MFA everywhere it is offered; be suspicious of urgency, too-good-to-be-true offers, and requests for credentials or payment; verify a sender through a second channel before acting on a financial or credential request; keep devices and apps updated; do not plug in unknown USB devices; report anything suspicious to security@lausd.net rather than trying to investigate it yourself.', 'cybersafety,awareness,public']
];

function seedReferenceKB_() {
  if (!_isAdminUserForSeed()) { _toast_('Admin only.'); return; }
  let inserted = 0, updated = 0;
  REFERENCE_KB_DOCS.forEach(function (row) {
    const result = _upsertKbRow_(row[0], row[1], row[2], row[3], row[4]);
    if (result === 'inserted') inserted++; else updated++;
  });
  _toast_('Reference KB seeded: ' + inserted + ' new, ' + updated + ' updated.');
  SpreadsheetApp.getUi().alert(
    'Reference KB Seeded',
    inserted + ' new doc(s), ' + updated + ' updated. These are the docs your agent personas already reference by name ' +
    '(01-nist-csf, 02-cis-controls, 03-k12-data, 06-identity-basics, 07-ctu-sources, 08-its-policy-catalog, 09-rup-obligations, ' +
    '10-cybersafety-public) plus a fuller 04-vendor-arb-checklist. Re-running this is safe — it updates in place by id.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

// Tone strings condensed from each agent's own existing PRIMARY ROLE / STRICT
// RULES text in 12_seed_all_risk_agents.js — not new claims, just distilled
// style guidance for _buildSystemPrompt_'s "Tone:" line.
var AGENT_TONE_DEFAULTS = {
  'security-helpdesk': 'Warm, patient, plain-language. Empathy first, then clear next steps. Never rushes the user.',
  'idm': 'Precise and procedural. Explains the reasoning behind approval workflows rather than shortcutting them.',
  'grc': 'Precise, conservative, evidence-based. States "I don\'t have enough information" rather than guessing.',
  'netsec': 'Technical and risk-aware. Zero-trust-minded; favors compensating controls over convenience.',
  'ctu': 'Factual and calm under pressure. Leads with severity and next steps, not speculation.',
  'security-architect': 'Strategic, defense-in-depth mindset. Balances rigor with practical delivery.',
  'soc-manager': 'Operational and metrics-driven. Speaks in process, ownership, and escalation terms.',
  'vendor-review-consultant': 'Thorough and checklist-driven. Firm about not waiving incomplete documentation.',
  'ai-systems-analyst': 'Analytical and transparency-focused. Probes for bias, logging, and human-oversight gaps.',
  'cybersop-forge': 'Structured, playbook-style writer. Every SOP has explicit triggers, steps, and escalation paths.',
  'sow-generator-lausd': 'Formal and contract-precise. Never omits deliverables, SLAs, or compliance language.',
  'vulnerability-reporter': 'Clear and business-friendly. Translates technical findings into severity, impact, fix steps.',
  'ultimate-dast-analyzer': 'Business impact first, technical detail second. Solution-oriented, never exploit-oriented.',
  'ip-url-health-analyzer': 'Verdict-first and concise. States confidence level and recommended action up front.',
  'exam-tutor': 'Encouraging and Socratic. Teaches the underlying concept rather than handing out answers.'
};

// Only agents that plausibly need more than bare KB lookup, based on role:
// live threat/CVE/IOC context (webpage/search) or deliverable hand-off (email).
var AGENT_SKILLS_RECOMMENDED = {
  'grc': 'kb,webpage',
  'ctu': 'kb,webpage,search',
  'security-architect': 'kb,webpage',
  'vendor-review-consultant': 'kb,webpage',
  'cybersop-forge': 'kb,email',
  'sow-generator-lausd': 'kb,email',
  'vulnerability-reporter': 'kb,webpage,search',
  'ultimate-dast-analyzer': 'kb,webpage',
  'ip-url-health-analyzer': 'kb,webpage,search'
};

function _agentsSheetColMap_() {
  const sheet = _getSs_().getSheetByName('Agents');
  if (!sheet || sheet.getLastRow() < 1) return null;
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h || '').trim().toLowerCase(); });
  const map = {};
  headers.forEach(function (h, i) { if (h) map[h] = i; });
  return { sheet: sheet, map: map };
}

// Only fills cells that are still blank — safe to re-run, never clobbers a
// tone an admin has since customized by hand.
function backfillAgentTone_() {
  if (!_isAdminUserForSeed()) { _toast_('Admin only.'); return; }
  const ctx = _agentsSheetColMap_();
  if (!ctx || ctx.map.slug === undefined || ctx.map.tone === undefined) {
    _toast_('Run "Migrate Agents (add skills/tone columns)" first.');
    return;
  }
  const numRows = ctx.sheet.getLastRow() - 1;
  if (numRows < 1) { _toast_('No agents found.'); return; }
  const data = ctx.sheet.getRange(2, 1, numRows, ctx.sheet.getLastColumn()).getValues();
  let filled = 0;
  data.forEach(function (row, i) {
    const slug = String(row[ctx.map.slug] || '').trim();
    const currentTone = String(row[ctx.map.tone] || '').trim();
    if (currentTone) return; // already set (manually or previously) — don't touch
    const suggested = AGENT_TONE_DEFAULTS[slug];
    if (!suggested) return;
    ctx.sheet.getRange(2 + i, ctx.map.tone + 1).setValue(suggested);
    filled++;
  });
  _toast_('Backfilled tone for ' + filled + ' agent(s). Existing non-blank tones left untouched.');
}

// Only upgrades rows still at the migration default 'kb' — safe to re-run,
// never clobbers a skills value an admin has since customized by hand.
function upgradeAgentSkills_() {
  if (!_isAdminUserForSeed()) { _toast_('Admin only.'); return; }
  const ctx = _agentsSheetColMap_();
  if (!ctx || ctx.map.slug === undefined || ctx.map.skills === undefined) {
    _toast_('Run "Migrate Agents (add skills/tone columns)" first.');
    return;
  }
  const numRows = ctx.sheet.getLastRow() - 1;
  if (numRows < 1) { _toast_('No agents found.'); return; }
  const data = ctx.sheet.getRange(2, 1, numRows, ctx.sheet.getLastColumn()).getValues();
  let upgraded = 0;
  data.forEach(function (row, i) {
    const slug = String(row[ctx.map.slug] || '').trim();
    const currentSkills = String(row[ctx.map.skills] || '').trim().toLowerCase();
    const recommended = AGENT_SKILLS_RECOMMENDED[slug];
    if (!recommended) return;
    if (currentSkills !== 'kb' && currentSkills !== '') return; // already customized — don't touch
    ctx.sheet.getRange(2 + i, ctx.map.skills + 1).setValue(recommended);
    upgraded++;
  });
  _toast_('Upgraded skills for ' + upgraded + ' agent(s) still at default "kb". Review before shipping "search" (needs Serper key).');
}

// ==========================================================================
// SKILLS CATALOG (v1.3.0) — reference tab: what's implemented vs. possible.
//
// The 'implemented' skills are ported from Khoj's own ConversationCommand
// set (src/khoj/utils/helpers.py in the local khoj-selfhosted clone at
// C:\Users\remon.attalla\Documents\GIT\risk) plus two GAS-native additions
// (email, and this catalog itself) that don't exist in Khoj. The 'proposed'
// rows are the rest of Khoj's command set, evaluated for portability to
// Apps Script -- code/research/image/operator all need infrastructure GAS
// doesn't have (a code sandbox, an image model, browser automation), so
// they're documented but NOT built. Don't wire one up without a SECURITY.md
// pass first -- 'code' in particular is exactly the kind of thing that
// bulletin exists to gate.
// ==========================================================================

var SKILLS_CATALOG = [
  // key, status, description, source, how_to_enable
  ['kb', 'implemented',
    'Injects matching rows from the KB tab (slug=shared + slug=<agent>) into the system prompt. Default skill if Agents.skills is blank.',
    'native', "Add 'kb' to the agent's skills column (or leave blank)."],
  ['memory', 'implemented',
    'Hourly sweep folds recent ChatLog exchanges into a short running note per (user, agent), re-injected into future sessions. No vector DB -- a sweep, not per-message.',
    'Khoj (UserMemory, ported)', "Add 'memory' to skills, then run AGENTS -> Memory -> Install Hourly Memory Trigger once."],
  ['webpage', 'implemented',
    'If the user pastes a URL, fetches it, strips HTML, injects up to 4000 chars as context.',
    'Khoj (Webpage / ReadWebpage)', "Add 'webpage' to skills."],
  ['search', 'implemented',
    'User prefixes a message with "search: <query>"; calls Serper.dev and injects the top results. Fails soft (tells the model search is unavailable) if no key is set.',
    'Khoj (Online / SearchWeb)', "Add 'search' to skills, then AGENTS -> Admin + Ingest -> Set Web Search API Key (Serper)."],
  ['email', 'implemented',
    'Emails the current chat transcript to the REQUESTING user\'s own address only -- no arbitrary recipient, to prevent the web app relaying mail to third parties.',
    'GAS-native (not in Khoj)', "Add 'email' to skills."],
  ['diagram', 'implemented (v1.3.0)',
    'Instructs the model to output a Mermaid diagram in a ```mermaid fenced block when a flowchart/relationship is best shown visually. Rendered as plain text/code in this chat UI, not drawn inline -- paste into a Mermaid viewer.',
    'Khoj (Diagram)', "Add 'diagram' to skills."],
  ['code', 'proposed - not built',
    'Run a script to do calculations, parse data, or generate a chart. Khoj runs this in an ephemeral E2B or Terrarium Python sandbox (your khoj-selfhosted docker-compose already runs the Terrarium sandbox container). Apps Script has no equivalent safe sandbox -- would need to call out to an external code-exec API.',
    'Khoj (Code / run_code)', 'NOT implemented. Needs a SECURITY.md review before building -- arbitrary code execution is the highest-risk skill on this list.'],
  ['research', 'proposed - not built',
    'Multi-step/iterative search+read loop for a deeper answer than one search pass. Could be approximated in GAS as a loop calling the existing webpage/search skill functions 2-3x before answering.',
    'Khoj (Research)', 'NOT implemented.'],
  ['image', 'not planned',
    'Generate an illustrative image from a text description.',
    'Khoj (Image)', 'NOT implemented -- no image-gen API wired up; out of scope for a text chat UI.'],
  ['operator', 'not planned',
    '"Operate a computer" -- browser/UI automation to complete a task.',
    'Khoj (Operator)', 'NOT implemented -- outside Apps Script\'s execution model and this platform\'s security posture.'],
  ['calendar', 'proposed - not built',
    "Read-only lookup of the REQUESTING user's own upcoming Calendar events, mirroring the 'email' skill's self-only safety pattern.",
    'GAS-native idea (not in Khoj)', 'NOT implemented. Needs a new calendar.readonly OAuth scope in appsscript.json -- re-authorization required.'],
  ['gmail-search', 'proposed - not built',
    "Read-only search of the REQUESTING user's own Gmail for context (e.g. \"did I already get an advisory about vendor X\").",
    'GAS-native idea (not in Khoj)', 'NOT implemented. Needs a new gmail.readonly OAuth scope -- re-authorization required.']
];

function _upsertSkillsCatalogRow_(key, status, description, source, howToEnable) {
  const sheet = _getOrCreateTab_('SkillsCatalog', ['skill_key', 'status', 'description', 'source', 'how_to_enable']);
  const data = sheet.getDataRange().getValues();
  const row = [key, status, description, source, howToEnable];
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]) === key) {
      sheet.getRange(i + 1, 1, 1, row.length).setValues([row]);
      return 'updated';
    }
  }
  sheet.appendRow(row);
  return 'inserted';
}

function seedSkillsCatalog_() {
  if (!_isAdminUserForSeed()) { _toast_('Admin only.'); return; }
  let inserted = 0, updated = 0;
  SKILLS_CATALOG.forEach(function (row) {
    const result = _upsertSkillsCatalogRow_(row[0], row[1], row[2], row[3], row[4]);
    if (result === 'inserted') inserted++; else updated++;
  });
  _toast_('SkillsCatalog seeded: ' + inserted + ' new, ' + updated + ' updated. See the SkillsCatalog tab.');
}
