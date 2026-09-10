# P2 Sheet MVP — Agent Platform (locked, shipping prep)

Tabs (current unified design):

- Agents (slug | name | status | org | model | personality | kb_tags | notes)
- KB (id | slug | title | body | tags)   // slug='shared' (global) or matches Agents.slug
- ChatLog (Timestamp | Email | Slug | Prompt | Response)
- Config
- Access_Control
- Audit_Log (etc from setup)

Old KB_* and KB_shared tabs should be deleted via cleanup script after migration.

## Code wiring complete in gas-agent-platform

- lib/01_config.js : SHEET_TABS (AGENTS, KB, ...)
- lib/02_helpers.js : getAgent_, getKbSnippets_ (reads unified KB for shared + slug), chatWithAgent_, getActiveAgents_, setup() creates correct headers
- template/shell.js : doGet(?agent=slug or ?page=admin), wrappers, menu with cleanup + ship checklist
- template/index.html : clean chat UI + Admin link for admins
- template/10_admin_ingest.js : admin CRUD, ingest (writes to KB with slug), migrate, cleanupOldKbTabs_
- template/12_seed_all_risk_agents.js : seeds 15 agents with kb_tags, targets single KB
- Menu: AGENTS → Admin + Ingest (seed, migrate, cleanup for shipping, admin page)

## Full roster (status=on)

security-helpdesk, idm, grc, netsec, ctu, security-architect, soc-manager, vendor-review-consultant, ai-systems-analyst, cybersop-forge, sow-generator-lausd, vulnerability-reporter, ultimate-dast-analyzer, ip-url-health-analyzer, exam-tutor

## Shipping steps (run these)

1. In Sheet: AGENTS → Setup & Config → Run Setup (ensures current tabs/headers).
2. AGENTS → Admin + Ingest → Seed all 15 Risk agents (populates Agents + KB).
3. (if old data) Migrate KB tabs to single KB.
4. AGENTS → Admin + Ingest → Cleanup old KB tabs (shipping) — removes legacy KB_* / KB_shared.
5. AGENTS → Setup & Config → Set Web App URL (paste current /exec URL).
6. Script Properties: AI_KEY_OPENROUTER (or equivalent) — never in Sheet.
7. clasp push (from lib/ and template/).
8. In GAS editor: Deploy → Manage deployments → edit the web app deployment (or new) with Execute as USER_ACCESSING, access DOMAIN.
9. Run menuShipChecklist and follow (pin lib version, devMode=false in appsscript.json, SECURITY.md, register Notion, disable test mode).
10. Smoke test:
    - /exec → agent list
    - /exec?agent=grc (etc)
    - /exec?page=admin → agent list editor (admin only)
    - Verify only one KB tab, Agents uses kb_tags column.

No keys in git/Sheet. Sheet is sole control plane.

Use /exec?page=admin (or sidebar) to manage agents/status.

Project cleaned and prepared for shipping.
