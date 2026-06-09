# Google Sheets Schema for Risk (GAS Port)

Use the provided spreadsheet:
https://docs.google.com/spreadsheets/d/1dgKGcCDgpwkEvZyFOsfpznOmLo3HvPp_T2oKQKgPhLY/edit

Create the following sheets (tabs). Use the exact column headers for easy Apps Script access via getRange + getValues.

## 1. Agents (main table - one row per agent)

Columns (A onwards):
A: slug (text, primary key, e.g. "incident-response", "phishing-triage". Must be URL-safe, lowercase with hyphens)
B: name (text, human friendly)
C: personality (text, LONG - this should contain the FULL guarded personality. Prepend the GUARDRAIL.txt content when creating/editing agents)
D: privacy_level (text: PUBLIC | PROTECTED | PRIVATE)
E: creator (text: email of the person who owns/created the agent)
F: style_color (text: hex like "#dc2626" or name)
G: style_icon (text: e.g. "shield" or emoji or class name)
H: input_tools (text: JSON array or comma-separated, e.g. ["web_search","code_execution"] )
I: output_modes (text: JSON or comma, e.g. ["text","summary"])
J: managed_by_admin (boolean: TRUE / FALSE)
K: is_hidden (boolean)
L: created_at (date or timestamp)
M: notes (text, optional)

Example row for a public security agent:
slug | name | personality | privacy_level | creator | ...
incident-response | Incident Response Advisor | # RISK-AI SAFETY GUARDRAIL ... [full guardrail + actual persona text] | PUBLIC | remon.attalla@.net | #dc2626 | shield | ["web_search"] | ["text"] | FALSE | FALSE | 2026-06-...

**Important for security**: When you create agents, always make sure column C starts with the content of GUARDRAIL.txt (or append it in your GAS code before sending to LLM).

## 2. Conversations (chat sessions - one row per conversation)

A: conversation_id (text, generate with Utilities.getUuid() or similar)
B: agent_slug (text, foreign key to Agents!A)
C: user_email (text)
D: created_at (timestamp)
E: updated_at (timestamp)
F: title (text, optional - can be first user message or generated)

## 3. Messages (full chat history + context for LLM + audit trail)

A: message_id (text, uuid or row number)
B: conversation_id (text)
C: role (text: "user" or "assistant")
D: content (text, the actual message)
E: timestamp (timestamp)
F: metadata (text, JSON string: {"blocked_injection": false, "model": "gemini-1.5-flash", "tokens": 123, "latency_ms": 2340, ...})

This sheet will grow. Use it both for:
- Building chat history when calling the LLM (last N messages for the conversation_id)
- Security/audit logs

## 4. Logs (security, usage, and operational audit - append-only)

A: timestamp
B: user_email (or "anonymous")
C: agent_slug
D: action (e.g. "get_agent", "create_session", "chat", "injection_blocked", "error")
E: details (text or JSON: the query, http status, error message, etc.)
F: ip_or_client (optional)

## 5. Admins (simple allowlist - or use Script Properties instead)

A: email (the two from original: remon.attalla@.net and rayattalla@gmail.com)
B: role (e.g. "admin")

For the GAS code, you can hardcode the admin emails in Script Properties for speed, or read this sheet.

## Setup Instructions (Direct in the Sheet)

1. Open the sheet link.
2. Rename the default first tab to "Agents" and add the columns above (row 1 = headers).
3. Create new tabs: "Conversations", "Messages", "Logs", "Admins".
4. For any agent you add, copy the entire GUARDRAIL.txt content + a newline + "---" + newline + the actual useful persona text into the "personality" cell.
5. Share the spreadsheet with the account that owns the Apps Script (or use the same account).

## How the GAS Code Will Use These Sheets

- SpreadsheetApp.openById('1dgKGcCDgpwkEvZyFOsfpznOmLo3HvPp_T2oKQKgPhLY')
- Get sheets by name: ss.getSheetByName('Agents')
- Use getDataRange().getValues() + filter, or getRange for specific rows.
- For writes: appendRow() on Logs/Messages/Conversations.
- For performance: keep Agents small and load into memory (CacheService) at startup if needed. Messages can get large — query with filters.

## Redaction Rule (implement in GAS)

When returning agent data for a user:
- If user email is in Admins OR user email == creator:
  - Return full persona, creator, etc.
- Else:
  - Return persona: null (or empty)
  - creator: null
  - chat_model / files if present: null or []

This matches the original api_agents.py redaction.

## Privacy Levels (from original)

- PUBLIC: Anyone with the link (or who can discover the slug) can use it (subject to your web app sharing settings).
- PROTECTED: Similar but perhaps requires being in the org.
- PRIVATE: Only the creator (and admins).

In the readonly agent lookup, the original allowed PUBLIC | PROTECTED | creator.

You can implement the same filter in GAS when listing or fetching an agent by slug.

This schema is the minimal faithful port of the data concepts from the original Risk hardening work.
