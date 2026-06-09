# risk-app-GAS

Personal port of the Risk AI hardened agent chat system to Google Apps Script + Google Sheets.

**This is a personal project (no organizational branding).**

## Features
- Shareable per-agent URLs: `?agent=your-slug`
- Strong security: 
  - Server-side redaction of persona/creator for non-owners/admins
  - Hardened non-disclosure guardrail prepended to every agent personality
  - Client-side + server-side prompt injection/jailbreak filtering
- Minimal locked-down UI for shared agents (no full Khoj sidebar, no agent switching)
- Google Sheets as DB:
  - Agents (with personality, privacy_level, etc.)
  - Conversations/sessions
  - Messages (history for context)
  - Logs (audit/usage)
- **Google API only (Gemini via Google AI API / UrlFetchApp)** — no local LLM (no Ollama, etc.)

## Setup
1. Open the Google Sheet: https://docs.google.com/spreadsheets/d/1dgKGcCDgpwkEvZyFOsfpznOmLo3HvPp_T2oKQKgPhLY
2. Create tabs per SHEETS_SCHEMA.md (Agents, Conversations, Messages, Logs)
3. In Apps Script (script ID from your project), set Script Properties:
   - `GEMINI_API_KEY` = your Google AI (Gemini) API key
4. Deploy as Web App (Execute as: Me, Who has access: Anyone with link)
5. Use the Web App URL + `?agent=slug` for shareable links
6. (Optional) Run `hardenAgent("slug")` in editor to apply guardrail to agents

## Files
- `Code.gs` - Main backend (redaction, sessions, chat with Gemini, harden, doGet)
- `RiskChat.html` - Minimal standalone chat UI (adapts the original risk-chat.html)
- `appsscript.json` - GAS project config for Web App
- `GUARDRAIL.txt` - The RISK-AI safety guardrail text (prepend to personalities)
- `INJECTION_PATTERNS.js` - Client-side jailbreak filter (from original)
- `SHEETS_SCHEMA.md` - Exact columns for the Google Sheet DB

## Development Notes
- All LLM calls use Google Gemini API (no local models).
- Agent access control based on privacy_level + creator/admin checks.
- Logging for security/audit.

Built from the risk-app hardening work, ported to pure Google stack.

Personal project by Ray Attalla.
