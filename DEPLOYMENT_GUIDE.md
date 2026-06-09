# Risk App - Deployment & Per-Agent URLs Guide (Google Apps Script)

## 1. Per-Agent Shareable URLs
The Web App provides a **unique URL for every agent** using a query parameter.

**Base Web App URL** (after deployment):
`https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec`

**Per-agent URL**:
`https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec?agent=your-agent-slug`

- Replace `YOUR_DEPLOYMENT_ID` with the ID from your Web App deployment.
- `your-agent-slug` comes from the `slug` column in your Agents sheet (e.g., `incident-response`).
- Anyone with the link can chat with that specific agent (subject to privacy_level).
- The UI is minimal/locked: no sidebar, no agent picker, no way to switch agents. Perfect for shareable links.

Example:
- Agent slug "phishing-advisor" → URL: `.../exec?agent=phishing-advisor`
- Visiting it loads the branded Risk AI chat directly for that agent.

## 2. Google Auth & Admin Powers
Google authentication is built-in using `Session.getActiveUser().getEmail()`.

- When someone visits the Web App (while logged into a Google account), the code knows who they are.
- **You are the admin** if your email matches `ADMIN_EMAILS` (currently only `rayattalla@gmail.com`).
- Only admins see the creation form and full agent details.
- Non-admins see redacted agents (no persona/creator) and can only chat with agents they have access to (based on `privacy_level` + creator).

**How auth works in practice**:
- Deploy the Web App with "Who has access: Anyone" (or "Anyone with the link").
- "Execute as: Me" (so it runs with your permissions for Sheets access).
- When **you** visit the base URL while logged into `rayattalla@gmail.com`, you get the admin dashboard.
- When others visit a `?agent=slug` link, they get the chat (Google may prompt them to sign in for the session).

To add more admins later: Edit the `ADMIN_EMAILS` array in Code.gs.

## 3. How to Deploy New Agents
You create agents **as the admin** through the web UI (no need to edit the sheet manually).

### Steps to Create & Deploy a New Agent:
1. **Deploy/Update the Web App** (one-time or when you change code):
   - In Apps Script editor → Deploy → New deployment (or Manage deployments).
   - Type: Web app
   - Execute as: Me
   - Who has access: Anyone (recommended for shareable links) or Anyone with the link.
   - Copy the Web App URL (this is your base + per-agent URLs).

2. **Visit as Admin**:
   - Go to the **base Web App URL** (no `?agent=` param) while logged into your Google account (`rayattalla@gmail.com`).
   - You will see the **Admin Dashboard** (because of the `isAdmin` check).

3. **Create the Agent**:
   - Fill the form:
     - **Slug**: Unique, URL-safe (e.g., `incident-response`, `phishing-advisor`). This becomes part of the URL.
     - **Name**: Human-friendly display name.
     - **Personality**: Your raw agent instructions (the guardrail is automatically prepended).
     - **Privacy Level**: PUBLIC (shareable by anyone with link), PROTECTED, or PRIVATE.
   - Click "Create Agent & Generate URL".
   - The system:
     - Adds the row to the Agents sheet.
     - Prepends the full RISK-AI SAFETY GUARDRAIL.
     - Returns the exact shareable URL.

4. **Test the New URL**:
   - Use the generated link: `https://.../exec?agent=your-new-slug`
   - It loads the locked Risk AI chat for that agent only.
   - Chat history is logged to the Messages sheet.
   - Attempts to jailbreak the prompt are blocked (client + server guardrails).

### Alternative (Advanced) - Direct Sheet + Harden
- Manually add a row to the Agents sheet.
- Run `hardenAgent("your-slug")` from the Apps Script editor (or a menu).
- The URL becomes available immediately.

## 4. Google Auth Flow Summary for Creation
- Only your Google account (the one in ADMIN_EMAILS) can access the creation form.
- When you visit the base URL:
  - GAS sees your email via `Session.getActiveUser()`.
  - `isAdmin()` returns true → you see the create form + list of all agents (with full details).
- Non-admins visiting the base URL see a message: "No agent specified. Ask the admin for a link."
- Shared agent URLs (`?agent=slug`) work for everyone (they don't need to be logged in, though Google may ask for consent on first use for the session).

## 5. Tips for Production Use
- **Redeploy** after any Code.gs changes (new deployment version recommended).
- Keep the Agents sheet private (share only with yourself or specific people).
- Use `privacy_level=PUBLIC` for truly shareable agents.
- Monitor the Logs sheet for usage / blocked attempts.
- For more advanced auth (OAuth scopes, etc.), you can expand the `isAdmin` logic or add a simple login gate.
- The chat uses Google Gemini API (set `GEMINI_API_KEY` in Project Settings > Script properties).

Your per-agent URLs are live as soon as you create the agent via the admin UI. You control creation exclusively via your Google login.

If you need a custom domain, more UI polish, or agent management features, let me know!
