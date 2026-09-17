# Branding assets (LAUSD / Risk AI)

Apps Script HtmlService cannot serve files from this folder. The live web app
embeds the same marks as inline SVG in `template/index.html`. These files are
the repo-side source of truth so Ray can drop official replacements later.

## Expected paths

| Mark | Path | Status |
| --- | --- | --- |
| LAUSD seal | `template/assets/lausd-seal.svg` (PNG also acceptable: `lausd-seal.png`) | **Stand-in only.** No official LAUSD seal was in the repo. The committed SVG is a typographic circular "LAUSD" badge in house navy `#1F3864` with amber `#C55A11` ring — not a replica of the official seal (torch / books / rim lettering). |
| Risk AI logo | `template/assets/risk-ai-logo.svg` | **In-repo mark reused.** Same sparkline already embedded in `template/index.html` (navy square, amber line, teal dot). |

## How to replace with official files

1. Drop the official LAUSD seal at `template/assets/lausd-seal.svg` (or `.png`).
2. Drop the official Risk AI wordmark/logo at `template/assets/risk-ai-logo.svg` (or `.png`) if it differs from the sparkline.
3. Re-embed the new artwork in `template/index.html` (inline SVG or data URI). GAS will not pick up this folder on `clasp push`.
4. Do not invent a fake official seal. Prefer the real district artwork.

House colors used by the stand-in: navy `#1F3864`, amber `#C55A11`. RISK wordmark in the UI is `#1F3864`; AI stays red `#dc2626`.
