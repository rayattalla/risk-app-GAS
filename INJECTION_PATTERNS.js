// Exact client-side prompt injection / jailbreak filter from risk-chat.html
// This runs in the browser before any call to the backend.
// It is a polite UX layer; the guardrail in the personality is the real defense.
const INJECTION_PATTERNS = [
  /\bignore\s+(all\s+)?(your\s+)?(previous|prior|above|the\s+above)\s+(instructions?|rules?|prompts?|messages?|directives?)\b/i,
  /\b(forget|disregard)\s+(your|the|all|previous|prior)\s+(instructions?|rules?|prompts?|training)/i,
  /\b(show|reveal|tell|give|share|expose|output|print|reproduce|repeat)\s+(me\s+)?(your\s+|the\s+|all\s+|full\s+)?(system\s+)?(prompt|instructions?|rules?|persona|guidelines?|training|configuration)\b/i,
  /\bwhat\s+(are\s+|is\s+)?(your|the)\s+(system\s+)?(instructions?|rules?|prompt|training|configuration|guidelines?|persona)\b/i,
  /\brepeat\s+(the\s+)?(text|words|content|message)\s+above\b/i,
  /\byou\s+are\s+now\s+(a|an|in)?/i,
  /\bact\s+as\s+(a|an|if|though)?/i,
  /\bpretend\s+(you\s+are|to\s+be)\b/i,
  /\bdeveloper\s+mode\b/i,
  /\bjailbreak/i,
  /\bDAN\s+mode\b/i,
  /\bswitch\s+(your\s+)?persona\b/i,
  /\bbreak\s+character\b/i,
  /\bsystem\s+message\b/i,
  /\binitial\s+prompt\b/i,
];

function looksLikeInjection(text) {
  for (const re of INJECTION_PATTERNS) {
    if (re.test(text)) return true;
  }
  return false;
}
