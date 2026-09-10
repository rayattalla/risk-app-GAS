/**
 * RiskAI-GAS — Email (04_email.js)
 *
 * LAUSD-branded HTML email layout + send wrapper.
 * All outgoing mail respects TEST_MODE — set via enableTestMode() in 01_config.js.
 *
 * Public API:
 *   sendMail(to, subject, htmlBody, opts)  — main send function
 *   emailLayout(title, bodyHtml)           — returns full HTML email string
 *   emailButton(label, url, color)         — CTA button HTML
 *   emailKV(label, value)                  — key/value row HTML (use in a <table>)
 *   emailSection(heading, contentHtml)     — labelled section HTML
 *
 * Usage pattern:
 *   const body = '<table>' + emailKV('Status', 'Approved') + '</table>'
 *              + emailButton('Open Dashboard', webAppUrl);
 *   sendMail(submitterEmail, 'Record approved', emailLayout('Approved', body));
 */

// ==========================================================================
// BRANDING CONSTANTS
// ==========================================================================

const LAUSD_GRADIENT = 'linear-gradient(90deg,#F47B20 0%,#ED1C24 25%,#0072CE 50%,#00A3E0 75%,#5B6CB0 100%)';

const LAUSD_COLORS = {
  BLUE:       '#0072CE',
  DARK_BLUE:  '#1A237E',
  ORANGE:     '#F47B20',
  RED:        '#ED1C24',
  TEAL:       '#00A3E0',
  PURPLE:     '#5B6CB0',
  BACKGROUND: '#F5F7FA',
  CARD:       '#FFFFFF',
  TEXT:       '#1A1A2E',
  MUTED:      '#5F6368',
  BORDER:     '#E8EAED',
  CRITICAL:   '#B71C1C',
  HIGH:       '#E65100',
  MEDIUM:     '#F9A825',
  LOW:        '#2E7D32'
};

// ==========================================================================
// SEND WRAPPER
// ==========================================================================

/**
 * Send an HTML email. Respects TEST_MODE.
 * opts: { cc, bcc, replyTo }
 */
function sendMail(to, subject, htmlBody, opts) {
  if (!to) return;
  let recipient = to;
  if (isTestMode()) {
    recipient = getTestEmail();
    subject   = '[TEST → ' + to + '] ' + subject;
  }
  opts = opts || {};
  const cc = opts.cc || CONFIG.NOTIFY_EMAIL;
  try {
    MailApp.sendEmail({
      to:       recipient,
      cc:       isTestMode() ? '' : cc,
      bcc:      isTestMode() ? '' : (opts.bcc || ''),
      subject:  subject,
      htmlBody: htmlBody,
      name:     CONFIG.EMAIL_SENDER_NAME,
      replyTo:  opts.replyTo || CONFIG.NOTIFY_EMAIL
    });
  } catch (e) {
    Logger.log('sendMail failed to ' + recipient + ': ' + e.message);
  }
}

// ==========================================================================
// LAYOUT — full email wrapper with LAUSD gradient header
// ==========================================================================

/**
 * Wrap bodyHtml in the standard LAUSD email layout.
 * title appears as the H1 in the card header.
 */
function emailLayout(title, bodyHtml) {
  return ''
    + '<!doctype html><html><head><meta charset="utf-8">'
    + '<meta name="viewport" content="width=device-width,initial-scale=1"></head>'
    + '<body style="margin:0;padding:0;background:' + LAUSD_COLORS.BACKGROUND + ';'
    +   'font-family:-apple-system,Segoe UI,Roboto,Arial,sans-serif;color:' + LAUSD_COLORS.TEXT + ';">'
    + '<table role="presentation" cellpadding="0" cellspacing="0" width="100%" '
    +   'style="background:' + LAUSD_COLORS.BACKGROUND + ';padding:24px 12px;"><tr><td align="center">'
    + '<table role="presentation" cellpadding="0" cellspacing="0" width="620" '
    +   'style="background:' + LAUSD_COLORS.CARD + ';border-radius:12px;'
    +   'box-shadow:0 4px 14px rgba(0,0,0,.06);overflow:hidden;">'
    // Gradient accent bar
    +   '<tr><td style="height:6px;background:' + LAUSD_GRADIENT + ';"></td></tr>'
    // Header
    +   '<tr><td style="padding:24px 32px 8px;">'
    +     '<div style="font-size:.78rem;text-transform:uppercase;letter-spacing:.08em;color:' + LAUSD_COLORS.MUTED + ';">'
    +       _esc_(CONFIG.ORG_SHORT) + ' &middot; ' + _esc_(CONFIG.PROGRAM_NAME)
    +     '</div>'
    +     '<h1 style="font-size:1.4rem;margin:.3rem 0 0;color:' + LAUSD_COLORS.DARK_BLUE + ';">'
    +       _esc_(title)
    +     '</h1>'
    +   '</td></tr>'
    // Body
    +   '<tr><td style="padding:8px 32px 32px;font-size:.95rem;line-height:1.6;">'
    +     bodyHtml
    +   '</td></tr>'
    // Footer
    +   '<tr><td style="padding:12px 32px;background:#F0F4F8;font-size:.78rem;color:' + LAUSD_COLORS.MUTED + ';">'
    +     'Automated message from the LAUSD ' + _esc_(CONFIG.PROGRAM_NAME) + ' system. '
    +     'Reply to <a href="mailto:' + _esc_(CONFIG.NOTIFY_EMAIL) + '" style="color:' + LAUSD_COLORS.BLUE + ';">'
    +       _esc_(CONFIG.NOTIFY_EMAIL)
    +     '</a>.'
    +   '</td></tr>'
    + '</table></td></tr></table></body></html>';
}

// ==========================================================================
// BUILDING BLOCKS — compose email bodies with these
// ==========================================================================

/**
 * A CTA button.
 * color defaults to LAUSD blue.
 */
function emailButton(label, url, color) {
  color = color || LAUSD_COLORS.BLUE;
  return '<p style="margin:1.2rem 0;">'
    + '<a href="' + _esc_(url) + '" '
    + 'style="display:inline-block;background:' + color + ';color:#fff;'
    + 'padding:.55rem 1.2rem;border-radius:6px;text-decoration:none;font-weight:600;">'
    + _esc_(label)
    + '</a></p>';
}

/**
 * A key/value row for use inside a <table> element.
 * Wrap multiple emailKV() calls in <table style="border-collapse:collapse;margin:.5rem 0 1rem;">...</table>
 */
function emailKV(label, value) {
  return '<tr>'
    + '<td style="padding:4px 16px 4px 0;color:' + LAUSD_COLORS.MUTED + ';font-size:.85rem;width:140px;vertical-align:top;">'
    +   _esc_(label)
    + '</td>'
    + '<td style="padding:4px 0;font-weight:500;">'
    +   _esc_(value || '—')
    + '</td>'
    + '</tr>';
}

/**
 * A labelled section divider (horizontal rule + heading).
 */
function emailSection(heading, contentHtml) {
  return '<hr style="border:none;border-top:1px solid ' + LAUSD_COLORS.BORDER + ';margin:1.2rem 0;">'
    + '<h2 style="font-size:1rem;color:' + LAUSD_COLORS.DARK_BLUE + ';margin:0 0 .5rem;">'
    +   _esc_(heading)
    + '</h2>'
    + contentHtml;
}

/**
 * Inline badge/chip with background color. Useful for status/risk tier labels.
 */
function emailBadge(label, color, textColor) {
  color     = color     || LAUSD_COLORS.BLUE;
  textColor = textColor || '#fff';
  return '<span style="display:inline-block;background:' + color + ';color:' + textColor + ';'
    + 'padding:.15rem .5rem;border-radius:4px;font-size:.8rem;font-weight:600;">'
    + _esc_(label)
    + '</span>';
}
