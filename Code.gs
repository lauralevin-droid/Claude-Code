/**
 * Thrive Market Email Copy Template Generator
 *
 * Reads a Wrike brief, auto-selects an email copy template, copies it to
 * Google Drive, pre-fills known fields, and returns an edit link.
 *
 * Setup (run once):
 *   1. Open Apps Script editor → select setupWrikeToken → Run
 *   2. Enter your Wrike permanent access token when prompted
 *      (Wrike profile → Apps & Integrations → API → Create token)
 *
 * Deploy as Web App:
 *   Deploy → New deployment → Web App
 *   Execute as: Me  |  Who has access: Anyone within Thrive Market
 */

// ─── Template IDs ─────────────────────────────────────────────────────────────
// Google Doc IDs extracted from the template URLs

var TEMPLATES = {
  content:  { label: 'Content Send (non-promo)', id: '1f6uzg9TZCKMwJdSAUC7pmKyBQOYpWzUwUxAsnGGBMDg' },
  slice:    { label: 'Slice',                    id: '1z6j5We_agYn7ew8N1bbjUYSkZpjTeg2JCQQf2kGij5M' },
  seamless: { label: 'Seamless',                 id: '1DpQ1abtKielLnR1sF8eR6qkgpOCvoJwwsDGZe2vprd4' },
  gwp:      { label: 'GWP / BAGO',              id: '1RegQ-F-5AgcA-rlYP2UvS2wyUlNinysvEd-8d_ptndw' },
  pctoff:   { label: '% Off',                    id: '1RuW0XDoT48Hr3Sv08l9IBCtbBf6PZCLC8Gh73I2KxDY' },
};

// ─── One-time Wrike token setup ───────────────────────────────────────────────

// Called from the web app settings panel
function saveWrikeToken(token) {
  if (!token || !token.trim()) return { error: 'Token cannot be empty.' };
  PropertiesService.getScriptProperties().setProperty('WRIKE_TOKEN', token.trim());
  return { ok: true };
}

function hasWrikeToken() {
  return !!PropertiesService.getScriptProperties().getProperty('WRIKE_TOKEN');
}

function getWrikeToken() {
  return PropertiesService.getScriptProperties().getProperty('WRIKE_TOKEN') || '';
}

// ─── Web App entry point ──────────────────────────────────────────────────────

function doGet() {
  var html = getFormHtml(hasWrikeToken());
  return HtmlService.createHtmlOutput(html)
    .setTitle('Thrive Market Email Copy Template Generator')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Called from the HTML form via google.script.run
function processForm(form) {
  var wrikeUrl      = (form.wrikeUrl      || '').trim();
  var templateOverride = (form.templateType || 'auto').trim();

  if (!wrikeUrl) return { error: 'Please paste a Wrike brief URL.' };
  if (!wrikeUrl.includes('wrike.com')) return { error: 'Please enter a valid Wrike URL.' };

  var brief = fetchWrikeBrief(wrikeUrl);
  if (brief.error) return { error: brief.error };

  var templateKey = templateOverride !== 'auto'
    ? templateOverride
    : detectTemplateType(brief.title, brief.description);

  var template = TEMPLATES[templateKey] || TEMPLATES.content;
  var docName  = brief.title || 'Email Copy – ' + wrikeUrl;

  try {
    var newFile = DriveApp.getFileById(template.id).makeCopy(docName);
    var doc     = DocumentApp.openById(newFile.getId());
    fillTemplate(doc, brief);
    doc.saveAndClose();
    return {
      url:          newFile.getUrl(),
      name:         docName,
      templateUsed: template.label,
      briefData:    brief,
    };
  } catch (e) {
    return { error: 'Could not copy template: ' + e.message };
  }
}

// ─── Wrike integration ────────────────────────────────────────────────────────

function fetchWrikeBrief(wrikeUrl) {
  var token = getWrikeToken();
  if (!token) {
    return { error: 'No Wrike API token found. Run setupWrikeToken() from the Apps Script editor first.' };
  }

  // Extract numeric ID from open.htm?id= or from path-based URLs
  var taskId = parseWrikeTaskId(wrikeUrl);
  if (!taskId) {
    return { error: 'Could not parse a task ID from the URL: ' + wrikeUrl };
  }

  // Wrike API: get task by permalink
  var apiUrl = 'https://www.wrike.com/api/v4/tasks?permalink=' + encodeURIComponent(wrikeUrl);
  var response;
  try {
    response = UrlFetchApp.fetch(apiUrl, {
      muteHttpExceptions: true,
      headers: { 'Authorization': 'Bearer ' + token },
    });
  } catch (e) {
    return { error: 'Wrike API request failed: ' + e.message };
  }

  if (response.getResponseCode() === 401) {
    return { error: 'Wrike token is invalid or expired. Run setupWrikeToken() again.' };
  }
  if (response.getResponseCode() !== 200) {
    return { error: 'Wrike API returned HTTP ' + response.getResponseCode() };
  }

  var data;
  try {
    data = JSON.parse(response.getContentText());
  } catch (e) {
    return { error: 'Could not parse Wrike API response.' };
  }

  var tasks = data.data;
  if (!tasks || tasks.length === 0) {
    return { error: 'No task found for that Wrike URL. Make sure the link is a direct task link and your token has access.' };
  }

  var task = tasks[0];
  return extractBriefFromTask(task, wrikeUrl);
}

function parseWrikeTaskId(url) {
  // https://www.wrike.com/open.htm?id=4486916737
  var m = url.match(/[?&]id=(\d+)/);
  if (m) return m[1];
  // https://www.wrike.com/workspace.htm#path=folder&id=4486916737
  m = url.match(/[#&]id=(\d+)/);
  if (m) return m[1];
  return null;
}

function extractBriefFromTask(task, sourceUrl) {
  var title       = task.title || '';
  var description = stripHtml(task.description || '');
  var dates       = parseDatesFromTitle(title) || parseDatesFromDescription(description);
  var promoFocus  = parsePromoFocus(title, description);
  var brand       = parseBrand(title, description);

  return {
    title:       title,
    description: description,
    dates:       dates,
    promoFocus:  promoFocus,
    brand:       brand,
    wrikeUrl:    sourceUrl,
    taskId:      task.id,
    error:       null,
  };
}

// ─── Brief parsing helpers ────────────────────────────────────────────────────

// Parses "08.07" or "08/07" from a Wrike title like "08.07 | Bobo's GWP"
function parseDatesFromTitle(title) {
  var m = title.match(/\b(\d{1,2})[./](\d{1,2})\b/);
  if (!m) return '';
  var month = parseInt(m[1], 10);
  var day   = parseInt(m[2], 10);
  var months = ['January','February','March','April','May','June',
                'July','August','September','October','November','December'];
  return months[month - 1] + ' ' + day;
}

function parseDatesFromDescription(desc) {
  // Look for patterns like "August 7", "8/7", "08-07", etc.
  var m = desc.match(/\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2}/i);
  return m ? m[0] : '';
}

// "08.07 | Bobo's GWP" → "Bobo's GWP"
function parsePromoFocus(title, description) {
  var m = title.match(/\|\s*(.+)$/);
  if (m) return m[1].trim();
  // Look for discount patterns in description
  m = description.match(/(\d+%\s*off|GWP|BOGO|BAGO|gift with purchase|free shipping)/i);
  return m ? m[0] : '';
}

// Try to extract a brand name from the promo focus or description
function parseBrand(title, description) {
  var focus = '';
  var m = title.match(/\|\s*(.+)$/);
  if (m) focus = m[1].trim();

  // Strip known promo suffixes to isolate brand
  var stripped = focus
    .replace(/\bGWP\b/gi, '')
    .replace(/\bBAGO\b/gi, '')
    .replace(/\bBOGO\b/gi, '')
    .replace(/\d+%\s*off/gi, '')
    .replace(/\bfree\b/gi, '')
    .trim();

  return stripped || '';
}

function stripHtml(html) {
  return html.replace(/<[^>]+>/g, ' ').replace(/\s{2,}/g, ' ').trim();
}

// ─── Template type detection ──────────────────────────────────────────────────

function detectTemplateType(title, description) {
  var text = (title + ' ' + description).toLowerCase();

  if (/\bgwp\b|gift with purchase/.test(text))       return 'gwp';
  if (/\bbago\b|\bbogo\b/.test(text))                return 'gwp';
  if (/\bslice\b/.test(text))                        return 'slice';
  if (/\bseamless\b/.test(text))                     return 'seamless';
  if (/\d+\s*%\s*off|\bpercent\s*off\b/.test(text)) return 'pctoff';
  if (/\$\d+\s*off|\bdollar(s)?\s*off\b/.test(text)) return 'pctoff';

  return 'content';
}

// ─── Template fill ────────────────────────────────────────────────────────────

// Replaces common placeholder tokens in the copied template doc.
// Tokens are case-insensitive and wrapped in [brackets] or {{braces}}.
function fillTemplate(doc, brief) {
  var body = doc.getBody();

  var replacements = [
    ['[DATE]',          brief.dates      || '[DATE]'],
    ['{{DATE}}',        brief.dates      || '[DATE]'],
    ['[PROMO DATE]',    brief.dates      || '[PROMO DATE]'],
    ['[SEND DATE]',     brief.dates      || '[SEND DATE]'],
    ['[PROMO FOCUS]',   brief.promoFocus || '[PROMO FOCUS]'],
    ['{{PROMO FOCUS}}', brief.promoFocus || '[PROMO FOCUS]'],
    ['[BRAND]',         brief.brand      || '[BRAND]'],
    ['{{BRAND}}',       brief.brand      || '[BRAND]'],
    ['[OFFER]',         brief.promoFocus || '[OFFER]'],
    ['{{OFFER}}',       brief.promoFocus || '[OFFER]'],
    ['[WRIKE LINK]',    brief.wrikeUrl],
    ['{{WRIKE LINK}}',  brief.wrikeUrl],
  ];

  replacements.forEach(function(pair) {
    try {
      body.replaceText(escapeRegex(pair[0]), pair[1]);
    } catch (_) {}
  });
}

function escapeRegex(str) {
  return str.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, '\\$&');
}

// ─── Web app HTML ─────────────────────────────────────────────────────────────

function getFormHtml(tokenSaved) {
  var templateOptions = Object.keys(TEMPLATES).map(function(key) {
    return '<option value="' + key + '">' + TEMPLATES[key].label + '</option>';
  }).join('');

  var tokenBanner = tokenSaved
    ? '<div class="token-ok">✓ Wrike connected — <a href="#" onclick="showSettings();return false;">update token</a></div>'
    : '<div class="token-warn">⚠ Wrike token not set. <a href="#" onclick="showSettings();return false;">Add token to get started</a></div>';

  var settingsPanel =
    '<div id="settings" style="display:' + (tokenSaved ? 'none' : 'block') + '">' +
    '<h2>Connect Wrike</h2>' +
    '<p class="sub">Create a permanent access token in Wrike: <strong>Profile → Apps &amp; Integrations → API → Create new token</strong>. Paste it below — it\'s saved to this script only.</p>' +
    '<label>Wrike API Token</label>' +
    '<input id="wrikeToken" type="password" placeholder="Paste token here…" />' +
    '<button id="saveBtn" onclick="saveToken()">Save Token</button>' +
    '<p id="tokenStatus"></p>' +
    '</div>';

  var mainForm =
    '<div id="main" style="display:' + (tokenSaved ? 'block' : 'none') + '">' +
    '<label>Brief <span>(Wrike link)</span></label>' +
    '<input id="wrikeUrl" type="url" placeholder="https://www.wrike.com/open.htm?id=..." />' +
    '<label>Template <span>(auto-detected from brief, or override)</span></label>' +
    '<select id="templateType">' +
    '<option value="auto">Auto-detect from brief</option>' +
    templateOptions +
    '</select>' +
    '<button id="btn" onclick="generate()">Generate Copy Doc</button>' +
    '<p id="status"></p>' +
    '<div id="result"></div>' +
    '<p id="errMsg" class="error"></p>' +
    '</div>';

  return '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<title>Thrive Market Email Copy Template Generator</title>' +
    '<style>' +
    'body{font-family:Google Sans,Arial,sans-serif;max-width:580px;margin:48px auto;padding:0 20px;color:#1f2937}' +
    'h1{color:#1B4332;font-size:22px;margin-bottom:4px}' +
    'h2{color:#1B4332;font-size:16px;margin:0 0 6px}' +
    'p.sub{color:#6b7280;font-size:13px;margin:0 0 20px}' +
    'label{display:block;font-size:13px;font-weight:600;margin:18px 0 5px}' +
    'label span{font-weight:400;color:#6b7280}' +
    'input,select{width:100%;box-sizing:border-box;padding:9px 11px;border:1px solid #d1d5db;border-radius:6px;font-size:14px;color:#111827}' +
    'select{background:#fff}' +
    'button{margin-top:20px;background:#1B4332;color:#fff;border:none;padding:11px 28px;border-radius:6px;font-size:15px;font-weight:600;cursor:pointer;width:100%}' +
    'button:disabled{opacity:.5;cursor:default}' +
    '.token-ok{font-size:13px;color:#166534;background:#f0fdf4;border:1px solid #bbf7d0;padding:8px 12px;border-radius:6px;margin-bottom:20px}' +
    '.token-ok a,.token-warn a{color:inherit;font-weight:600}' +
    '.token-warn{font-size:13px;color:#92400e;background:#fffbeb;border:1px solid #fde68a;padding:8px 12px;border-radius:6px;margin-bottom:20px}' +
    '#status{margin-top:18px;font-size:14px;color:#374151;min-height:20px}' +
    '#result{margin-top:14px;padding:14px 16px;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px;display:none}' +
    '#result a{color:#1B4332;font-weight:700;font-size:15px;text-decoration:none}' +
    '#result a:hover{text-decoration:underline}' +
    '.meta{font-size:12px;color:#6b7280;margin-top:6px}' +
    '.error{color:#b91c1c;margin-top:12px;font-size:14px}' +
    '#tokenStatus{font-size:13px;margin-top:12px}' +
    '</style></head><body>' +
    '<h1>Email Copy Template Generator</h1>' +
    '<p class="sub" style="margin-bottom:16px">Paste a Wrike brief link to generate a pre-filled email copy doc.</p>' +
    tokenBanner +
    settingsPanel +
    mainForm +
    '<script>' +
    'function showSettings(){' +
    '  document.getElementById("settings").style.display="block";' +
    '  document.getElementById("main").style.display="none";' +
    '}' +
    'function saveToken(){' +
    '  var t=document.getElementById("wrikeToken").value.trim();' +
    '  if(!t){alert("Please paste a token.");return;}' +
    '  document.getElementById("saveBtn").disabled=true;' +
    '  document.getElementById("tokenStatus").textContent="Saving…";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(r){' +
    '      if(r.error){document.getElementById("tokenStatus").textContent="Error: "+r.error;document.getElementById("saveBtn").disabled=false;return;}' +
    '      document.getElementById("settings").style.display="none";' +
    '      document.getElementById("main").style.display="block";' +
    '      document.querySelector(".token-warn") && (document.querySelector(".token-warn").outerHTML="<div class=\'token-ok\'>✓ Wrike connected — <a href=\'#\' onclick=\'showSettings();return false;\'>update token</a></div>");' +
    '      document.getElementById("saveBtn").disabled=false;' +
    '    })' +
    '    .withFailureHandler(function(e){' +
    '      document.getElementById("tokenStatus").textContent="Error: "+e.message;' +
    '      document.getElementById("saveBtn").disabled=false;' +
    '    })' +
    '    .saveWrikeToken(t);' +
    '}' +
    'function generate(){' +
    '  var url=document.getElementById("wrikeUrl").value.trim();' +
    '  var tmpl=document.getElementById("templateType").value;' +
    '  if(!url){alert("Please paste a Wrike brief URL.");return;}' +
    '  document.getElementById("btn").disabled=true;' +
    '  document.getElementById("result").style.display="none";' +
    '  document.getElementById("errMsg").textContent="";' +
    '  document.getElementById("status").textContent="Fetching brief from Wrike and building your copy doc…";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(r){' +
    '      document.getElementById("status").textContent="";' +
    '      document.getElementById("btn").disabled=false;' +
    '      if(r.error){document.getElementById("errMsg").textContent=r.error;return;}' +
    '      var div=document.getElementById("result");' +
    '      div.innerHTML="<a href=\'"+r.url+"\' target=\'_blank\'>↗ Open: "+escHtml(r.name)+"</a>"' +
    '        +"<div class=\'meta\'>Template: "+escHtml(r.templateUsed)+"</div>"' +
    '        +(r.briefData&&r.briefData.dates?"<div class=\'meta\'>Date: "+escHtml(r.briefData.dates)+"</div>":"")' +
    '        +(r.briefData&&r.briefData.promoFocus?"<div class=\'meta\'>Promo focus: "+escHtml(r.briefData.promoFocus)+"</div>":"");' +
    '      div.style.display="block";' +
    '    })' +
    '    .withFailureHandler(function(e){' +
    '      document.getElementById("status").textContent="";' +
    '      document.getElementById("errMsg").textContent="Error: "+e.message;' +
    '      document.getElementById("btn").disabled=false;' +
    '    })' +
    '    .processForm({wrikeUrl:url,templateType:tmpl});' +
    '}' +
    'function escHtml(s){return s?s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;"):""}' +
    '<\/script></body></html>';
}
