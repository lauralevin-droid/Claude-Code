/**
 * Thrive Market Email Copy Template Generator
 *
 * Setup (one-time, in the Apps Script editor):
 *   1. Get your permanent access token from Wrike:
 *      profile avatar -> Apps & Integrations -> open your app -> copy "Permanent access token"
 *   2. Temporarily add this function to Code.gs, run it once, then delete it:
 *        function tempSetToken() {
 *          PropertiesService.getScriptProperties().setProperty('WRIKE_TOKEN', 'your-token-here');
 *        }
 *   3. Deploy as Web App (Execute as: Me, Who has access: Anyone within Thrive Market)
 *   4. Open the Web App URL and start generating docs
 */

// - Template IDs (Google Doc IDs) -

var TEMPLATES = {
  content:  { label: 'Content Send (non-promo)', id: '1f6uzg9TZCKMwJdSAUC7pmKyBQOYpWzUwUxAsnGGBMDg' },
  slice:    { label: 'Slice',                    id: '1z6j5We_agYn7ew8N1bbjUYSkZpjTeg2JCQQf2kGij5M' },
  seamless: { label: 'Seamless',                 id: '1DpQ1abtKielLnR1sF8eR6qkgpOCvoJwwsDGZe2vprd4' },
  gwp:      { label: 'GWP / BAGO',              id: '1RegQ-F-5AgcA-rlYP2UvS2wyUlNinysvEd-8d_ptndw' },
  pctoff:   { label: '% Off',                    id: '1RuW0XDoT48Hr3Sv08l9IBCtbBf6PZCLC8Gh73I2KxDY' },
};

var WRIKE_API_BASE = 'https://www.wrike.com/api/v4';

// - Token helpers -

function hasWrikeToken() {
  return !!PropertiesService.getScriptProperties().getProperty('WRIKE_TOKEN');
}

function getWrikeToken_() {
  var token = PropertiesService.getScriptProperties().getProperty('WRIKE_TOKEN');
  if (!token) throw new Error('Wrike token not set. See setup instructions in Code.gs.');
  return token;
}

// - Web App entry point -

function doGet() {
  var html = getFormHtml_(hasWrikeToken());
  return HtmlService.createHtmlOutput(html)
    .setTitle('Thrive Market Email Copy Template Generator')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// - Main form handler -

function processForm(form) {
  var wrikeUrl         = (form.wrikeUrl     || '').trim();
  var templateOverride = (form.templateType || 'auto').trim();

  if (!wrikeUrl) return { error: 'Please paste a Wrike brief URL.' };
  if (wrikeUrl.indexOf('wrike.com') === -1) return { error: 'Please enter a valid Wrike URL.' };
  if (!hasWrikeToken()) return { error: 'Wrike token not configured. See setup instructions in Code.gs.' };

  var brief = fetchWrikeBrief_(wrikeUrl);
  if (brief.error) return { error: brief.error };

  var templateKey = (templateOverride !== 'auto')
    ? templateOverride
    : detectTemplateType_(brief.title, brief.description);

  var template = TEMPLATES[templateKey] || TEMPLATES.content;
  var docName  = brief.title || ('Email Copy - ' + wrikeUrl);

  try {
    var newFile = DriveApp.getFileById(template.id).makeCopy(docName);
    var doc     = DocumentApp.openById(newFile.getId());
    fillTemplate_(doc, brief);
    doc.saveAndClose();
    return {
      url:          newFile.getUrl(),
      name:         docName,
      templateUsed: template.label,
      briefData:    brief,
    };
  } catch (err) {
    return { error: 'Could not copy template: ' + err.message };
  }
}

// - Wrike API: fast permalink lookup -

function fetchWrikeBrief_(wrikeUrl) {
  var url      = wrikeUrl.trim();
  var resolved = resolvePermalink_(url);
  if (!resolved) {
    return { error: 'Could not find that brief in Wrike. Make sure you copied the correct task link.' };
  }

  var task = fetchFullItem_(resolved.id, !!resolved._isFolder);
  if (!task) {
    return { error: 'Found the task in Wrike but could not load its details. Try again.' };
  }

  return extractBriefFromTask_(task, url);
}

function resolvePermalink_(wrikeUrl) {
  var encoded    = encodeURIComponent(wrikeUrl);
  var rawEncoded = wrikeUrl.replace(/&/g, '%26');

  try {
    var d1 = wrikeFetch_('/tasks?permalink=' + encoded);
    if (d1.data && d1.data[0]) return d1.data[0];
  } catch (e1) { /* fall through */ }

  try {
    var d2 = wrikeFetch_('/tasks?permalink=' + rawEncoded);
    if (d2.data && d2.data[0]) return d2.data[0];
  } catch (e2) { /* fall through */ }

  try {
    var d3 = wrikeFetch_('/folders?permalink=' + encoded);
    if (d3.data && d3.data[0]) {
      var folder = d3.data[0];
      folder._isFolder = true;
      return folder;
    }
  } catch (e3) { /* fall through */ }

  return null;
}

function fetchFullItem_(itemId, isFolder) {
  var base     = isFolder ? '/folders/' : '/tasks/';
  var attempts = [['description', 'customFields'], ['customFields'], []];

  for (var i = 0; i < attempts.length; i++) {
    try {
      var ep   = base + itemId;
      if (attempts[i].length) ep += '?fields=' + encodeURIComponent(JSON.stringify(attempts[i]));
      var data = wrikeFetch_(ep);
      if (data.data && data.data[0]) return data.data[0];
    } catch (e) {
      if (!e.message || e.message.indexOf('400') === -1) throw e;
    }
  }

  if (!isFolder) {
    try {
      var fd = wrikeFetch_('/folders/' + itemId);
      if (fd.data && fd.data[0]) return fd.data[0];
    } catch (_) {}
  }

  return null;
}

function wrikeFetch_(endpoint) {
  var token = getWrikeToken_();
  var resp  = UrlFetchApp.fetch(WRIKE_API_BASE + endpoint, {
    method: 'GET',
    headers: { 'Authorization': 'Bearer ' + token },
    muteHttpExceptions: true,
  });
  var code = resp.getResponseCode();
  if (code !== 200) throw new Error('Wrike API ' + code + ': ' + resp.getContentText().slice(0, 200));
  return JSON.parse(resp.getContentText());
}

// - Brief extraction -

function extractBriefFromTask_(task, sourceUrl) {
  var title       = task.title || '';
  var description = stripHtml_(task.description || '');
  return {
    title:       title,
    description: description,
    dates:       parseDatesFromTitle_(title) || parseDatesFromDescription_(description),
    promoFocus:  parsePromoFocus_(title, description),
    brand:       parseBrand_(title, description),
    wrikeUrl:    sourceUrl,
    taskId:      task.id,
    error:       null,
  };
}

function stripHtml_(html) {
  if (!html) return '';
  return html
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n')
    .replace(/<[^>]+>/g, ' ')
    .replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>')
    .replace(/&nbsp;/g, ' ').replace(/&#39;/g, "'").replace(/&quot;/g, '"')
    .replace(/\s{2,}/g, ' ').trim();
}

function parseDatesFromTitle_(title) {
  var m = title.match(/\b(\d{1,2})[./](\d{1,2})\b/);
  if (!m) return '';
  var months = ['January','February','March','April','May','June',
                'July','August','September','October','November','December'];
  return (months[parseInt(m[1], 10) - 1] || '') + ' ' + parseInt(m[2], 10);
}

function parseDatesFromDescription_(desc) {
  var m = desc.match(/\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2}/i);
  return m ? m[0] : '';
}

function parsePromoFocus_(title, description) {
  var m = title.match(/\|\s*(.+)$/);
  if (m) return m[1].trim();
  m = description.match(/(\d+%\s*off|GWP|BOGO|BAGO|gift with purchase|free shipping)/i);
  return m ? m[0] : '';
}

function parseBrand_(title, description) {
  var m = title.match(/\|\s*(.+)$/);
  if (!m) return '';
  return m[1].trim()
    .replace(/\bGWP\b/gi, '').replace(/\bBAGO\b/gi, '').replace(/\bBOGO\b/gi, '')
    .replace(/\d+%\s*off/gi, '').replace(/\bfree\b/gi, '').trim();
}

// - Template detection and fill -

function detectTemplateType_(title, description) {
  var text = (title + ' ' + description).toLowerCase();
  if (/\bgwp\b|gift with purchase/.test(text))        return 'gwp';
  if (/\bbago\b|\bbogo\b/.test(text))                 return 'gwp';
  if (/\bslice\b/.test(text))                         return 'slice';
  if (/\bseamless\b/.test(text))                      return 'seamless';
  if (/\d+\s*%\s*off|\bpercent\s*off\b/.test(text))  return 'pctoff';
  if (/\$\d+\s*off|\bdollar(s)?\s*off\b/.test(text)) return 'pctoff';
  return 'content';
}

function fillTemplate_(doc, brief) {
  var body  = doc.getBody();
  var pairs = [
    ['[DATE]',        brief.dates      || '[DATE]'],
    ['{{DATE}}',      brief.dates      || '[DATE]'],
    ['[PROMO DATE]',  brief.dates      || '[PROMO DATE]'],
    ['[SEND DATE]',   brief.dates      || '[SEND DATE]'],
    ['[PROMO FOCUS]', brief.promoFocus || '[PROMO FOCUS]'],
    ['[BRAND]',       brief.brand      || '[BRAND]'],
    ['[OFFER]',       brief.promoFocus || '[OFFER]'],
    ['[WRIKE LINK]',  brief.wrikeUrl   || ''],
  ];
  pairs.forEach(function(p) {
    try { body.replaceText(escapeRegex_(p[0]), p[1]); } catch (_) {}
  });
}

function escapeRegex_(str) {
  return str.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, '\\$&');
}

// - HTML UI -

function getFormHtml_(connected) {
  var templateOptions = Object.keys(TEMPLATES).map(function(k) {
    return '<option value="' + k + '">' + TEMPLATES[k].label + '</option>';
  }).join('');

  var banner = connected
    ? '<div class="pill ok">Wrike connected</div>'
    : '<div class="pill warn">Wrike token not set. See setup instructions in Code.gs.</div>';

  return '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<title>Thrive Market Email Copy Template Generator</title>' +
    '<style>' +
    'body{font-family:Google Sans,Arial,sans-serif;max-width:580px;margin:48px auto;padding:0 20px;color:#1f2937}' +
    'h1{color:#1B4332;font-size:22px;margin-bottom:4px}' +
    'p.sub{color:#6b7280;font-size:13px;margin:0 0 20px;line-height:1.5}' +
    'label{display:block;font-size:13px;font-weight:600;margin:16px 0 4px}' +
    'span.light{font-weight:400;color:#6b7280}' +
    'input,select{width:100%;box-sizing:border-box;padding:9px 11px;border:1px solid #d1d5db;border-radius:6px;font-size:14px;color:#111827;background:#fff}' +
    'button{margin-top:16px;background:#1B4332;color:#fff;border:none;padding:11px 24px;border-radius:6px;font-size:14px;font-weight:600;cursor:pointer;width:100%}' +
    'button:disabled{opacity:.5;cursor:default}' +
    '.pill{font-size:13px;padding:8px 12px;border-radius:6px;margin-bottom:20px}' +
    '.pill.ok{color:#166534;background:#f0fdf4;border:1px solid #bbf7d0}' +
    '.pill.warn{color:#92400e;background:#fffbeb;border:1px solid #fde68a}' +
    '#result a{color:#1B4332;font-weight:700;font-size:15px;text-decoration:none}' +
    '#result a:hover{text-decoration:underline}' +
    '.meta{font-size:12px;color:#6b7280;margin-top:6px}' +
    '</style></head><body>' +
    '<h1>Email Copy Template Generator</h1>' +
    '<p class="sub">Paste a Wrike brief link to generate a pre-filled email copy doc.</p>' +
    banner +
    '<label>Wrike Brief URL</label>' +
    '<input id="wrikeUrl" type="url" placeholder="https://www.wrike.com/open.htm?id=..." />' +
    '<label>Template <span class="light">(auto-detected, or override)</span></label>' +
    '<select id="templateType"><option value="auto">Auto-detect from brief</option>' + templateOptions + '</select>' +
    '<button id="btn" onclick="generate()">Generate Copy Doc &rarr;</button>' +
    '<p id="status" style="margin-top:18px;font-size:14px;color:#374151;min-height:20px"></p>' +
    '<div id="result" style="display:none;margin-top:14px;padding:14px 16px;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px"></div>' +
    '<p id="errMsg" style="color:#b91c1c;margin-top:12px;font-size:14px;min-height:16px"></p>' +
    '<script>' +
    'function generate(){' +
    '  var url=document.getElementById("wrikeUrl").value.trim();' +
    '  var tmpl=document.getElementById("templateType").value;' +
    '  if(!url){alert("Please paste a Wrike brief URL.");return;}' +
    '  document.getElementById("btn").disabled=true;' +
    '  document.getElementById("result").style.display="none";' +
    '  document.getElementById("errMsg").textContent="";' +
    '  document.getElementById("status").textContent="Fetching brief from Wrike... (10-20 seconds)";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(r){' +
    '      document.getElementById("status").textContent="";' +
    '      document.getElementById("btn").disabled=false;' +
    '      if(r.error){document.getElementById("errMsg").textContent=r.error;return;}' +
    '      var d=document.getElementById("result");' +
    '      d.innerHTML=' +
    '        "<a href=\'"+r.url+"\' target=\'_blank\'>Open: "+esc(r.name)+" &rarr;</a>"' +
    '        +"<div class=\'meta\'>Template: "+esc(r.templateUsed)+"</div>"' +
    '        +(r.briefData&&r.briefData.dates?"<div class=\'meta\'>Date: "+esc(r.briefData.dates)+"</div>":"")' +
    '        +(r.briefData&&r.briefData.promoFocus?"<div class=\'meta\'>Promo: "+esc(r.briefData.promoFocus)+"</div>":"");' +
    '      d.style.display="block";' +
    '    })' +
    '    .withFailureHandler(function(e){' +
    '      document.getElementById("status").textContent="";' +
    '      document.getElementById("errMsg").textContent="Error: "+e.message;' +
    '      document.getElementById("btn").disabled=false;' +
    '    })' +
    '    .processForm({wrikeUrl:url,templateType:tmpl});' +
    '}' +
    'function esc(s){return s?s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;"):""}' +
    '<\/script></body></html>';
}
