/**
 * Thrive Market Email Copy Template Generator
 *
 * Setup (one-time):
 *   1. Open the web app URL
 *   2. Click "Connect Wrike" and enter your Client ID + Client Secret
 *      (from Wrike -> profile -> Apps & Integrations -> open one of your apps)
 *   3. Click Authorize -- approve in the popup
 *   4. You're done. Start pasting Wrike brief links.
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
  if (!token) throw new Error('Wrike not connected. Open the Settings panel and authorize.');
  return token;
}

function saveWrikeToken_(token) {
  PropertiesService.getScriptProperties().setProperty('WRIKE_TOKEN', token);
}

// - Web App entry point -

function doGet(e) {
  // OAuth callback from Wrike: ?code=...&state=wrike_oauth
  if (e && e.parameter && e.parameter.code && e.parameter.state === 'wrike_oauth') {
    return handleOAuthCallback_(e.parameter.code);
  }
  var html = getFormHtml_(hasWrikeToken());
  return HtmlService.createHtmlOutput(html)
    .setTitle('Thrive Market Email Copy Template Generator')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// - OAuth flow -

function handleOAuthCallback_(code) {
  var props        = PropertiesService.getScriptProperties();
  var clientId     = props.getProperty('WRIKE_CLIENT_ID')     || '';
  var clientSecret = props.getProperty('WRIKE_CLIENT_SECRET') || '';
  var redirectUri  = ScriptApp.getService().getUrl();

  var resp;
  try {
    resp = UrlFetchApp.fetch('https://login.wrike.com/oauth2/token', {
      method: 'post',
      muteHttpExceptions: true,
      payload: {
        client_id:     clientId,
        client_secret: clientSecret,
        grant_type:    'authorization_code',
        code:          code,
        redirect_uri:  redirectUri,
      },
    });
  } catch (err) {
    return simplePage_('Connection failed', '<p class="err">Network error: ' + err.message + '</p>');
  }

  var data;
  try { data = JSON.parse(resp.getContentText()); } catch (_) { data = {}; }

  if (!data.access_token) {
    return simplePage_('Connection failed',
      '<p class="err">Wrike error: ' + (data.error_description || data.error || resp.getContentText().slice(0,200)) + '</p>' +
      '<p><a href="' + ScriptApp.getService().getUrl() + '">&larr; Try again</a></p>');
  }

  saveWrikeToken_(data.access_token);
  if (data.refresh_token) props.setProperty('WRIKE_REFRESH_TOKEN', data.refresh_token);

  return simplePage_('Wrike connected!',
    '<p class="ok">Authorization successful. You can close this tab.</p>' +
    '<p><a href="' + ScriptApp.getService().getUrl() + '">&larr; Open the generator</a></p>');
}

// Called from the UI to save credentials and get the auth URL
function saveCredentialsAndGetAuthUrl(clientId, clientSecret) {
  if (!clientId || !clientSecret) return { error: 'Both Client ID and Client Secret are required.' };

  var redirectUri = ScriptApp.getService().getUrl();
  var props = PropertiesService.getScriptProperties();
  props.setProperty('WRIKE_CLIENT_ID',    clientId.trim());
  props.setProperty('WRIKE_CLIENT_SECRET', clientSecret.trim());

  var authUrl = 'https://login.wrike.com/oauth2/authorize' +
    '?client_id='     + encodeURIComponent(clientId.trim()) +
    '&response_type=code' +
    '&redirect_uri='  + encodeURIComponent(redirectUri) +
    '&scope=Default' +
    '&state=wrike_oauth';

  return { authUrl: authUrl, redirectUri: redirectUri };
}

// Called from the UI on page load so the redirect URI can be displayed
function getRedirectUri() {
  return ScriptApp.getService().getUrl();
}

// Token refresh (called automatically on 401)
function refreshWrikeToken_() {
  var props        = PropertiesService.getScriptProperties();
  var refreshToken = props.getProperty('WRIKE_REFRESH_TOKEN') || '';
  var clientId     = props.getProperty('WRIKE_CLIENT_ID')     || '';
  var clientSecret = props.getProperty('WRIKE_CLIENT_SECRET') || '';
  if (!refreshToken || !clientId || !clientSecret) return false;

  var resp;
  try {
    resp = UrlFetchApp.fetch('https://login.wrike.com/oauth2/token', {
      method: 'post',
      muteHttpExceptions: true,
      payload: {
        client_id:     clientId,
        client_secret: clientSecret,
        grant_type:    'refresh_token',
        refresh_token: refreshToken,
      },
    });
  } catch (_) { return false; }

  var data;
  try { data = JSON.parse(resp.getContentText()); } catch (_) { return false; }
  if (!data.access_token) return false;

  saveWrikeToken_(data.access_token);
  if (data.refresh_token) props.setProperty('WRIKE_REFRESH_TOKEN', data.refresh_token);
  return true;
}

// - Main form handler -

function processForm(form) {
  var wrikeUrl         = (form.wrikeUrl     || '').trim();
  var templateOverride = (form.templateType || 'auto').trim();

  if (!wrikeUrl) return { error: 'Please paste a Wrike brief URL.' };
  if (wrikeUrl.indexOf('wrike.com') === -1) return { error: 'Please enter a valid Wrike URL.' };
  if (!hasWrikeToken()) return { error: 'Wrike not connected. Open Settings and authorize.' };

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
  if (code === 401) {
    if (!refreshWrikeToken_()) throw new Error('Wrike session expired. Re-authorize in Settings.');
    token = getWrikeToken_();
    resp  = UrlFetchApp.fetch(WRIKE_API_BASE + endpoint, {
      method: 'GET',
      headers: { 'Authorization': 'Bearer ' + token },
      muteHttpExceptions: true,
    });
    code = resp.getResponseCode();
  }
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
  return (months[parseInt(m[1],10) - 1] || '') + ' ' + parseInt(m[2], 10);
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

// - Utility -

function simplePage_(title, body) {
  return HtmlService.createHtmlOutput(
    '<html><head><meta charset="utf-8"><style>' +
    'body{font-family:Google Sans,sans-serif;max-width:520px;margin:60px auto;padding:0 20px;color:#1f2937}' +
    'h2{color:#1B4332}' +
    '.ok{color:#166534;background:#f0fdf4;border:1px solid #bbf7d0;padding:12px 16px;border-radius:8px}' +
    '.err{color:#b91c1c;background:#fef2f2;border:1px solid #fecaca;padding:12px 16px;border-radius:8px}' +
    'a{color:#1B4332;font-weight:600}' +
    '</style></head><body><h2>' + title + '</h2>' + body + '</body></html>'
  ).setTitle(title);
}

// - HTML UI -

function getFormHtml_(connected) {
  var templateOptions = Object.keys(TEMPLATES).map(function(k) {
    return '<option value="' + k + '">' + TEMPLATES[k].label + '</option>';
  }).join('');

  var banner = connected
    ? '<div class="pill ok">Wrike connected &nbsp;&middot;&nbsp; <a href="#" onclick="showSettings();return false;">re-authorize</a></div>'
    : '<div class="pill warn">Wrike not connected &mdash; <a href="#" onclick="showSettings();return false;">connect now</a></div>';

  var settings =
    '<div id="settings" style="display:' + (connected ? 'none' : 'block') + ';background:#f9fafb;border:1px solid #e5e7eb;border-radius:10px;padding:20px;margin-bottom:24px">' +
    '<h2 style="margin:0 0 12px">Connect Wrike</h2>' +
    '<p class="sub">In Wrike: <strong>profile avatar &rarr; Apps &amp; Integrations</strong> &rarr; open one of your apps &rarr; copy the <strong>Client ID</strong> and <strong>Client Secret</strong>.</p>' +
    '<p style="font-size:13px;font-weight:600;margin:0 0 4px">Step 1 &mdash; Add this Redirect URI to your Wrike app (before clicking Authorize):</p>' +
    '<div id="uriDisplay" style="font-size:12px;background:#f3f4f6;border:1px solid #d1d5db;padding:8px 10px;border-radius:6px;word-break:break-all;margin-bottom:4px;color:#111827;font-family:monospace">Loading...</div>' +
    '<p style="font-size:11px;color:#6b7280;margin:0 0 14px">In your Wrike app settings &rarr; Redirect URIs &rarr; paste that URL &rarr; Save.</p>' +
    '<p style="font-size:13px;font-weight:600;margin:0 0 4px">Step 2 &mdash; Enter your app credentials:</p>' +
    '<label>Client ID</label><input id="clientId" type="text" placeholder="Paste Client ID..." />' +
    '<label>Client Secret</label><input id="clientSecret" type="password" placeholder="Paste Client Secret..." />' +
    '<button id="authBtn" onclick="authorize()">Authorize with Wrike &rarr;</button>' +
    '<p id="authStatus" style="font-size:13px;margin-top:10px;color:#374151"></p>' +
    '</div>';

  var form =
    '<div id="main" style="display:' + (connected ? 'block' : 'none') + '">' +
    '<label>Wrike Brief URL</label>' +
    '<input id="wrikeUrl" type="url" placeholder="https://www.wrike.com/open.htm?id=..." />' +
    '<label>Template <span style="font-weight:400;color:#6b7280">(auto-detected, or override)</span></label>' +
    '<select id="templateType"><option value="auto">Auto-detect from brief</option>' + templateOptions + '</select>' +
    '<button id="btn" onclick="generate()">Generate Copy Doc &rarr;</button>' +
    '<p id="status" style="margin-top:18px;font-size:14px;color:#374151;min-height:20px"></p>' +
    '<div id="result" style="display:none;margin-top:14px;padding:14px 16px;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px"></div>' +
    '<p id="errMsg" style="color:#b91c1c;margin-top:12px;font-size:14px;min-height:16px"></p>' +
    '</div>';

  return '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<title>Thrive Market Email Copy Template Generator</title>' +
    '<style>' +
    'body{font-family:Google Sans,Arial,sans-serif;max-width:580px;margin:48px auto;padding:0 20px;color:#1f2937}' +
    'h1{color:#1B4332;font-size:22px;margin-bottom:4px}' +
    'h2{font-size:16px;font-weight:700;color:#1B4332}' +
    'p.sub{color:#6b7280;font-size:13px;margin:0 0 16px;line-height:1.5}' +
    'label{display:block;font-size:13px;font-weight:600;margin:12px 0 4px}' +
    'input,select{width:100%;box-sizing:border-box;padding:9px 11px;border:1px solid #d1d5db;border-radius:6px;font-size:14px;color:#111827;background:#fff}' +
    'button{margin-top:16px;background:#1B4332;color:#fff;border:none;padding:11px 24px;border-radius:6px;font-size:14px;font-weight:600;cursor:pointer;width:100%}' +
    'button:disabled{opacity:.5;cursor:default}' +
    '.pill{font-size:13px;padding:8px 12px;border-radius:6px;margin-bottom:20px}' +
    '.pill.ok{color:#166534;background:#f0fdf4;border:1px solid #bbf7d0}' +
    '.pill.warn{color:#92400e;background:#fffbeb;border:1px solid #fde68a}' +
    '.pill a,.pill.ok a{color:inherit;font-weight:700}' +
    '#result a{color:#1B4332;font-weight:700;font-size:15px;text-decoration:none}' +
    '#result a:hover{text-decoration:underline}' +
    '.meta{font-size:12px;color:#6b7280;margin-top:6px}' +
    '</style></head><body>' +
    '<h1>Email Copy Template Generator</h1>' +
    '<p class="sub">Paste a Wrike brief link to generate a pre-filled email copy doc.</p>' +
    banner + settings + form +
    '<script>' +
    'google.script.run.withSuccessHandler(function(uri){' +
    '  var el=document.getElementById("uriDisplay");if(el)el.textContent=uri;' +
    '}).getRedirectUri();' +
    'function showSettings(){document.getElementById("settings").style.display="block";}' +
    'function authorize(){' +
    '  var id=document.getElementById("clientId").value.trim();' +
    '  var secret=document.getElementById("clientSecret").value.trim();' +
    '  if(!id||!secret){alert("Enter both Client ID and Client Secret.");return;}' +
    '  document.getElementById("authBtn").disabled=true;' +
    '  document.getElementById("authStatus").textContent="Saving credentials...";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(r){' +
    '      if(r.error){document.getElementById("authStatus").innerHTML="Error: "+esc(r.error);document.getElementById("authBtn").disabled=false;return;}' +
    '      document.getElementById("authStatus").textContent="Redirecting to Wrike...";' +
    '      window.top.location.href=r.authUrl;' +
    '    })' +
    '    .withFailureHandler(function(e){' +
    '      document.getElementById("authStatus").textContent="Error: "+e.message;' +
    '      document.getElementById("authBtn").disabled=false;' +
    '    })' +
    '    .saveCredentialsAndGetAuthUrl(id,secret);' +
    '}' +
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
