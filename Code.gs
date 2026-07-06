/**
 * Thrive Market Email Copy Template Generator
 *
 * Setup (one-time, in the web app):
 *   1. Deploy as Web App (Execute as: Me, Who has access: Anyone within Thrive Market)
 *   2. Open the web app URL -> click "Connect Wrike"
 *   3. Enter the Client ID and Client Secret from your Wrike app
 *      (wrike.com -> profile avatar -> Apps & Integrations -> API -> Create new app)
 *   4. Click "Authorize with Wrike" -> approve in the popup -> done
 */

// Template IDs (Google Doc IDs from template URLs) -

var TEMPLATES = {
  content:  { label: 'Content Send (non-promo)', id: '1f6uzg9TZCKMwJdSAUC7pmKyBQOYpWzUwUxAsnGGBMDg' },
  slice:    { label: 'Slice',                    id: '1z6j5We_agYn7ew8N1bbjUYSkZpjTeg2JCQQf2kGij5M' },
  seamless: { label: 'Seamless',                 id: '1DpQ1abtKielLnR1sF8eR6qkgpOCvoJwwsDGZe2vprd4' },
  gwp:      { label: 'GWP / BAGO',              id: '1RegQ-F-5AgcA-rlYP2UvS2wyUlNinysvEd-8d_ptndw' },
  pctoff:   { label: '% Off',                    id: '1RuW0XDoT48Hr3Sv08l9IBCtbBf6PZCLC8Gh73I2KxDY' },
};

// Web App entry points -

function doGet(e) {
  // Wrike OAuth2 callback: ?code=...&state=wrike_oauth
  if (e && e.parameter && e.parameter.code && e.parameter.state === 'wrike_oauth') {
    return handleOAuthCallback(e.parameter.code);
  }
  var html = getFormHtml(hasWrikeToken());
  return HtmlService.createHtmlOutput(html)
    .setTitle('Thrive Market Email Copy Template Generator')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Exchanges authorization code for access + refresh tokens
function handleOAuthCallback(code) {
  var props = PropertiesService.getScriptProperties();
  var clientId     = props.getProperty('WRIKE_CLIENT_ID')     || '';
  var clientSecret = props.getProperty('WRIKE_CLIENT_SECRET') || '';
  var redirectUri  = props.getProperty('WRIKE_REDIRECT_URI')  || ScriptApp.getService().getUrl();

  var response;
  try {
    response = UrlFetchApp.fetch('https://login.wrike.com/oauth2/token', {
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
  } catch (e) {
    return htmlPage('Connection failed', '<p class="err">Network error: ' + e.message + '</p>');
  }

  var data;
  try { data = JSON.parse(response.getContentText()); } catch (_) { data = {}; }

  if (!data.access_token) {
    return htmlPage('Connection failed',
      '<p class="err">Wrike returned an error: ' + (data.error_description || data.error || response.getContentText()) + '</p>' +
      '<p><a href="' + ScriptApp.getService().getUrl() + '">&larr; Try again</a></p>');
  }

  props.setProperty('WRIKE_ACCESS_TOKEN',  data.access_token);
  props.setProperty('WRIKE_REFRESH_TOKEN', data.refresh_token || '');

  return htmlPage('Wrike connected!',
    '<p class="ok">Wrike authorization was successful.</p>' +
    '<p><a href="' + ScriptApp.getService().getUrl() + '">&larr; Open the generator</a></p>');
}

function htmlPage(title, body) {
  return HtmlService.createHtmlOutput(
    '<html><head><meta charset="utf-8"><style>' +
    'body{font-family:Google Sans,sans-serif;max-width:520px;margin:60px auto;padding:0 20px;color:#1f2937}' +
    'h2{color:#1B4332}.ok{color:#166534;background:#f0fdf4;border:1px solid #bbf7d0;padding:12px 16px;border-radius:8px}' +
    '.err{color:#b91c1c;background:#fef2f2;border:1px solid #fecaca;padding:12px 16px;border-radius:8px}' +
    'a{color:#1B4332;font-weight:600}</style></head><body>' +
    '<h2>' + title + '</h2>' + body + '</body></html>'
  ).setTitle(title);
}

// Returns the redirect URI this script uses -- must match exactly in the Wrike app settings
// --- Called from the HTML form -------------------------------------------

// Returns the canonical redirect URI this script will use -- always server-side.
function getRedirectUri() {
  return ScriptApp.getService().getUrl();
}

// Save Client ID + Secret and build the Wrike authorization URL.
// redirectUri is always derived server-side to guarantee consistency.
function saveCredentialsAndGetAuthUrl(clientId, clientSecret) {
  if (!clientId || !clientSecret) return { error: 'Both Client ID and Client Secret are required.' };

  var redirectUri = ScriptApp.getService().getUrl();
  var props = PropertiesService.getScriptProperties();
  props.setProperty('WRIKE_CLIENT_ID',     clientId.trim());
  props.setProperty('WRIKE_CLIENT_SECRET', clientSecret.trim());
  props.setProperty('WRIKE_REDIRECT_URI',  redirectUri);

  var authUrl = 'https://login.wrike.com/oauth2/authorize' +
    '?client_id='    + encodeURIComponent(clientId.trim()) +
    '&response_type=code' +
    '&redirect_uri=' + encodeURIComponent(redirectUri) +
    '&scope=Default' +
    '&state=wrike_oauth';

  return { authUrl: authUrl, redirectUri: redirectUri };
}

function processForm(form) {
  var wrikeUrl         = (form.wrikeUrl      || '').trim();
  var templateOverride = (form.templateType  || 'auto').trim();

  if (!wrikeUrl) return { error: 'Please paste a Wrike brief URL.' };
  if (!wrikeUrl.includes('wrike.com')) return { error: 'Please enter a valid Wrike URL.' };

  var brief = fetchWrikeBrief(wrikeUrl);
  if (brief.error) return { error: brief.error };

  var templateKey = templateOverride !== 'auto'
    ? templateOverride
    : detectTemplateType(brief.title, brief.description);

  var template = TEMPLATES[templateKey] || TEMPLATES.content;
  var docName  = brief.title || 'Email Copy - ' + wrikeUrl;

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
  } catch (err) {
    return { error: 'Could not copy template: ' + err.message };
  }
}

// Wrike token management -

function hasWrikeToken() {
  return !!PropertiesService.getScriptProperties().getProperty('WRIKE_ACCESS_TOKEN');
}

function getWrikeToken() {
  return PropertiesService.getScriptProperties().getProperty('WRIKE_ACCESS_TOKEN') || '';
}

// Refreshes the access token using the stored refresh token
function refreshWrikeToken() {
  var props        = PropertiesService.getScriptProperties();
  var refreshToken = props.getProperty('WRIKE_REFRESH_TOKEN') || '';
  var clientId     = props.getProperty('WRIKE_CLIENT_ID')     || '';
  var clientSecret = props.getProperty('WRIKE_CLIENT_SECRET') || '';

  if (!refreshToken || !clientId || !clientSecret) return false;

  var response;
  try {
    response = UrlFetchApp.fetch('https://login.wrike.com/oauth2/token', {
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
  try { data = JSON.parse(response.getContentText()); } catch (_) { return false; }

  if (!data.access_token) return false;

  props.setProperty('WRIKE_ACCESS_TOKEN', data.access_token);
  if (data.refresh_token) props.setProperty('WRIKE_REFRESH_TOKEN', data.refresh_token);
  return true;
}

// Wrike API -

function fetchWrikeBrief(wrikeUrl) {
  if (!hasWrikeToken()) {
    return { error: 'Wrike is not connected. Open the Settings panel and authorize.' };
  }

  var numericId = parseWrikeTaskId(wrikeUrl);
  if (!numericId) return { error: 'Could not parse a task ID from: ' + wrikeUrl };

  // Strategy 1: follow the open.htm redirect to get the real Wrike API task ID
  var apiTaskId = resolveTaskIdViaRedirect(numericId);
  if (apiTaskId) {
    var r1 = callWrikeApi('https://www.wrike.com/api/v4/tasks/' + apiTaskId);
    if (!r1.error && r1.data && r1.data.length > 0) {
      return extractBriefFromTask(r1.data[0], wrikeUrl);
    }
  }

  // Strategy 2: permalink search with the full open.htm URL
  var r2 = callWrikeApi(
    'https://www.wrike.com/api/v4/tasks?permalink=' + encodeURIComponent(wrikeUrl)
  );
  if (!r2.error && r2.data && r2.data.length > 0) {
    return extractBriefFromTask(r2.data[0], wrikeUrl);
  }

  // Build a diagnostic message from what the API actually returned
  var diag = r2.error
    ? r2.error
    : ('API returned 0 tasks. Raw: ' + JSON.stringify(r2).substring(0, 200));
  return {
    error: 'Could not find that task in Wrike (' + diag + '). ' +
      'Make sure the Wrike app has access to the workspace containing this task.'
  };
}

// Follow the open.htm short-link (without redirects) to extract the Wrike API task ID
function resolveTaskIdViaRedirect(numericId) {
  try {
    var resp = UrlFetchApp.fetch(
      'https://www.wrike.com/open.htm?id=' + numericId,
      {
        muteHttpExceptions: true,
        followRedirects: false,
        headers: { 'Authorization': 'Bearer ' + getWrikeToken() },
      }
    );
    var location = resp.getHeaders()['Location'] || resp.getHeaders()['location'] || '';
    // Workspace URLs contain the task API ID: ...&id=IEAAAAAAKQAAAABY...
    var m = location.match(/[?&#]id=([A-Z0-9]{10,})/);
    return m ? m[1] : null;
  } catch (_) {
    return null;
  }
}

// Makes a GET request; retries once with a refreshed token on 401
function callWrikeApi(url) {
  var token    = getWrikeToken();
  var response = wrikeGet(url, token);

  if (response.getResponseCode() === 401) {
    if (!refreshWrikeToken()) {
      return { error: 'Wrike token expired and could not be refreshed. Re-authorize in Settings.' };
    }
    response = wrikeGet(url, getWrikeToken());
  }

  if (response.getResponseCode() !== 200) {
    return { error: 'Wrike API returned HTTP ' + response.getResponseCode() };
  }

  var data;
  try { data = JSON.parse(response.getContentText()); } catch (_) {
    return { error: 'Could not parse Wrike API response.' };
  }
  return data;
}

function wrikeGet(url, token) {
  return UrlFetchApp.fetch(url, {
    muteHttpExceptions: true,
    headers: { 'Authorization': 'Bearer ' + token },
  });
}

function parseWrikeTaskId(url) {
  var m = url.match(/[?&]id=(\d+)/);
  if (m) return m[1];
  m = url.match(/[#&]id=(\d+)/);
  if (m) return m[1];
  return null;
}

function extractBriefFromTask(task, sourceUrl) {
  var title       = task.title || '';
  var description = stripHtml(task.description || '');
  return {
    title:      title,
    description: description,
    dates:      parseDatesFromTitle(title) || parseDatesFromDescription(description),
    promoFocus: parsePromoFocus(title, description),
    brand:      parseBrand(title, description),
    wrikeUrl:   sourceUrl,
    taskId:     task.id,
    error:      null,
  };
}

// Brief parsing -

function parseDatesFromTitle(title) {
  var m = title.match(/\b(\d{1,2})[./](\d{1,2})\b/);
  if (!m) return '';
  var months = ['January','February','March','April','May','June',
                'July','August','September','October','November','December'];
  var month = parseInt(m[1], 10);
  var day   = parseInt(m[2], 10);
  return (months[month - 1] || '') + ' ' + day;
}

function parseDatesFromDescription(desc) {
  var m = desc.match(/\b(January|February|March|April|May|June|July|August|September|October|November|December)\s+\d{1,2}/i);
  return m ? m[0] : '';
}

function parsePromoFocus(title, description) {
  var m = title.match(/\|\s*(.+)$/);
  if (m) return m[1].trim();
  m = description.match(/(\d+%\s*off|GWP|BOGO|BAGO|gift with purchase|free shipping)/i);
  return m ? m[0] : '';
}

function parseBrand(title, description) {
  var m = title.match(/\|\s*(.+)$/);
  if (!m) return '';
  return m[1].trim()
    .replace(/\bGWP\b/gi, '').replace(/\bBAGO\b/gi, '').replace(/\bBOGO\b/gi, '')
    .replace(/\d+%\s*off/gi, '').replace(/\bfree\b/gi, '').trim();
}

function stripHtml(html) {
  return html.replace(/<[^>]+>/g, ' ').replace(/\s{2,}/g, ' ').trim();
}

// Template detection & fill -

function detectTemplateType(title, description) {
  var text = (title + ' ' + description).toLowerCase();
  if (/\bgwp\b|gift with purchase/.test(text))        return 'gwp';
  if (/\bbago\b|\bbogo\b/.test(text))                 return 'gwp';
  if (/\bslice\b/.test(text))                         return 'slice';
  if (/\bseamless\b/.test(text))                      return 'seamless';
  if (/\d+\s*%\s*off|\bpercent\s*off\b/.test(text))  return 'pctoff';
  if (/\$\d+\s*off|\bdollar(s)?\s*off\b/.test(text)) return 'pctoff';
  return 'content';
}

function fillTemplate(doc, brief) {
  var body = doc.getBody();
  var pairs = [
    ['[DATE]',         brief.dates      || '[DATE]'],
    ['{{DATE}}',       brief.dates      || '[DATE]'],
    ['[PROMO DATE]',   brief.dates      || '[PROMO DATE]'],
    ['[SEND DATE]',    brief.dates      || '[SEND DATE]'],
    ['[PROMO FOCUS]',  brief.promoFocus || '[PROMO FOCUS]'],
    ['[BRAND]',        brief.brand      || '[BRAND]'],
    ['[OFFER]',        brief.promoFocus || '[OFFER]'],
    ['[WRIKE LINK]',   brief.wrikeUrl],
  ];
  pairs.forEach(function(p) {
    try { body.replaceText(escapeRegex(p[0]), p[1]); } catch (_) {}
  });
}

function escapeRegex(str) {
  return str.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, '\\$&');
}

// HTML -

function getFormHtml(connected) {
  var templateOptions = Object.keys(TEMPLATES).map(function(k) {
    return '<option value="' + k + '">' + TEMPLATES[k].label + '</option>';
  }).join('');

  var banner = connected
    ? '<div class="pill ok">Wrike connected &nbsp;&middot;&nbsp; <a href="#" onclick="showSettings();return false;">re-authorize</a></div>'
    : '<div class="pill warn">Wrike not connected &mdash; <a href="#" onclick="showSettings();return false;">connect now</a></div>';

  var settings =
    '<div id="settings" style="display:' + (connected ? 'none' : 'block') + ';background:#f9fafb;border:1px solid #e5e7eb;border-radius:10px;padding:20px;margin-bottom:24px">' +
    '<h2 style="margin:0 0 4px">Connect Wrike</h2>' +
    '<p class="sub">In Wrike: <strong>profile avatar &rarr; Apps &amp; Integrations &rarr; API &rarr; Create new app</strong>. ' +
    'Follow the steps below in order.</p>' +
    '<p style="font-size:13px;font-weight:600;margin:0 0 4px">Step 1 &mdash; Register this Redirect URI in your Wrike app <em>before</em> clicking Authorize:</p>' +
    '<div id="uriDisplay" style="font-size:12px;background:#f3f4f6;border:1px solid #d1d5db;padding:8px 10px;border-radius:6px;word-break:break-all;margin-bottom:4px;color:#111827;font-family:monospace">Loading...</div>' +
    '<p style="font-size:11px;color:#6b7280;margin:0 0 12px">Wrike app &rarr; Redirect URIs &rarr; paste exactly as shown &rarr; Save.</p>' +
    '<p style="font-size:13px;font-weight:600;margin:0 0 4px">Step 2 &mdash; Paste your app credentials:</p>' +
    '<label>Client ID</label><input id="clientId" type="text" placeholder="e.g. XXXXXXXXXXXXXXXX" />' +
    '<label>Client Secret</label><input id="clientSecret" type="password" placeholder="Paste client secret..." />' +
    '<button id="authBtn" onclick="authorize()" style="margin-top:16px">Authorize with Wrike &rarr;</button>' +
    '<p id="authStatus" style="font-size:13px;margin-top:10px;color:#374151"></p>' +
    '</div>';

  var form =
    '<div id="main" style="display:' + (connected ? 'block' : 'none') + '">' +
    '<label>Brief <span style="font-weight:400;color:#6b7280">(Wrike link)</span></label>' +
    '<input id="wrikeUrl" type="url" placeholder="https://www.wrike.com/open.htm?id=..." />' +
    '<label>Template <span style="font-weight:400;color:#6b7280">(auto-detected, or override)</span></label>' +
    '<select id="templateType"><option value="auto">Auto-detect from brief</option>' + templateOptions + '</select>' +
    '<button id="btn" onclick="generate()">Generate Copy Doc</button>' +
    '<p id="status" style="margin-top:18px;font-size:14px;color:#374151;min-height:20px"></p>' +
    '<div id="result" style="display:none;margin-top:14px;padding:14px 16px;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px"></div>' +
    '<p id="errMsg" style="color:#b91c1c;margin-top:12px;font-size:14px"></p>' +
    '</div>';

  return '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<title>Thrive Market Email Copy Template Generator</title>' +
    '<style>' +
    'body{font-family:Google Sans,Arial,sans-serif;max-width:580px;margin:48px auto;padding:0 20px;color:#1f2937}' +
    'h1{color:#1B4332;font-size:22px;margin-bottom:4px}' +
    'h2{font-size:15px;font-weight:700}' +
    'p.sub{color:#6b7280;font-size:13px;margin:0 0 16px;line-height:1.5}' +
    'label{display:block;font-size:13px;font-weight:600;margin:16px 0 4px}' +
    'input,select{width:100%;box-sizing:border-box;padding:9px 11px;border:1px solid #d1d5db;border-radius:6px;font-size:14px;color:#111827;background:#fff}' +
    'button{margin-top:16px;background:#1B4332;color:#fff;border:none;padding:11px 24px;border-radius:6px;font-size:14px;font-weight:600;cursor:pointer;width:100%}' +
    'button:disabled{opacity:.5;cursor:default}' +
    '.pill{font-size:13px;padding:8px 12px;border-radius:6px;margin-bottom:20px}' +
    '.pill.ok{color:#166534;background:#f0fdf4;border:1px solid #bbf7d0}' +
    '.pill.warn{color:#92400e;background:#fffbeb;border:1px solid #fde68a}' +
    '.pill a{color:inherit;font-weight:700}' +
    '#result a{color:#1B4332;font-weight:700;font-size:15px;text-decoration:none}' +
    '#result a:hover{text-decoration:underline}' +
    '.meta{font-size:12px;color:#6b7280;margin-top:6px}' +
    '</style></head><body>' +
    '<h1>Email Copy Template Generator</h1>' +
    '<p class="sub" style="margin-bottom:16px">Paste a Wrike brief link to generate a pre-filled email copy doc.</p>' +
    banner + settings + form +
    '<script>' +
    // Load the canonical redirect URI from the server and display it
    'google.script.run.withSuccessHandler(function(uri){' +
    '  var el=document.getElementById("uriDisplay");' +
    '  if(el)el.textContent=uri;' +
    '}).getRedirectUri();' +
    'function showSettings(){' +
    '  document.getElementById("settings").style.display="block";' +
    '}' +
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
    '      window.location.href=r.authUrl;' +
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
    '  document.getElementById("status").textContent="Fetching brief from Wrike and building your copy doc...";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(r){' +
    '      document.getElementById("status").textContent="";' +
    '      document.getElementById("btn").disabled=false;' +
    '      if(r.error){document.getElementById("errMsg").textContent=r.error;return;}' +
    '      var d=document.getElementById("result");' +
    '      d.innerHTML="<a href=\'"+r.url+"\' target=\'_blank\'>Open: "+esc(r.name)+" &rarr;</a>"' +
    '        +"<div class=\'meta\'>Template: "+esc(r.templateUsed)+"</div>"' +
    '        +(r.briefData&&r.briefData.dates?"<div class=\'meta\'>Date: "+esc(r.briefData.dates)+"</div>":"")' +
    '        +(r.briefData&&r.briefData.promoFocus?"<div class=\'meta\'>Promo focus: "+esc(r.briefData.promoFocus)+"</div>":"");' +
    '      d.style.display="block";' +
    '    })' +
    '    .withFailureHandler(function(e){' +
    '      document.getElementById("status").textContent="";' +
    '      document.getElementById("errMsg").textContent="Error: "+e.message;' +
    '      document.getElementById("btn").disabled=false;' +
    '    })' +
    '    .processForm({wrikeUrl:url,templateType:tmpl});' +
    // authorize() call needs the redirectUri arg
    '' +
    '}' +
    'function esc(s){return s?s.replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;"):""}' +
    '<\/script></body></html>';
}
