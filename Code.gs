/**
 * Thrive Market Email Copy Template Generator
 *
 * Setup (one-time, in the Apps Script editor):
 *   1. Get your permanent access token from Wrike:
 *      profile avatar -> Apps & Integrations -> open your app -> copy "Permanent access token"
 *   2. Temporarily add this to Code.gs, run it once, then delete it:
 *        function tempSetToken() {
 *          PropertiesService.getScriptProperties().setProperty('WRIKE_TOKEN', 'your-token-here');
 *        }
 *   3. Deploy as Web App (Execute as: Me, Who has access: Anyone within Thrive Market)
 *   4. Open the Web App URL, paste your Wrike brief folder URL in Settings, click Save Folder
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

// - Folder setting (called from the UI) -

function getBriefFolderUrl() {
  return PropertiesService.getScriptProperties().getProperty('BRIEF_FOLDER_URL') || '';
}

/**
 * Save the brief folder. Accepts either:
 *   (a) A Wrike folder URL  (https://www.wrike.com/open.htm?id=...)
 *   (b) An alphanumeric Wrike folder ID (e.g. IEADZYVOI46XHRTX)
 *
 * For (b), the ID is saved immediately. This is the recommended path --
 * run setupBriefFolderFromUrl() in the Script Editor to find your ID first.
 */
function saveBriefFolder(input) {
  input = (input || '').trim();
  if (!input) return { error: 'Please paste a Wrike folder URL or ID.' };

  // If it looks like an alphanumeric Wrike ID, save it directly
  if (/^[A-Z0-9]{10,}$/.test(input)) {
    PropertiesService.getScriptProperties().setProperty('BRIEF_FOLDER_ID', input);
    PropertiesService.getScriptProperties().setProperty('BRIEF_FOLDER_URL', input);
    return { ok: true };
  }

  PropertiesService.getScriptProperties().setProperty('BRIEF_FOLDER_URL', input);
  PropertiesService.getScriptProperties().deleteProperty('BRIEF_FOLDER_ID');

  return {
    error: 'Could not auto-resolve that URL. Run setupBriefFolderFromUrl() in the ' +
           'Apps Script editor with your folder URL to find and save the folder ID. ' +
           'See Code.gs setup instructions for details.'
  };
}

function getCachedFolderId_() {
  return PropertiesService.getScriptProperties().getProperty('BRIEF_FOLDER_ID') || '';
}

/**
 * ONE-TIME SETUP: Run this function in the Apps Script editor (not the web app).
 * It finds the alphanumeric Wrike ID for your brief folder and saves it.
 *
 * How to run:
 *   1. Paste your Wrike folder URL as the argument below (replace the placeholder)
 *   2. Click the Run button in the Apps Script editor
 *   3. Check the Execution Log -- it will print the found ID and save it automatically
 *   4. After this runs successfully, the web app will work for all future brief lookups
 *
 * Example: setupBriefFolderFromUrl('https://www.wrike.com/open.htm?id=819530475')
 */
function setupBriefFolderFromUrl(folderUrl) {
  folderUrl = (folderUrl || '').trim();
  if (!folderUrl) {
    Logger.log('ERROR: Pass your folder URL as the argument, e.g.:');
    Logger.log('  setupBriefFolderFromUrl("https://www.wrike.com/open.htm?id=819530475")');
    return;
  }

  var m = folderUrl.match(/[?&#]id=(\d+)/);
  if (!m) {
    Logger.log('ERROR: URL does not contain a numeric id parameter.');
    return;
  }
  var targetPermalink = 'https://www.wrike.com/open.htm?id=' + m[1];
  Logger.log('Searching for: ' + targetPermalink);

  var spaces = wrikeFetch_('/spaces').data || [];
  Logger.log('Found ' + spaces.length + ' spaces. Checking each...');

  for (var s = 0; s < spaces.length; s++) {
    var spaceName = spaces[s].title || spaces[s].id;
    Logger.log('Checking space: ' + spaceName);

    // Get all folders in this space (no fields param -- adding it breaks some spaces)
    var allFolders = wrikeFetch_('/spaces/' + spaces[s].id + '/folders').data || [];
    Logger.log('  ' + allFolders.length + ' folders. Batch-checking permalinks...');

    // Check in batches of 100 using the multi-get endpoint
    for (var i = 0; i < allFolders.length; i += 100) {
      var batch = allFolders.slice(i, i + 100).map(function(f) { return f.id; }).join(',');
      try {
        var result = wrikeFetch_('/folders/' + batch).data || [];
        for (var j = 0; j < result.length; j++) {
          if (result[j].permalink === targetPermalink) {
            var foundId = result[j].id;
            var foundTitle = result[j].title;
            Logger.log('FOUND: "' + foundTitle + '" ID=' + foundId);
            PropertiesService.getScriptProperties().setProperty('BRIEF_FOLDER_ID', foundId);
            PropertiesService.getScriptProperties().setProperty('BRIEF_FOLDER_URL', folderUrl);
            Logger.log('Saved to script properties. Setup complete!');
            return foundId;
          }
        }
      } catch (e) {
        Logger.log('  Batch error at index ' + i + ': ' + e.message.slice(0, 80));
      }
    }
  }

  Logger.log('NOT FOUND in any space. The folder may use a different URL format.');
  Logger.log('Try: setupBriefFolderById("YOUR_ALPHANUMERIC_ID") if you know the API ID.');
}

/**
 * Alternative one-time setup: use if you already know the alphanumeric Wrike folder ID.
 * Run in the Apps Script editor:  setupBriefFolderById('IEADZYVOI46XHRTX')
 */
function setupBriefFolderById(alphanumericId) {
  alphanumericId = (alphanumericId || '').trim();
  if (!alphanumericId) {
    Logger.log('ERROR: Pass the alphanumeric folder ID as the argument.');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('BRIEF_FOLDER_ID', alphanumericId);
  Logger.log('Saved folder ID: ' + alphanumericId + '. Setup complete!');
}

// - Web App entry point -

function doGet() {
  var html = getFormHtml_(hasWrikeToken(), getBriefFolderUrl());
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
  if (!getCachedFolderId_()) return { error: 'No brief folder saved yet. Open Settings, paste your folder URL, and click Save Folder.' };

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

// - Wrike brief lookup -

/**
 * Find a brief by its permalink URL within the saved folder.
 * Searches direct children, then one level deeper (month folders -> briefs).
 * Checks both folders (briefs stored as projects) and tasks.
 */
function fetchWrikeBrief_(wrikeUrl) {
  var targetPermalink = wrikeUrl.trim();
  var folderId        = getCachedFolderId_();

  if (!folderId) {
    return { error: 'No brief folder configured. Open Settings and save your folder URL.' };
  }

  var task = findByPermalinkInFolder_(folderId, targetPermalink);
  if (!task) {
    return { error: 'Could not find that brief in your configured folder. Make sure the link is from a brief inside that folder.' };
  }

  return extractBriefFromTask_(task, targetPermalink);
}

/**
 * Search for an item with the given permalink within a folder.
 * Checks: direct child folders, direct child tasks,
 *         then grandchild folders and tasks (one level deeper).
 *
 * Note: /folders/{id}/folders returns permalink by default -- no fields param needed.
 */
function findByPermalinkInFolder_(folderId, targetPermalink) {
  // Level 1: direct child folders (permalink returned by default)
  var childFolders = [];
  try { childFolders = wrikeFetch_('/folders/' + folderId + '/folders').data || []; } catch (_) {}

  for (var i = 0; i < childFolders.length; i++) {
    if (childFolders[i].permalink === targetPermalink) {
      return fetchFullItem_(childFolders[i].id, true);
    }
  }

  // Level 1: direct child tasks
  var childTasks = [];
  try { childTasks = wrikeFetch_('/folders/' + folderId + '/tasks').data || []; } catch (_) {}

  for (var i = 0; i < childTasks.length; i++) {
    if (childTasks[i].permalink === targetPermalink) {
      return fetchFullItem_(childTasks[i].id, false);
    }
  }

  // Level 2: grandchild folders and tasks (brief inside a month folder)
  for (var i = 0; i < childFolders.length; i++) {
    var grandFolders = [];
    try { grandFolders = wrikeFetch_('/folders/' + childFolders[i].id + '/folders').data || []; } catch (_) {}

    for (var j = 0; j < grandFolders.length; j++) {
      if (grandFolders[j].permalink === targetPermalink) {
        return fetchFullItem_(grandFolders[j].id, true);
      }
    }

    var grandTasks = [];
    try { grandTasks = wrikeFetch_('/folders/' + childFolders[i].id + '/tasks').data || []; } catch (_) {}

    for (var j = 0; j < grandTasks.length; j++) {
      if (grandTasks[j].permalink === targetPermalink) {
        return fetchFullItem_(grandTasks[j].id, false);
      }
    }
  }

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

function getFormHtml_(connected, savedFolderUrl) {
  var templateOptions = Object.keys(TEMPLATES).map(function(k) {
    return '<option value="' + k + '">' + TEMPLATES[k].label + '</option>';
  }).join('');

  var folderSaved  = !!getCachedFolderId_();
  var wrikeStatus  = !connected
    ? '<div class="pill warn">Wrike token not set. See setup instructions in Code.gs.</div>'
    : folderSaved
      ? '<div class="pill ok">Wrike connected &nbsp;&middot;&nbsp; <a href="#" onclick="showSettings();return false;">settings</a></div>'
      : '<div class="pill warn">Wrike connected, but no brief folder set &mdash; <a href="#" onclick="showSettings();return false;">open settings</a></div>';

  var settings =
    '<div id="settings" style="display:' + (folderSaved ? 'none' : 'block') + ';background:#f9fafb;border:1px solid #e5e7eb;border-radius:10px;padding:20px;margin-bottom:24px">' +
    '<h2 style="margin:0 0 8px;font-size:15px;color:#1B4332">Settings</h2>' +
    '<p class="sub"><strong>One-time setup required.</strong> To link your Wrike brief folder, do one of the following:</p>' +
    '<p class="sub" style="margin-top:-8px"><strong>Option A (recommended):</strong> In the Apps Script editor, run <code>setupBriefFolderFromUrl("YOUR_WRIKE_FOLDER_URL")</code>. This takes 1-3 minutes and auto-saves your folder.</p>' +
    '<p class="sub" style="margin-top:-8px"><strong>Option B:</strong> Paste your folder\'s alphanumeric Wrike ID below (looks like IEADZYVOI46XHRTX). You can find it by running <code>setupBriefFolderFromUrl()</code> in the editor first.</p>' +
    '<label>Alphanumeric Wrike Folder ID</label>' +
    '<input id="folderUrl" type="text" value="' + (savedFolderUrl || '') + '" placeholder="e.g. IEADZYVOI46XHRTX" />' +
    '<button id="saveBtn" onclick="saveFolder()">Save Folder ID</button>' +
    '<p id="saveStatus" style="font-size:13px;margin-top:10px;color:#374151;min-height:16px"></p>' +
    '</div>';

  return '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<title>Thrive Market Email Copy Template Generator</title>' +
    '<style>' +
    'body{font-family:Google Sans,Arial,sans-serif;max-width:580px;margin:48px auto;padding:0 20px;color:#1f2937}' +
    'h1{color:#1B4332;font-size:22px;margin-bottom:4px}' +
    'p.sub{color:#6b7280;font-size:13px;margin:0 0 16px;line-height:1.5}' +
    'label{display:block;font-size:13px;font-weight:600;margin:12px 0 4px}' +
    'span.light{font-weight:400;color:#6b7280}' +
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
    '<p class="sub">Paste a Wrike brief link to generate a pre-filled email copy doc.</p>' +
    wrikeStatus + settings +
    '<label>Wrike Brief URL</label>' +
    '<input id="wrikeUrl" type="url" placeholder="https://www.wrike.com/open.htm?id=..." />' +
    '<label>Template <span class="light">(auto-detected, or override)</span></label>' +
    '<select id="templateType"><option value="auto">Auto-detect from brief</option>' + templateOptions + '</select>' +
    '<button id="btn" onclick="generate()">Generate Copy Doc &rarr;</button>' +
    '<p id="status" style="margin-top:18px;font-size:14px;color:#374151;min-height:20px"></p>' +
    '<div id="result" style="display:none;margin-top:14px;padding:14px 16px;background:#f0fdf4;border:1px solid #bbf7d0;border-radius:8px"></div>' +
    '<p id="errMsg" style="color:#b91c1c;margin-top:12px;font-size:14px;min-height:16px"></p>' +
    '<script>' +
    'function showSettings(){document.getElementById("settings").style.display="block";}' +
    'function saveFolder(){' +
    '  var url=document.getElementById("folderUrl").value.trim();' +
    '  if(!url){alert("Please paste a Wrike folder URL.");return;}' +
    '  document.getElementById("saveBtn").disabled=true;' +
    '  document.getElementById("saveStatus").textContent="Resolving folder in Wrike... this takes about 20-30 seconds.";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(r){' +
    '      document.getElementById("saveBtn").disabled=false;' +
    '      if(r.error){document.getElementById("saveStatus").textContent="Error: "+r.error;return;}' +
    '      document.getElementById("saveStatus").textContent="Folder saved! You can now generate docs.";' +
    '      document.getElementById("settings").style.display="none";' +
    '    })' +
    '    .withFailureHandler(function(e){' +
    '      document.getElementById("saveBtn").disabled=false;' +
    '      document.getElementById("saveStatus").textContent="Error: "+e.message;' +
    '    })' +
    '    .saveBriefFolder(url);' +
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
