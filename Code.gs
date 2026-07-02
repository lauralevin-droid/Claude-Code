/**
 * Thrive Market Product Promotion Brief Generator
 *
 * Two ways to run:
 *   1. From the Apps Script editor: select createPromoBrief and click Run.
 *      Uses Browser.inputBox() — no spreadsheet or doc required.
 *   2. As a Web App: deploy as web app (Execute as: Me, Who has access: Anyone).
 *      Opens a form in the browser; submit creates the doc and returns a link.
 */

// ─── Web App entry point ──────────────────────────────────────────────────────

function doGet() {
  return HtmlService.createHtmlOutput(getFormHtml())
    .setTitle("Thrive Market Promo Brief Generator")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** Called from the HTML form via google.script.run */
function processForm(form) {
  const productUrl = (form.productUrl || "").trim();
  const discount   = (form.discount   || "").trim();
  const promoDates = (form.promoDates || "").trim();

  if (!productUrl.includes("thrivemarket.com")) {
    return { error: "Please enter a valid Thrive Market product URL." };
  }

  const product = fetchProductData(productUrl);
  const doc     = buildDoc(product, productUrl, discount, promoDates);
  return { url: doc.getUrl(), name: doc.getName() };
}

// ─── Editor / standalone entry point ─────────────────────────────────────────

function createPromoBrief() {
  const productUrl = Browser.inputBox(
    "Thrive Market Promo Brief Generator",
    "Paste the Thrive Market product URL:",
    Browser.Buttons.OK_CANCEL
  );
  if (productUrl === "cancel" || productUrl === "") return;

  if (!productUrl.includes("thrivemarket.com")) {
    Browser.msgBox("Please enter a valid Thrive Market product URL.");
    return;
  }

  const discount = Browser.inputBox(
    "Discount Details",
    "What is the discount? (e.g. 20% off, $5 off, BOGO):",
    Browser.Buttons.OK_CANCEL
  );
  if (discount === "cancel") return;

  const promoDates = Browser.inputBox(
    "Promotion Dates",
    "Enter the promotion date range (e.g. July 4–7, 2026):",
    Browser.Buttons.OK_CANCEL
  );
  if (promoDates === "cancel") return;

  Browser.msgBox("Fetching product info — this may take a moment. Click OK to continue.");

  const product = fetchProductData(productUrl);
  const doc     = buildDoc(product, productUrl, discount, promoDates);

  Browser.msgBox("Done! Your promo brief is ready:\n\n" + doc.getUrl());
}

// ─── Custom menu (when bound to a Sheet or Doc) ───────────────────────────────

function onOpen() {
  try {
    SpreadsheetApp.getUi()
      .createMenu("Thrive Promo")
      .addItem("Create Promotion Brief…", "createPromoBrief")
      .addToUi();
  } catch (_) {
    try {
      DocumentApp.getUi()
        .createMenu("Thrive Promo")
        .addItem("Create Promotion Brief…", "createPromoBrief")
        .addToUi();
    } catch (_) {}
  }
}

// ─── Product data fetching ────────────────────────────────────────────────────

function fetchProductData(url) {
  let html;
  try {
    const response = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: {
        "User-Agent":
          "Mozilla/5.0 (compatible; ThrivePromoBot/1.0; internal tooling)",
      },
    });
    if (response.getResponseCode() !== 200) {
      return fallbackProduct(url, "HTTP " + response.getResponseCode());
    }
    html = response.getContentText();
  } catch (e) {
    return fallbackProduct(url, e.message);
  }

  return {
    name:        extractMeta(html, "og:title")       || extractTitle(html)       || "Unknown Product",
    description: extractMeta(html, "og:description") || extractDescription(html) || "",
    imageUrl:    extractMeta(html, "og:image")       || "",
    brand:       extractBrand(html),
    price:       extractPrice(html),
    category:    extractCategory(html),
    badges:      extractBadges(html),
    ingredients: extractIngredients(html),
    fetchError:  null,
  };
}

function fallbackProduct(url, reason) {
  return {
    name: "Product (could not fetch page: " + reason + ")",
    description: "",
    imageUrl: "",
    brand: "",
    price: "",
    category: "",
    badges: [],
    ingredients: "",
    fetchError: reason,
  };
}

// ─── HTML extraction helpers ──────────────────────────────────────────────────

function extractMeta(html, property) {
  const patterns = [
    new RegExp('<meta[^>]+property=["\']' + property + '["\'][^>]+content=["\']([^"\']+)["\']', "i"),
    new RegExp('<meta[^>]+content=["\']([^"\']+)["\'][^>]+property=["\']' + property + '["\']', "i"),
  ];
  for (const re of patterns) {
    const m = html.match(re);
    if (m) return decodeHtmlEntities(m[1].trim());
  }
  return null;
}

function extractTitle(html) {
  const m = html.match(/<title[^>]*>([^<]+)<\/title>/i);
  return m ? decodeHtmlEntities(m[1].trim()) : null;
}

function extractDescription(html) {
  const m = html.match(/<meta[^>]+name=["\']description["\'][^>]+content=["\']([^"\']+)["\']>/i)
    || html.match(/<meta[^>]+content=["\']([^"\']+)["\'][^>]+name=["\']description["\']>/i);
  return m ? decodeHtmlEntities(m[1].trim()) : null;
}

function extractBrand(html) {
  const jsonLd = extractJsonLd(html);
  if (jsonLd && jsonLd.brand) {
    return typeof jsonLd.brand === "string" ? jsonLd.brand : jsonLd.brand.name || "";
  }
  const m = html.match(/"brand"\s*:\s*"([^"]+)"/i);
  return m ? decodeHtmlEntities(m[1]) : "";
}

function extractPrice(html) {
  const jsonLd = extractJsonLd(html);
  if (jsonLd && jsonLd.offers) {
    const offer = Array.isArray(jsonLd.offers) ? jsonLd.offers[0] : jsonLd.offers;
    if (offer && offer.price) return "$" + offer.price;
  }
  const m = html.match(/"price"\s*:\s*"?([\d.]+)"?/i);
  return m ? "$" + m[1] : "";
}

function extractCategory(html) {
  const jsonLd = extractJsonLd(html, "BreadcrumbList");
  if (jsonLd && jsonLd.itemListElement) {
    const crumbs = jsonLd.itemListElement;
    if (crumbs.length >= 2) return crumbs[crumbs.length - 2].name || "";
  }
  return "";
}

function extractBadges(html) {
  const knownBadges = [
    "Non-GMO", "Organic", "Gluten-Free", "Vegan", "Paleo",
    "Keto", "Kosher", "Fair Trade", "Whole30", "BPA-Free",
  ];
  return knownBadges.filter((b) =>
    new RegExp(b.replace("-", "[- ]?"), "i").test(html)
  );
}

function extractIngredients(html) {
  const m = html.match(/ingredients\s*[:\-]?\s*<[^>]*>([^<]{20,})/i)
    || html.match(/Ingredients[^:]*:\s*([A-Za-z,\s\(\)\.]{20,})/i);
  return m ? decodeHtmlEntities(m[1].trim()).substring(0, 500) : "";
}

function extractJsonLd(html, type) {
  const scriptPattern = /<script[^>]+type=["\']application\/ld\+json["\'][^>]*>([\s\S]*?)<\/script>/gi;
  let match;
  while ((match = scriptPattern.exec(html)) !== null) {
    try {
      const data = JSON.parse(match[1]);
      const records = Array.isArray(data) ? data : [data];
      for (const r of records) {
        if (!type || r["@type"] === type) return r;
      }
    } catch (_) {}
  }
  return null;
}

function decodeHtmlEntities(str) {
  return str
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'")
    .replace(/&nbsp;/g, " ");
}

// ─── Document builder ─────────────────────────────────────────────────────────

function buildDoc(product, productUrl, discount, promoDates) {
  const docTitle = "Promo Brief – " + product.name + " – " + promoDates;
  const doc  = DocumentApp.create(docTitle);
  const body = doc.getBody();

  const H1 = DocumentApp.ParagraphHeading.HEADING1;
  const H2 = DocumentApp.ParagraphHeading.HEADING2;
  const H3 = DocumentApp.ParagraphHeading.HEADING3;

  const titlePara = body.appendParagraph("Promotion Brief");
  titlePara.setHeading(H1);
  titlePara.editAsText().setForegroundColor("#1B4332");

  appendDivider(body);
  appendHeading(body, "Overview", H2);

  const overviewTable = body.appendTable();
  overviewTable.setBorderWidth(0);

  const rows = [
    ["Product",      product.name],
    ["Brand",        product.brand    || "—"],
    ["Category",     product.category || "—"],
    ["Member Price", product.price    || "—"],
    ["Discount",     discount         || "—"],
    ["Promo Dates",  promoDates       || "—"],
  ];

  rows.forEach(function(row) {
    const tableRow   = overviewTable.appendTableRow();
    const labelCell  = tableRow.appendTableCell(row[0]);
    labelCell.editAsText().setBold(true);
    tableRow.appendTableCell(row[1]);
  });

  styleTable(overviewTable);
  body.appendParagraph("");

  appendHeading(body, "Product Description", H2);
  body.appendParagraph(product.description || "[Add product description here]");
  body.appendParagraph("");

  appendHeading(body, "Certifications & Badges", H2);
  if (product.badges && product.badges.length > 0) {
    product.badges.forEach(function(b) {
      body.appendListItem(b).setGlyphType(DocumentApp.GlyphType.BULLET);
    });
  } else {
    body.appendParagraph("[List relevant certifications, e.g. Non-GMO, Organic]");
  }
  body.appendParagraph("");

  if (product.ingredients) {
    appendHeading(body, "Key Ingredients", H2);
    body.appendParagraph(product.ingredients);
    body.appendParagraph("");
  }

  appendDivider(body);
  appendHeading(body, "Promotion Messaging", H2);

  appendHeading(body, "Suggested Headline", H3);
  body.appendParagraph(buildHeadline(product.name, discount)).editAsText().setItalic(true);

  appendHeading(body, "Suggested Subhead", H3);
  body.appendParagraph(buildSubhead(product.name, product.brand, promoDates)).editAsText().setItalic(true);

  appendHeading(body, "Email Body Copy", H3);
  body.appendParagraph(buildEmailCopy(product, discount, promoDates, productUrl));
  body.appendParagraph("");

  appendHeading(body, "Short Social / Push Copy", H3);
  body.appendParagraph(buildSocialCopy(product.name, discount, promoDates)).editAsText().setItalic(true);
  body.appendParagraph("");

  appendDivider(body);
  appendHeading(body, "Campaign Checklist", H2);

  var checklist = [
    "Confirm discount code / pricing with Growth team",
    "Pull product image assets",
    "QA product page URL before send",
    "Schedule email send in ESP",
    "Schedule push notification",
    "Post to social channels",
    "Set up promo in Thrive backend / coupon system",
    "Monitor revenue and CTR during promo window",
    "Post-mortem debrief after promo ends",
  ];

  checklist.forEach(function(item) {
    body.appendListItem("☐  " + item).setGlyphType(DocumentApp.GlyphType.BULLET);
  });

  body.appendParagraph("");
  appendDivider(body);
  appendHeading(body, "Source", H2);
  body.appendParagraph("Product URL: " + productUrl);
  body.appendParagraph(
    "Brief generated: " + Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy 'at' h:mm a z"
    )
  );

  if (product.fetchError) {
    body.appendParagraph(
      "Note: product data could not be fully fetched (" + product.fetchError +
      "). Please fill in the fields manually."
    ).editAsText().setForegroundColor("#B91C1C");
  }

  doc.saveAndClose();
  return doc;
}

// ─── Copy generators ──────────────────────────────────────────────────────────

function buildHeadline(name, discount) {
  return discount ? "Save " + discount + " on " + name : "Members-Only Deal: " + name;
}

function buildSubhead(name, brand, dates) {
  var brandPart = brand ? "from " + brand + " " : "";
  return "Shop " + name + " " + brandPart +
    (dates ? "— available " + dates + " only." : "for a limited time.");
}

function buildEmailCopy(product, discount, dates, url) {
  var discountLine = discount
    ? "For a limited time, members save " + discount + " on " + product.name + "."
    : product.name + " is on promotion for members.";
  var badgeLine = product.badges && product.badges.length
    ? "Certified " + product.badges.join(", ") + "."
    : "";
  var descLine = product.description
    ? product.description.substring(0, 200) + (product.description.length > 200 ? "…" : "")
    : "";
  var datesLine = dates ? "Offer valid " + dates + "." : "";

  return [discountLine, descLine, badgeLine, datesLine, "Shop now: " + url]
    .filter(Boolean)
    .join("\n\n");
}

function buildSocialCopy(name, discount, dates) {
  var save = discount ? "Save " + discount + " " : "";
  var when = dates ? " " + dates + " only." : ".";
  return save + "on " + name + " — members-only deal" + when + " #ThriveMarket";
}

// ─── Doc style helpers ────────────────────────────────────────────────────────

function appendHeading(body, text, level) {
  var p = body.appendParagraph(text);
  p.setHeading(level);
  return p;
}

function appendDivider(body) {
  var p = body.appendParagraph("──────────────────────────────────────────────────────────");
  p.editAsText().setForegroundColor("#9CA3AF").setFontSize(8);
  return p;
}

function styleTable(table) {
  for (var r = 0; r < table.getNumRows(); r++) {
    var row = table.getRow(r);
    for (var c = 0; c < row.getNumCells(); c++) {
      var cell = row.getCell(c);
      cell.setPaddingTop(4).setPaddingBottom(4).setPaddingLeft(8).setPaddingRight(8);
      if (r % 2 === 0) cell.setBackgroundColor("#F0FDF4");
    }
  }
}

// ─── Web app HTML form ────────────────────────────────────────────────────────

function getFormHtml() {
  return '<!DOCTYPE html><html><head><meta charset="utf-8">' +
    '<title>Thrive Market Promo Brief Generator</title>' +
    '<style>' +
    'body{font-family:Google Sans,sans-serif;max-width:540px;margin:48px auto;padding:0 16px;color:#1f2937}' +
    'h1{color:#1B4332;font-size:22px;margin-bottom:24px}' +
    'label{display:block;font-size:13px;font-weight:600;margin:16px 0 4px}' +
    'input{width:100%;box-sizing:border-box;padding:8px 10px;border:1px solid #d1d5db;border-radius:6px;font-size:14px}' +
    'button{margin-top:24px;background:#1B4332;color:#fff;border:none;padding:10px 24px;border-radius:6px;font-size:15px;cursor:pointer}' +
    'button:disabled{opacity:.5;cursor:default}' +
    '#status{margin-top:16px;font-size:14px;color:#374151}' +
    '#result{margin-top:12px;font-size:15px}' +
    '#result a{color:#1B4332;font-weight:600}' +
    '</style></head><body>' +
    '<h1>Thrive Market Promo Brief Generator</h1>' +
    '<label>Product URL</label>' +
    '<input id="url" type="url" placeholder="https://thrivemarket.com/p/..." />' +
    '<label>Discount</label>' +
    '<input id="discount" type="text" placeholder="e.g. 20% off, $5 off, BOGO" />' +
    '<label>Promotion Dates</label>' +
    '<input id="dates" type="text" placeholder="e.g. July 4–7, 2026" />' +
    '<br><button id="btn" onclick="submit()">Generate Brief</button>' +
    '<p id="status"></p><p id="result"></p>' +
    '<script>' +
    'function submit(){' +
    '  var url=document.getElementById("url").value.trim();' +
    '  var discount=document.getElementById("discount").value.trim();' +
    '  var dates=document.getElementById("dates").value.trim();' +
    '  if(!url){alert("Please enter a product URL.");return;}' +
    '  document.getElementById("btn").disabled=true;' +
    '  document.getElementById("status").textContent="Fetching product info and building your brief — this may take 20-30 seconds…";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(r){' +
    '      document.getElementById("status").textContent="";' +
    '      document.getElementById("btn").disabled=false;' +
    '      if(r.error){document.getElementById("result").textContent="Error: "+r.error;return;}' +
    '      document.getElementById("result").innerHTML="Done! <a href=\'"+r.url+"\' target=\'_blank\'>Open your promo brief</a>";' +
    '    })' +
    '    .withFailureHandler(function(e){' +
    '      document.getElementById("status").textContent="Error: "+e.message;' +
    '      document.getElementById("btn").disabled=false;' +
    '    })' +
    '    .processForm({productUrl:url,discount:discount,promoDates:dates});' +
    '}' +
    '<\/script></body></html>';
}
