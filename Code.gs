/**
 * Thrive Market Product Promotion Brief Generator
 *
 * Usage: Run createPromoBrief() from the Apps Script editor,
 * or call it from a custom menu after opening the spreadsheet.
 */

// ─── Entry point ─────────────────────────────────────────────────────────────

function createPromoBrief() {
  const ui = SpreadsheetApp.getUi
    ? SpreadsheetApp.getUi()
    : DocumentApp.getUi();

  const urlResponse = ui.prompt(
    "Thrive Market Promo Brief Generator",
    "Paste the Thrive Market product URL:",
    ui.ButtonSet.OK_CANCEL
  );

  if (urlResponse.getSelectedButton() !== ui.Button.OK) return;
  const productUrl = urlResponse.getResponseText().trim();

  if (!productUrl.includes("thrivemarket.com")) {
    ui.alert("Please enter a valid Thrive Market product URL.");
    return;
  }

  const discountResponse = ui.prompt(
    "Discount Details",
    "What is the discount? (e.g. 20% off, $5 off, BOGO):",
    ui.ButtonSet.OK_CANCEL
  );
  if (discountResponse.getSelectedButton() !== ui.Button.OK) return;
  const discount = discountResponse.getResponseText().trim();

  const dateResponse = ui.prompt(
    "Promotion Dates",
    "Enter the promotion date range (e.g. July 4–7, 2026):",
    ui.ButtonSet.OK_CANCEL
  );
  if (dateResponse.getSelectedButton() !== ui.Button.OK) return;
  const promoDates = dateResponse.getResponseText().trim();

  ui.alert("Fetching product info and building your brief — this may take a moment.");

  const product = fetchProductData(productUrl);
  const doc = buildDoc(product, productUrl, discount, promoDates);

  ui.alert(
    "Done!",
    "Your promo brief is ready:\n" + doc.getUrl(),
    ui.ButtonSet.OK
  );
}

// ─── Product data fetching ────────────────────────────────────────────────────

/**
 * Fetches and parses product data from a Thrive Market product page.
 * Returns a plain object with the fields we care about.
 */
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
    name:         extractMeta(html, "og:title")        || extractTitle(html)    || "Unknown Product",
    description:  extractMeta(html, "og:description")  || extractDescription(html) || "",
    imageUrl:     extractMeta(html, "og:image")        || "",
    brand:        extractBrand(html),
    price:        extractPrice(html),
    category:     extractCategory(html),
    badges:       extractBadges(html),
    ingredients:  extractIngredients(html),
    fetchError:   null,
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
  // Matches both og: and name= meta tags
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
  // Thrive product pages often embed brand in JSON-LD
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
  // Breadcrumb or category meta
  const jsonLd = extractJsonLd(html, "BreadcrumbList");
  if (jsonLd && jsonLd.itemListElement) {
    const crumbs = jsonLd.itemListElement;
    if (crumbs.length >= 2) {
      return crumbs[crumbs.length - 2].name || "";
    }
  }
  return "";
}

function extractBadges(html) {
  // Look for common Thrive badges: Non-GMO, Organic, Gluten-Free, etc.
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
  const docTitle =
    "Promo Brief – " + product.name + " – " + promoDates;
  const doc = DocumentApp.create(docTitle);
  const body = doc.getBody();

  // Header style helpers
  const H1 = DocumentApp.ParagraphHeading.HEADING1;
  const H2 = DocumentApp.ParagraphHeading.HEADING2;
  const H3 = DocumentApp.ParagraphHeading.HEADING3;

  // ── Title ──────────────────────────────────────────────────────────────────
  const titlePara = body.appendParagraph("🛒 Promotion Brief");
  titlePara.setHeading(H1);
  titlePara.editAsText().setForegroundColor("#1B4332");

  appendDivider(body);

  // ── Overview table ─────────────────────────────────────────────────────────
  appendHeading(body, "Overview", H2);

  const overviewTable = body.appendTable();
  overviewTable.setBorderWidth(0);

  const rows = [
    ["Product",        product.name],
    ["Brand",         product.brand || "—"],
    ["Category",      product.category || "—"],
    ["Member Price",  product.price || "—"],
    ["Discount",      discount || "—"],
    ["Promo Dates",   promoDates || "—"],
  ];

  rows.forEach(([label, value]) => {
    const row = overviewTable.appendTableRow();
    const labelCell = row.appendTableCell(label);
    labelCell.editAsText().setBold(true);
    row.appendTableCell(value);
  });

  styleTable(overviewTable);
  body.appendParagraph("");

  // ── Product description ────────────────────────────────────────────────────
  appendHeading(body, "Product Description", H2);
  body.appendParagraph(
    product.description || "[Add product description here]"
  );
  body.appendParagraph("");

  // ── Certifications / badges ────────────────────────────────────────────────
  appendHeading(body, "Certifications & Badges", H2);
  if (product.badges && product.badges.length > 0) {
    product.badges.forEach((b) => {
      body.appendListItem(b).setGlyphType(
        DocumentApp.GlyphType.BULLET
      );
    });
  } else {
    body.appendParagraph("[List relevant certifications, e.g. Non-GMO, Organic]");
  }
  body.appendParagraph("");

  // ── Key ingredients ────────────────────────────────────────────────────────
  if (product.ingredients) {
    appendHeading(body, "Key Ingredients", H2);
    body.appendParagraph(product.ingredients);
    body.appendParagraph("");
  }

  // ── Promotion messaging ────────────────────────────────────────────────────
  appendDivider(body);
  appendHeading(body, "Promotion Messaging", H2);

  appendHeading(body, "Suggested Headline", H3);
  body.appendParagraph(
    buildHeadline(product.name, discount)
  ).editAsText().setItalic(true);

  appendHeading(body, "Suggested Subhead", H3);
  body.appendParagraph(
    buildSubhead(product.name, product.brand, promoDates)
  ).editAsText().setItalic(true);

  appendHeading(body, "Email Body Copy", H3);
  body.appendParagraph(
    buildEmailCopy(product, discount, promoDates, productUrl)
  );
  body.appendParagraph("");

  appendHeading(body, "Short Social/Push Copy", H3);
  body.appendParagraph(
    buildSocialCopy(product.name, discount, promoDates)
  ).editAsText().setItalic(true);
  body.appendParagraph("");

  // ── Campaign checklist ─────────────────────────────────────────────────────
  appendDivider(body);
  appendHeading(body, "Campaign Checklist", H2);

  const checklist = [
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

  checklist.forEach((item) => {
    body.appendListItem("☐  " + item).setGlyphType(
      DocumentApp.GlyphType.BULLET
    );
  });

  body.appendParagraph("");

  // ── Source / metadata ──────────────────────────────────────────────────────
  appendDivider(body);
  appendHeading(body, "Source", H2);
  body.appendParagraph("Product URL: " + productUrl);
  body.appendParagraph(
    "Brief generated: " + Utilities.formatDate(
      new Date(), Session.getScriptTimeZone(), "MMMM d, yyyy 'at' h:mm a z"
    )
  );

  if (product.fetchError) {
    body
      .appendParagraph(
        "⚠️  Note: product data could not be fully fetched (" +
          product.fetchError +
          "). Please fill in the fields manually."
      )
      .editAsText()
      .setForegroundColor("#B91C1C");
  }

  doc.saveAndClose();
  return doc;
}

// ─── Copy generators ──────────────────────────────────────────────────────────

function buildHeadline(name, discount) {
  if (discount) return "Save " + discount + " on " + name;
  return "Members-Only Deal: " + name;
}

function buildSubhead(name, brand, dates) {
  const brandPart = brand ? "from " + brand + " " : "";
  return (
    "Shop " + name + " " + brandPart +
    (dates ? "— available " + dates + " only." : "for a limited time.")
  );
}

function buildEmailCopy(product, discount, dates, url) {
  const discountLine = discount
    ? "For a limited time, members save " + discount + " on " + product.name + "."
    : product.name + " is on promotion for members.";
  const badgeLine =
    product.badges && product.badges.length
      ? "Certified " + product.badges.join(", ") + "."
      : "";
  const descLine = product.description
    ? product.description.substring(0, 200) + (product.description.length > 200 ? "…" : "")
    : "";
  const datesLine = dates ? "Offer valid " + dates + "." : "";

  return [discountLine, descLine, badgeLine, datesLine, "Shop now → " + url]
    .filter(Boolean)
    .join("\n\n");
}

function buildSocialCopy(name, discount, dates) {
  const save = discount ? "Save " + discount + " " : "";
  const when = dates ? " 🗓️ " + dates + " only." : ".";
  return save + "on " + name + " — members-only deal" + when + " 🌿 #ThriveMarket";
}

// ─── Doc style helpers ────────────────────────────────────────────────────────

function appendHeading(body, text, level) {
  const p = body.appendParagraph(text);
  p.setHeading(level);
  return p;
}

function appendDivider(body) {
  const p = body.appendParagraph("─".repeat(60));
  p.editAsText().setForegroundColor("#9CA3AF").setFontSize(8);
  return p;
}

function styleTable(table) {
  for (let r = 0; r < table.getNumRows(); r++) {
    const row = table.getRow(r);
    for (let c = 0; c < row.getNumCells(); c++) {
      const cell = row.getCell(c);
      cell.setPaddingTop(4).setPaddingBottom(4)
          .setPaddingLeft(8).setPaddingRight(8);
      if (r % 2 === 0) {
        cell.setBackgroundColor("#F0FDF4");
      }
    }
  }
}

// ─── Custom menu (for Sheets or Docs trigger) ─────────────────────────────────

function onOpen() {
  const ui = SpreadsheetApp.getUi
    ? SpreadsheetApp.getUi()
    : DocumentApp.getUi();
  ui.createMenu("Thrive Promo")
    .addItem("Create Promotion Brief…", "createPromoBrief")
    .addToUi();
}
