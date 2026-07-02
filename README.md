# Thrive Market Promo Brief Generator

An Apps Script that generates a formatted Google Doc promotion brief for a Thrive Market product, pulling live product data from the product page.

## Setup

1. Go to [script.google.com](https://script.google.com) and create a new project.
2. Paste the contents of `Code.gs` into the editor (replacing the default `myFunction`).
3. Save the project (name it anything, e.g. *Thrive Promo Brief*).
4. Run `createPromoBrief` once — Google will ask you to authorize access to Google Docs and external URLs.

You can also bind this script to a Google Sheet or Doc via **Extensions → Apps Script**, in which case a **Thrive Promo** menu will appear automatically on open.

## How it works

1. Prompts for a Thrive Market product URL, the discount, and the promo date range.
2. Fetches the product page and extracts:
   - Product name, brand, category
   - Member price
   - Description
   - Certifications/badges (Organic, Non-GMO, Vegan, etc.)
   - Key ingredients (if present on the page)
3. Creates a new Google Doc with:
   - **Overview table** — all key deal details at a glance
   - **Product description & certifications**
   - **Promotion messaging** — suggested headline, subhead, email body copy, and short social/push copy
   - **Campaign checklist** — standard steps from setup to post-mortem
   - **Source & metadata**

## Notes

- The script uses `UrlFetchApp` to fetch the product page. If Thrive Market's CDN blocks the request, the doc is still created with placeholder text and a warning.
- No external APIs or API keys required — everything runs inside Apps Script.
- The generated doc is saved to your Google Drive root.
