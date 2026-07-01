# WhatsApp Delivery Monitor — engageSPARK

A static, client-side dashboard that reads a weekly engageSPARK campaign
export (`.xlsx`) and shows delivery health for the WhatsApp messaging
campaign: enrollment, messages sent, response rate, delivery status
breakdown, Meta blocks, and Saturday prompt-button engagement.

No backend, no build step. Everything (parsing the Excel file, computing
metrics, drawing charts) runs in the visitor's browser via
[SheetJS](https://sheetjs.com/) and [Chart.js](https://www.chartjs.org/),
both loaded from a public CDN. Nothing is uploaded anywhere.

## 1. Publish it on GitHub Pages

1. Push this whole folder to a GitHub repository.
2. Repo → **Settings → Pages** → Source: `main` branch, `/ (root)`.
3. Your dashboard will be live at `https://<user>.github.io/<repo>/`.

## 2. Updating the report every week

The dashboard always reads **`data/latest.xlsx`**. To refresh it:

1. Download the new report from engageSPARK for campaign `747707`.
2. Rename the downloaded file to exactly `latest.xlsx`.
3. Replace `data/latest.xlsx` in the repo with it (same path, same name).
4. Commit and push. Nothing else needs to change — the dashboard
   recalculates everything from the new file automatically.

If you ever open `index.html` directly as a local file (not through a web
server), the browser will block the automatic fetch of `data/latest.xlsx`
for security reasons (CORS). In that case the dashboard shows an
**"Upload report manually"** button as a fallback — pick the file there to
preview it locally before pushing. Once served through GitHub Pages (or
any local dev server such as `python -m http.server`), the automatic load
works normally.

## 3. What the dashboard expects in the workbook

Two sheets from the standard engageSPARK export are used:

- **`Campaign Report`** — one row per enrolled participant (`Contact ID`,
  `First Name`, `Subscription Time (America/Chicago)`,
  `Subscription Status`).
- **`WhatsApp Log`** — one row per individual WhatsApp message
  (`Contact ID`, `Direction`, `Message Label`,
  `Time of Message (America/Chicago)`, `Delivery Status`,
  `Delivery Error`, `Message`).

If engageSPARK ever renames these sheets or columns, update the constants
at the top of `js/app.js` (`CONFIG.SHEET_CAMPAIGN`, `CONFIG.SHEET_WHATSAPP`)
or the field names used in `buildCampaignData()` / `buildWhatsAppData()`.

## 4. How the metrics are computed

- **Participants enrolled** — distinct `Contact ID` in `Campaign Report`.
- **Messages sent to date** — rows in `WhatsApp Log` with
  `Direction = outbound`.
- **Participants who replied ≥1 time** — distinct `Contact ID` with at
  least one `Direction = inbound` row, ever (not filtered by week/round).
- **Delivery status chart** — outbound messages only, grouped by
  `Delivery Status` (`Read`, `Delivered`, `Not Delivered`, `Unknown`).
  "Delivered" here means delivered-but-not-read.
- **Blocked by Meta** — outbound messages with
  `Delivery Error = "Blocked by Meta"`, bucketed by campaign week.
- **Saturday button clicks** — inbound messages whose text matches a
  known WhatsApp quick-reply phrase (see `CONFIG.BUTTON_REPLY_TEXTS` in
  `js/app.js`). **If the button wording changes**, add the new variant to
  that list so the click chart keeps counting correctly.
- **Campaign week** — messages are bucketed into 7-day windows starting
  from the very first message timestamp in the log (`Week 1` = first 7
  days, `Week 2` = the next 7, etc.). This applies uniformly to every
  chart with a week axis or a week filter.
- **Enrollment rounds** — detected automatically by clustering distinct
  enrollment dates: any gap larger than `CONFIG.ROUND_GAP_DAYS` (default
  3 days) between consecutive enrollment dates starts a new round. As new
  people enroll in future weeks, new rounds will appear on their own —
  no manual configuration needed.

## 5. Filters

- **Enrollment round** — "All" shows the overall total. Selecting one or
  more specific rounds restricts every chart to those rounds; selecting
  more than one switches the delivery-status chart into side-by-side
  comparison mode (one set of bars per round).
- **Specific participants** — optional, narrows every chart down to the
  selected people only (search by name or contact ID).
- **Weeks** — restricts which campaign weeks are included. For the
  delivery-status chart this changes the aggregate totals; for the
  weekly-axis charts (Blocked by Meta, Button clicks) it simply hides the
  non-selected week bars.

All three filters combine (AND logic) and apply to the three charts —
the scorecards at the top always show unfiltered, campaign-to-date totals.

## 6. File structure

```
index.html          Page markup
css/styles.css       Styling
js/app.js            Parsing, metrics, filters, charts (all logic lives here)
data/latest.xlsx     The current weekly report — overwrite this each week
```
