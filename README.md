# Lead Reporting Dashboard on Google Sheet - Google Apps Script

Automated multi-source lead consolidation and reporting dashboard built with Google Apps Script and Google Sheets.

---

## Project overview

This system consolidates leads from **multiple Google Sheets sources** (connected via Make/Integromat, Zapier, or Elementor webhook) into a single, unified dashboard, with deduplication, key performance indicator tracking, and dynamic filtering.

**Technologies used:** Google Sheets · Google Apps Script

---

## Features

| Feature | Description |
|---|---|
| **Multi-source import** | Fetches leads from up to N Google Sheets sources automatically |
| **Smart deduplication** | Priority order: Email → Phone |
| **Universal date parsing** | Handles FR, EN, ISO, timestamp, serial formats |
| **Date format detection** | Auto-detects DD/MM/YYYY vs MM/DD/YYYY per source file |
| **Dynamic dashboard** | KPIs + cross-table by period (Daily / Weekly / Monthly / Yearly) |
| **Custom date filter** | Interactive dual-calendar with quick-select shortcuts |
| **Multi-sheet support** | Each source can point to a specific tab via `gid` in URL |
| **Source merging** | Same source name in Config = merged into one column |
| **Nightly auto-sync** | Configurable trigger for automatic overnight consolidation |

---

## Spreadsheet Structure

### Tabs

| Tab | Description |
|---|---|
| **Dashboard** | Main view — KPIs + dynamic leads table |
| **Data** | Unique leads after deduplication |
| **Raw** | All imported leads before deduplication |
| **Duplicates** | Leads removed during deduplication |
| **Config** | List of sources (Name + Google Sheets URL) |
| **LeadCache** | *(hidden)* Internal cache for fast dashboard rendering |

### Config Tab Format

| Column A | Column B |
|---|---|
| Source Name | Full Google Sheets URL (with `#gid=` for specific tab) |


> **Same name = merged column.** Two rows with the same name will be combined into one dashboard column.

---

## Setup Instructions

### 1. Create the Google Sheets file

Create a new Google Sheets file named **"Reporting Leads"** and add these tabs manually:
- `Dashboard`
- `Data`
- `Raw`
- `Duplicates`
- `Config`

### 2. Add your sources in Config

Fill in the **Config** tab starting from row 2:
- **Column A** → Source display name
- **Column B** → Full URL of the Google Sheet source

To get the URL of a specific tab:
1. Open the source Google Sheet
2. Click on the target tab
3. Copy the full URL from the browser (it includes `#gid=XXXXXXXXX`)

### 3. Install the script

1. In your Reporting Leads file: **Extensions → Apps Script**
2. Delete all existing code in `Code.gs`
3. Paste the full script content
4. Click **Save** (💾)

### 4. Authorize permissions

1. Run `consolidateLeads` for the first time
2. Click **Review permissions** when prompted
3. Grant access to Google Sheets (needed to read source files)

### 5. Run Full Consolidation

From the **"Lead Reporting"** menu:
> **Lead Reporting → Full Consolidation**

---

## Menu Options

| Menu Item | Action |
|---|---|
| **Full Consolidation** | Re-fetches all sources, rebuilds dashboard from scratch |
| **Refresh Dashboard** | Re-fetches all sources, updates data without rebuilding layout |
| **Date Filter** | Opens interactive calendar for custom date range filtering |
| **Enable nightly auto-sync** | Sets a daily trigger at midnight for automatic consolidation |

---

## Dashboard Periods

| Period | Display | Limit |
|---|---|---|
| **Daily** | `25 Feb 2026` | Last 90 days |
| **Weekly** | `February 16 - 22, 2026` | Last 6 months |
| **Monthly** | `February 2026` | Last 24 months |
| **Yearly** | `2026` | All years |

---

## KPI Cards

| KPI | Description |
|---|---|
| **Total Leads** | All unique leads |
| **Today** | Leads with today's date |
| **This Week** | Leads in the current ISO week (Mon → Sun) |
| **This Month** | Leads in the current calendar month |
| **Year XXXX** | Leads in the current year (updates automatically) |
| **Duplicates** | Total leads removed by deduplication |

---

### Critical date format setting

In Forms, the Date field must use **`Y-m-d`** format to output ISO dates:

```
Y-m-d  →  2026-02-25  ✅ (recommended, no ambiguity)
d/m/Y  →  25/02/2026  ⚠️ (works but may cause edge cases)
```

> **Why it matters:** Google Sheets may auto-interpret `9/2/2026` as September 2 (US format) instead of February 9 (FR format). ISO format `2026-02-09` is never ambiguous.

---

## Technical Notes

### Deduplication logic

```
1. If Email exists     → deduplicate by Email (normalized lowercase)
2. Else if Phone exists → deduplicate by Phone (digits only)
3. Else if Name exists  → deduplicate by Name (lowercase, no spaces)

Chronological sort applied BEFORE dedup → oldest lead is kept
```

### Date parsing priority

```
1. Native Date object (Google Sheets auto-conversion)
2. ISO with timezone  : 2026-02-09T23:00:00.000Z
3. ISO date only      : 2026-02-09
4. FR separator       : 25/02/2026 or 25-02-2026
5. FR text            : 25 février 2026
6. EN text            : February 25, 2026 or 25 Feb 2026
7. Google Sheets serial number
```

### Multi-tab source URLs

To target a specific tab, include the `gid` parameter in the URL:

```
https://docs.google.com/spreadsheets/d/{ID}/edit#gid={GID}
```

The `GID` is visible in the browser URL when you click on the tab.

### Cache system

After each consolidation, leads are stored in the hidden **LeadCache** tab as `YYYY-MM-DD` + source name pairs. This allows instant period switching (`onEdit` trigger) without re-fetching all sources.

---

## File Structure

```
Code.gs
│
├── consolidateLeads()       — Full import + rebuild
├── refreshDashboard()       — Re-fetch + update only
├── fetchAndProcess()        — Core import engine
│   ├── detectDateFormat()   — Auto-detect DD/MM vs MM/DD
│   └── toDateObj()          — Universal date parser
│
├── buildDashboard()         — Full dashboard construction
├── drawDynamicTable()       — Cross-table rendering
├── groupBySource()          — Data grouping by period
├── filterCustom()           — Custom date range filtering
│
├── storeCache()             — Write LeadCache tab
├── loadCache()              — Read LeadCache tab
│
├── openDateFilter()         — HTML calendar dialog
├── applyDateFilter()        — Apply filter from dialog
│
├── onEdit()                 — Period dropdown trigger
├── onOpen()                 — Menu creation
└── setNightlyTrigger()      — Auto-sync setup
```

---

## Troubleshooting

| Problem | Likely cause | Fix |
|---|---|---|
| "No data for this selection" | Cache empty or corrupted | Run Full Consolidation |
| Lead not imported | Invalid or future date | Check source sheet date format |
| Wrong lead count | Deduplication working correctly | Check Duplicates tab |
| New source not showing | Cache not refreshed | Run Full Consolidation |
| Specific tab not loading | URL missing `#gid=` | Copy URL with tab active in browser |
| Date parsing error | Ambiguous format (e.g. `9/2/2026`) | Set Forms to `Y-m-d` format |

---

## License

Free to use and modify.
