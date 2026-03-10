# Leads Reporting Dashboard — Google Apps Script

A multi-source lead aggregation and reporting system built with Google Apps Script and Google Sheets. It consolidates leads from multiple Google Forms/Sheets into a single dashboard with automatic deduplication, business type classification, GA4 website traffic integration, and period-based analytics.

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Setup](#setup)
- [Configuration](#configuration)
- [How It Works](#how-it-works)
- [Dashboard](#dashboard)
- [Menu Reference](#menu-reference)
- [Field Mapping](#field-mapping)
- [Business Type Classification](#business-type-classification)
- [GA4 Integration](#ga4-integration)
- [Deduplication Logic](#deduplication-logic)
- [Caching System](#caching-system)
- [Triggers](#triggers)
- [Sheet Structure](#sheet-structure)

---

## Overview

This project automates lead reporting for businesses collecting leads from multiple sources (websites, landing pages, forms). Instead of checking each source manually, the dashboard pulls all data into one place, removes duplicates, classifies leads by business type, and displays analytics broken down by period and source.

It is designed to run entirely inside Google Sheets with no external backend or paid services.

---

## Features

- Aggregates leads from up to 8+ Google Sheets sources
- Deduplicates by email, phone number, and name
- Keeps the most recent submission when duplicates are found
- Classifies leads by business type with configurable merge rules
- Displays 5 KPI counters: total, today, this week, this month, this year
- Displays 3 performance KPIs: traffic-to-lead rate, month-over-month growth, top source this month
- Period selector: Daily, Weekly, Monthly, Yearly views
- Tag-based source filtering
- Custom date range filter with calendar picker
- GA4 website sessions and users table per source
- Auto-sync every 8 hours via time-based trigger
- Automatic re-authorization prompt on open

---

## Setup

**Step 1 — Copy the script**

Open your Google Sheet, go to Extensions > Apps Script, paste the contents of `Code.gs`, and save.

**Step 2 — Add the manifest scopes**

In Apps Script, go to Project Settings, enable "Show appsscript.json in editor", then paste the following into `appsscript.json`:

```json
{
  "timeZone": "Europe/Paris",
  "dependencies": {},
  "exceptionLogging": "STACKDRIVER",
  "runtimeVersion": "V8",
  "oauthScopes": [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/script.container.ui",
    "https://www.googleapis.com/auth/analytics.readonly",
    "https://www.googleapis.com/auth/script.external_request",
    "https://www.googleapis.com/auth/script.scriptapp"
  ]
}
```

**Step 3 — Authorize the script**

Reload the spreadsheet. A dialog will appear asking for authorization. Click the link and accept all permissions. If the dialog does not appear, use Lead Reporting > Authorize / Re-authorize script from the menu.

**Step 4 — Set up triggers**

Go to Lead Reporting > Setup triggers. This creates:
- A time-based trigger that runs a full consolidation every 8 hours
- An installable onEdit trigger that refreshes the dashboard (including GA4) when the period or tag selector is changed

**Step 5 — Run the first consolidation**

Go to Lead Reporting > Full Consolidation. This imports all leads, deduplicates them, and builds the dashboard.

---

## Configuration

The `Config` sheet drives everything. Each row defines one data source.

| Column | Description |
|--------|-------------|
| A | Source name (displayed in the dashboard) |
| B | Full URL of the source Google Sheet (including `gid=` parameter if needed) |
| C | Tag (optional, used to group sources in the tag filter) |
| D | GA4 Property ID (optional, numbers only, e.g. `526737266`) |

Row 1 is a header row. Data starts from row 2.

The script reads columns A through D. Column F is used internally to display the last sync date.

---

## How It Works

When a consolidation or refresh is triggered, the script:

1. Reads all source URLs from the Config sheet
2. Opens each source spreadsheet and reads all rows
3. Detects the column layout of each source using the field map (see Field Mapping)
4. Normalizes each row into a standard 13-field format
5. Parses and validates dates, filtering out invalid or future entries
6. Sorts all leads by date descending to keep the most recent version of each person
7. Runs deduplication using email, phone, and name as identifiers
8. Writes unique leads to the Data sheet, all leads to Raw, duplicates to Duplicates
9. Stores a lightweight cache in a hidden sheet called LeadCache
10. Builds or refreshes the Dashboard sheet

---

## Dashboard

The dashboard is organized in vertical sections.

**Rows 2-4** — Title bar with last updated date and a separator line.

**Rows 6-7** — Five KPI cards: Total Leads, Today, This Week, This Month, Year to Date. Each card has a tooltip note explaining the metric.

**Rows 9-10** — Three performance KPIs: Traffic to Lead Rate (requires GA4), Lead Growth month-over-month, and Top Source This Month.

**Row 13** — Period selector (Daily / Weekly / Monthly / Yearly) and Tag selector. Changing either value triggers an automatic redraw of all tables below.

**Rows 16+** — Three data tables stacked vertically:
- Leads by Source: period rows, source columns, totals
- Leads by Business Type: period rows, business type columns, totals
- Website Visits GA4: sessions and unique users per source per period (only shown if GA4 property IDs are configured)

All tables respect the current period and tag selection. A custom date filter (accessible from the menu) overrides the period selector and shows daily granularity for the selected range.

---

## Menu Reference

The Lead Reporting menu is split into two sections.

**Daily use**

- Refresh Dashboard — re-imports all leads from all sources and rebuilds the dashboard
- Date Filter — opens a calendar picker to apply a custom date range to all tables

**Technical / admin**

- Full Consolidation — same as Refresh but also rebuilds the entire dashboard layout from scratch
- Setup triggers — installs the 8-hour auto-sync and the installable onEdit trigger
- Test GA4 Access — runs a diagnostic on all configured GA4 properties and shows results in a popup
- Authorize / Re-authorize script — forces the OAuth authorization flow

---

## Field Mapping

The script maps source column headers to a standard set of fields using a case-insensitive, accent-tolerant lookup. Smart apostrophes and typographic quotes in header names are automatically normalized before comparison.

| Internal Field | Recognized Headers (partial list) |
|----------------|-----------------------------------|
| date | date, created at, date de creation |
| nom | nom, name, last name |
| prenom | first name, prenom |
| tel | telephone, phone, phone/whatsapp, numéro de téléphone (gsm) |
| email | email, e-mail, adresse e-mail |
| entreprise | entreprise, company name, nom de l'entreprise |
| ville | ville, city |
| pays | country, pays |
| adresse | adresse, address, address line 1 |
| type | business type, type d'activite, other business type |
| message | additional details, message, request |
| quantity | production_volume, estimated order quantity |
| prodtype | product_type, toys category |

To add a new alias for any field, add it to the corresponding array in the `FIELD_MAP` variable at the top of `Code.gs`.

The standard output format for all sheets is: Date, Source, Type, Full Name, Phone, Email, Company, City, Country, Address, Message, Quantity, Product Type.

---

## Business Type Classification

Raw business type values from source forms are normalized into canonical categories using the `canonicalBT` function.

**Priority order (displayed left to right in the table)**

1. Wholesaler
2. Vape Shop
3. Tobacco & Smoke Shop
4. Retail Store
5. Convenience Store
6. E-Commerce Retail
7. Consumer
8. Unknown

Each category absorbs its aliases via substring or exact match (case-insensitive). For example, "online retailer", "online vape store", and "ecommerce" all map to E-Commerce Retail. "Retail shop" and "retailer" map to Retail Store.

Unrecognized values are kept as-is and displayed alphabetically between Consumer and Unknown. Unknown is always last.

The priority order matters because it prevents incorrect absorption: E-Commerce Retail is evaluated before Retail Store, so "online retailer" is not incorrectly captured by the broader "retailer" alias.

To add or modify categories, edit the `BT_MERGE` object and the `BT_PRIORITY` array in `Code.gs`.

---

## GA4 Integration

The GA4 table fetches sessions and unique users from the Google Analytics Data API v1beta for each source that has a property ID configured in column D of Config.

**Requirements**

- The Google Analytics Data API must be added in Apps Script under Services
- The script must be authorized with the `analytics.readonly` scope
- The GA4 property must be accessible by the Google account running the script

**Dimension granularity**

The GA4 query dimension automatically adapts to the selected period:

| Period | GA4 Dimension |
|--------|---------------|
| Daily | date |
| Weekly | isoYearIsoWeek |
| Monthly | yearMonth |
| Yearly | year |
| Custom date filter | date (always daily) |

**Trigger requirement**

GA4 calls require OAuth tokens (`ScriptApp.getOAuthToken()`), which are not available in simple `onEdit` triggers. The installable `onDashboardEdit` trigger runs as the script owner and has full OAuth access. The simple `onEdit` fallback redraws only the Leads and Business Type tables without calling GA4.

To diagnose GA4 issues, use Lead Reporting > Test GA4 Access.

---

## Deduplication Logic

Deduplication runs after all sources are merged.

1. All leads are sorted by date descending so the most recent submission is processed first
2. The script iterates through the sorted list, tracking seen emails, phone numbers, and names
3. A lead is marked as a duplicate if its normalized email, normalized phone, or (when both are absent) its name matches a previously seen value
4. Unique leads are written to the Data sheet. Duplicates go to the Duplicates sheet.

Phone normalization strips spaces, dashes, dots, parentheses, and the `+` prefix before comparison. Email normalization lowercases and strips whitespace. The result is that if the same person submits a form twice, only their most recent submission is counted.

---

## Caching System

After each consolidation, the script writes a lightweight cache to a hidden sheet called `LeadCache`. This avoids re-fetching all source sheets when only the dashboard display needs to change (for example when switching periods or applying a date filter).

The cache stores:
- Row 1: JSON metadata (source names, total counts, last updated timestamp, GA4 property map)
- Rows 2+: One row per unique lead with date, source name, and business type

When the period or tag selector is changed, the `onDashboardEdit` trigger reads from the cache instead of re-fetching all sources. The `applyDateFilter` function also reads from the cache.

The cache is rebuilt on every Full Consolidation and every Refresh Dashboard.

---

## Triggers

**Time-based trigger**

Created by `setNightlyTrigger`. Calls `consolidateLeads` every 8 hours. Runs as the script owner with full OAuth. Keeps the dashboard current without manual intervention.

**Installable onEdit trigger**

Also created by `setNightlyTrigger`. Calls `onDashboardEdit` when any cell in the Dashboard sheet is edited. Because it is installable, it runs with full OAuth and can call the GA4 API. It only reacts to edits on the period selector (row 13, column 3) and the tag selector (row 13, column 5).

**Simple onEdit fallback**

The built-in `onEdit` function responds to the same cells without OAuth. It redraws only the Leads by Source and Leads by Business Type tables, skipping GA4. This acts as a fallback if the installable trigger has not been set up.

To reset all triggers, run Lead Reporting > Setup triggers again. It deletes all existing time-based and onEdit installable triggers before creating new ones.

---

## Sheet Structure

| Sheet | Description |
|-------|-------------|
| Config | Source configuration: name, URL, tag, GA4 property ID |
| Dashboard | Main reporting view, auto-generated |
| Data | Unique (deduplicated) leads, sorted most recent first |
| Raw | All leads including duplicates, sorted most recent first |
| Duplicates | Leads removed by deduplication |
| LeadCache | Hidden. Lightweight cache used by dashboard triggers |
