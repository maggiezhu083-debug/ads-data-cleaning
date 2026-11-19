---
description: Clean, validate, aggregate, and summarize Meta Ads and
  Google Ads reports from CSV, TSV, and XLSX files. Use for paid-media
  data cleaning, Campaign × Ad reporting, recalculated metrics, video
  metrics, and editable native Excel PivotTables when supported by
  available tools.
name: ad-data-cleaning
---

# Ad Data Cleaning

## 1. Purpose and scope

This Skill is primarily designed for **Meta Ads and Google Ads**
reports. It may process other platforms when explicitly requested, but
must not assume that Meta or Google metric definitions apply to other
platforms.

The workflow must preserve source data, identify platform-specific field
meanings, aggregate base metrics at the requested reporting grain,
recalculate ratios from aggregated base measures, generate a new XLSX
workbook, and validate the result before delivery.

Never invent missing values, silently remove uncertain records, conflate
conversion events, or claim that a native PivotTable was created unless
it has been verified.

## 2. Default output and reporting grain

Create a **new XLSX workbook**. Never overwrite the uploaded source
file.

Default workbook sheets: 1. `Raw Data` --- preserved source data. 2.
`Pivot` --- a genuine, editable native Excel PivotTable, only when
available tools can create and validate one. 3. `Weighted Summary` --- a
formula-based summary with correctly weighted ratios. 4. `QA` ---
optional validation details when useful.

Default reporting grain: **Campaign × Ad**.

Do not aggregate by Ad alone unless explicitly requested. Keep Campaign
and Ad names exactly as supplied. Do not merge names by case-insensitive
or fuzzy matching.

If the user requests Date, Ad Set, Placement, Device, Country, or
another dimension, include it in the actual grouping key. Do not merely
display it without grouping by it.

Process multiple files separately by default. Combine only when
requested and when platform, currency, reporting level, date range, and
metric definitions are compatible.

## 3. Input inspection and preflight

Inspect every input file and relevant worksheet. Identify: - file name
and type; - platform (Meta Ads or Google Ads, where identifiable); -
worksheet names, header row, and data range; - date range, row count,
and reporting level; - currency and possible mixed currencies; -
available dimensions and base metrics; - conversion and conversion-value
fields; - video-view fields; - blank, malformed, subtotal, total, title,
and note rows; - potential duplicate records and material ambiguities.

Do not assume the first row contains headers. Detect report titles,
metadata, blank rows, and totals.

When useful, provide a concise preflight summary before or alongside
processing. Do not present example values as detected facts.

Ask one focused clarification question only when an unresolved issue
could materially change the result, such as purchase-event definition,
currency, reporting level, duplicate status, or grouping dimension.
Otherwise proceed and document assumptions.

## 4. Preserve source data

For XLSX input: - preserve the original workbook and never edit it in
place; - retain the original raw-data worksheet name, columns, values,
and row order as closely as possible; - preserve formatting, formulas,
hidden rows/columns/sheets, filters, Excel Tables, merged cells, data
validation, and named ranges where supported by the tools; - do not
silently flatten formulas into values or remove workbook features; -
disclose any feature that could not be preserved.

For CSV/TSV input, retain the original column names, source values, and
row order in the `Raw Data` worksheet.

Do not rename, reorder, clean, overwrite, or delete source columns in
the preserved raw-data sheet. Normalized fields belong in separate
output sheets.

## 5. Source field mapping

Map fields by meaning, not name similarity alone. The names below are
candidates; confirm their meaning in the actual export.

### 5.1 Meta Ads

  Standard field              Common Meta Ads field
  --------------------------- -------------------------------
  Campaign                    Campaign name
  Ad Set                      Ad set name
  Ad                          Ad name
  Date                        Reporting starts / Day / Date
  Cost                        Amount spent
  Impressions                 Impressions
  Reach                       Reach
  Frequency                   Frequency
  Link Clicks                 Link clicks
  Outbound Clicks             Outbound clicks
  Landing Page Views          Landing page views
  Purchases                   Purchases
  Purchase Conversion Value   Purchases conversion value
  Adds to Cart                Adds to cart
  Checkouts                   Checkouts initiated
  ThruPlays                   ThruPlays
  3-second Plays              3-second video plays

Rules: - Prefer `Amount spent` as Cost. - For default click-based
reporting, use `Link clicks` and label it clearly. Do not silently
substitute `Clicks (all)`, outbound clicks, or landing page views. - Use
`Purchases` and `Purchases conversion value` only when present and
clearly defined. - Do not assume `Results`, generic `Conversions`, or
similarly named fields mean Purchases. - Keep purchases, leads,
add-to-carts, checkouts, and other events separate. - Preserve
attribution context and reporting-window information if included in the
export.

### 5.2 Google Ads

  -----------------------------------------------------------------------
  Standard field                      Common Google Ads field
  ----------------------------------- -----------------------------------
  Campaign                            Campaign

  Ad Group                            Ad group

  Ad                                  Ad name, asset/ad identifier, or
                                      another reliable ad-level field

  Date                                Day / Date

  Cost                                Cost

  Impressions                         Impr. / Impressions

  Clicks                              Clicks

  Interactions                        Interactions

  Conversions                         Conversions

  All Conversions                     All conv.

  Conversion Value                    Conv. value

  All Conversion Value                All conv. value

  Video Views                         Views / TrueView views, depending
                                      on report type

  Engagements                         Engagements, where applicable
  -----------------------------------------------------------------------

Rules: - Use the actual ad or creative name/identifier for Ad. Never
substitute Ad Group for Ad. - If the export has no reliable ad-level
field, do not fabricate one. Report the actual grain and ask whether
another grain is acceptable. - Keep `Conversions` and `All conversions`
separate. - Keep `Conversion value` and `All conversion value`
separate. - Never assume Google Ads `Conversions` means Purchases. It
may include any configured conversion actions. - Map Google Ads
Conversions to Purchase only when the export or user confirms that the
included conversion actions are purchase-only. - If conversion-action
segmentation is available, inspect it to identify purchase-specific
actions. If not, retain the generic Conversions metric and ask before
labeling it Purchase. - Preserve both `Clicks` and `Interactions` when
available. They are not always equivalent, especially across campaign
types.

### 5.3 Cross-platform comparability

Meta Link Clicks and Google Ads Clicks can have different definitions.
Preserve source labels and explain the difference; do not imply perfect
equivalence.

Do not combine Meta and Google totals unless requested and the currency,
reporting grain, date range, attribution/conversion definitions, and
metric scopes are compatible. Otherwise produce separate platform
summaries.

## 6. Data quality and duplicate handling

Check missing Campaign/Ad identifiers, malformed numeric fields, mixed
currencies, date formats, subtotal/total rows, repeated records, mixed
reporting levels, incompatible dimensions, and inconsistent conversion
definitions.

A repeated Campaign × Ad key is not a duplicate. Multiple rows may
represent dates, placements, devices, countries, or other legitimate
breakdowns.

Do not delete rows merely because all visible cell values appear
identical. Identical-looking rows may be legitimate source records.

Remove a record only when accidental duplicate import is confirmed by
reliable evidence, such as source import metadata or a verified
duplicate-export condition. Record the number removed, the rule used,
and the reason. If uncertain, retain the records and flag them.

Never add campaign-level or ad-set-level summary rows to ad-level detail
rows. Select the correct reporting level to avoid double counting.

## 7. Aggregation rules

Aggregate **base measures first** at the requested reporting grain.

Typical base measures: - Cost; - Impressions; - selected Clicks or Link
Clicks; - Interactions, when relevant; - Conversions and All
Conversions, kept separate; - confirmed Purchase count; - matching
Conversion Value and Purchase Conversion Value; - video-view counts; -
other requested event counts.

Never sum or average source-level ratios such as CPC, CTR, ROAS, CVR,
AOV, VTR, or CPV.

The grouping key must include all requested dimensions. Default:
Campaign × Ad.

## 8. Metric definitions and recalculation

Calculate ratios only after aggregating their underlying base measures.

### 8.1 Standard metrics

  Metric             Formula / rule
  ------------------ ---------------------------------------------------------
  Cost               Sum of source Cost
  Impressions        Sum of source Impressions
  Clicks             Sum of the selected source click metric
  CPC                Cost ÷ selected Clicks
  CTR                selected Clicks ÷ Impressions
  ROAS               matching Conversion Value ÷ Cost
  Purchase           Sum of confirmed purchase-event count only
  Purchase CVR       confirmed Purchase ÷ selected Clicks
  AOV                matching Purchase Conversion Value ÷ confirmed Purchase
  Conversions        Sum of source Conversions, distinct from Purchase
  Conversion Value   Sum of the matching source Conversion Value

If a denominator is zero, leave the ratio blank.

### 8.2 Conversion and purchase safeguards

-   `Conversions` is a platform-defined metric, not automatically
    Purchase.
-   `All conversions` is not interchangeable with `Conversions`.
-   Purchase count must come from a confirmed purchase event or
    purchase-only conversion-action scope.
-   Conversion value must use the same event scope as the count used for
    the associated ROAS or AOV.
-   If only generic Conversions and Conversion Value are available,
    calculate `ROAS = Conversion Value ÷ Cost` and label it according to
    that generic conversion scope. Do not label the count Purchase or
    calculate AOV unless purchase scope is confirmed.
-   For a unified custom Purchase CVR, use confirmed Purchase ÷ the
    selected click metric and document the click metric used for each
    platform.
-   If the user requests Google Ads' native Conversion Rate, use the
    applicable platform denominator (often Interactions, depending on
    campaign type) and label it `Google Ads Conversion Rate`. Do not
    silently call it Purchase CVR.
-   For Meta, Purchases ÷ Link Clicks may be used as a custom
    `Meta Purchase CVR`; document the formula and do not imply it is
    necessarily identical to a native Meta metric.

### 8.3 Video metrics

Inspect source headers and report type for video-view fields, including
Meta ThruPlays, Meta 3-second video plays, and Google Ads Views/TrueView
views.

Calculate each event separately:

  -----------------------------------------------------------------------
  Source view metric                  Output metrics
  ----------------------------------- -----------------------------------
  TrueView views                      TrueView VTR = TrueView views ÷
                                      Impressions; TrueView CPV = Cost ÷
                                      TrueView views

  ThruPlays                           ThruPlay VTR = ThruPlays ÷
                                      Impressions; ThruPlay CPV = Cost ÷
                                      ThruPlays

  3-second plays                      3s Play VTR = 3-second plays ÷
                                      Impressions; 3s Play CPV = Cost ÷
                                      3-second plays
  -----------------------------------------------------------------------

Do not combine different view events. Keep the source field name visible
or document the mapping. Do not assume Google Ads Views are equivalent
to Meta ThruPlays or 3-second plays. If a view field is ambiguous, ask
or retain it as an unclassified source metric without inventing a
VTR/CPV definition.

### 8.4 Other metrics

Calculate additional metrics only when required base fields and
definitions are available. Examples: - CPM = Cost ÷ Impressions ×
1,000 - Cost per Result = Cost ÷ clearly identified result count -
Landing Page View Rate = Landing Page Views ÷ selected Clicks, when
compatible fields are available.

Label custom metrics clearly and document formulas.

## 9. Missing values, zeroes, and currencies

-   Treat blank event counts or values as zero only when source-report
    semantics clearly mean no event or no amount.
-   Otherwise preserve missing state and leave affected calculations
    blank or ask.
-   Never silently replace missing values with zero.
-   Leave ratios blank when denominators are zero.
-   Do not aggregate currencies without a user-approved exchange basis
    and date.
-   Keep Cost and Conversion Value in source currency.
-   If multiple currencies appear, separate by currency or ask how to
    proceed.
-   If conversion value is zero or materially inconsistent with
    conversion counts, report source values and flag a possible
    tracking/export issue. Do not infer actual revenue.

## 10. Workbook structure and formatting

Create a new XLSX with: - preserved `Raw Data`; - `Pivot` when a genuine
native PivotTable can be created and validated; - `Weighted Summary`
with correct formulas at the requested grain; - optional `QA` sheet.

The `Weighted Summary` should include requested dimensions, base
measures, and supported recalculated metrics. Keep Conversions and
confirmed Purchase separate where both exist.

Formatting: - store CTR, CVR, VTR, and other rates as numeric
proportions and format as percentages; - format ROAS as a multiple; -
format Cost, Conversion Value, AOV, and CPV as currency in the source
currency; - format counts as numbers; - freeze header rows and use
sensible column widths; - enable filters on ordinary summary tables
where appropriate; - include notes for date range, currency, platform,
reporting grain, metric definitions, and material caveats.

Do not store numeric percentages as text strings containing `%`.

## 11. Native Excel PivotTable requirements

A styled summary table, AutoFilter range, formula table, or ordinary
chart is **not** a native Excel PivotTable.

Create a genuine, editable native Excel PivotTable when available tools
can create and validate one: - source it from retained Raw Data or a
clearly documented normalized range; - use Campaign and Ad as default
row fields, plus user-requested dimensions; - include relevant base
measures as sums; - ensure source range, cache, relationships, and field
configuration are valid; - verify that a compatible spreadsheet
application recognizes it as editable; - do not claim success based only
on PivotTable-related XML/ZIP entries.

Traditional Excel PivotTable calculated fields may calculate ratios on
source rows and then aggregate them, producing incorrectly weighted
results. Prefer verified Power Pivot/DAX measures or another calculation
at the pivot grain. Keep `Weighted Summary` as the authoritative ratio
summary when the PivotTable cannot reliably calculate weighted ratios.

### Capability and fallback

Inspect the actual available spreadsheet tools before promising a native
PivotTable.

If one cannot be created or safely preserved: 1. Do not fabricate
PivotTable structures or claim success. 2. Create the formula-based
`Weighted Summary` if possible. 3. State that the native PivotTable was
not created or could not be verified. 4. Explain the limitation briefly.
5. If a native PivotTable is essential, ask whether the user accepts a
formula summary as a substitute.

If the PivotTable cache refreshes only when opened in Excel, state this
and retain a current `Weighted Summary` companion.

## 12. Validation and QA

Before delivery, validate:

**Source preservation** - Raw Data sheet exists. - Original columns,
values, and row order are preserved as closely as possible. - Row count
matches the source, except for documented non-record rows or confirmed
duplicate handling. - No source values were silently changed. -
Unsupported workbook features are disclosed.

**Aggregation** - Campaign × Ad keys, or the user's requested keys, are
unique in the summary. - Summary rows were not double-counted with
detail rows. - Totals reconcile for Cost, Impressions, selected Clicks,
Conversions, Purchase, Conversion Value, and video-view counts where
applicable. - Differences are explained.

**Metrics** - Ratios are recalculated from aggregated base measures. -
Conversion and Purchase definitions are not conflated. - Conversion
counts and values have matching event scopes. - Currency is
consistent. - Zero denominators do not create errors. - Video metrics
use the correct separate view-event definitions.

**PivotTable** - Confirm whether a genuine native PivotTable exists. -
Confirm whether a compatible spreadsheet application recognizes it as
editable. - Confirm source range and fields. - State whether it
refreshes on open. - If not verified, mark it as not verified.

Include a concise QA report in the final response. Only mark a check as
Passed when it was actually performed.

## 13. Final delivery

Deliver XLSX as the primary output. A CSV copy of `Weighted Summary` may
be provided as a convenience, but it cannot replace the multi-sheet
workbook.

State: - output file link; - platform and date range; - reporting
grain; - currency; - key metric definitions and assumptions; - material
data issues; - whether the native PivotTable was created and verified; -
whether it refreshes on open, if applicable.

Keep the explanation concise but sufficiently detailed for the user to
review and trust the output.

## 14. Interaction rules

-   When the user provides a file and specifies dimensions/metrics,
    proceed without asking again about defaults.
-   Use Campaign × Ad by default.
-   Ask only about material ambiguities that could change results.
-   Do not ask the user to reconfirm rules already defined here.
-   Process multiple files separately unless combination is requested.
-   Never overwrite the source file.
-   Do not claim to have used Excel, LibreOffice, Power Pivot, or
    another application unless actually used.
-   Do not claim a native PivotTable is editable unless verified.
-   Distinguish confirmed facts, assumptions, and unresolved issues.
