# Ad Data Cleaning Skill

A data-cleaning and reporting Skill designed primarily for **Meta Ads and Google Ads**.

This Skill helps transform raw advertising reports into structured, reliable, and analysis-ready Excel workbooks. It standardizes platform-specific fields, aggregates campaign and ad-level performance data, recalculates key advertising metrics, identifies data quality issues, and generates organized reporting outputs.

## Key Features

- **Multi-format Data Support:** Process advertising reports in XLSX, CSV, and TSV formats.
- **Platform-specific Field Mapping:** Identify and map common Meta Ads and Google Ads reporting fields.
- **Campaign × Ad Reporting:** Use Campaign × Ad as the default reporting level.
- **Metric Recalculation:** Recalculate CPC, CTR, ROAS, CVR, AOV, and other supported metrics using aggregated base data.
- **Video Performance Analysis:** Detect video-view metrics such as TrueView views, ThruPlays, and 3-second video plays, and calculate corresponding VTR and CPV metrics when applicable.
- **Data Quality Checks:** Identify missing values, potential duplicates, inconsistent reporting levels, mixed currencies, and other data issues.
- **Excel Workbook Generation:** Create a new, structured XLSX workbook with preserved raw data and calculated summaries.
- **PivotTable Support:** Generate a genuine, editable native Excel PivotTable when supported by the available tools and successfully verified.

## Supported Platforms

| Platform | Supported Reporting Data |
|---|---|
| Meta Ads | Campaign, Ad Set, Ad, Cost, Impressions, Clicks, Purchases, Conversion Value, Video Views, etc. |
| Google Ads | Campaign, Ad Group, Ad, Cost, Impressions, Clicks, Interactions, Conversions, Conversion Value, Video Views, etc. |

The Skill primarily supports Meta Ads and Google Ads. Other advertising platforms can be processed when explicitly requested, but their metric definitions must be checked independently.

## Default Reporting Metrics

The Skill supports the following metrics when the required source fields are available:

| Metric | Calculation |
|---|---|
| CPC | Cost ÷ Clicks |
| CTR | Clicks ÷ Impressions |
| ROAS | Conversion Value ÷ Cost |
| CVR | Confirmed Purchase or matching Conversion Count ÷ Clicks |
| AOV | Purchase Conversion Value ÷ Purchases |
| CPM | Cost ÷ Impressions × 1,000 |
| VTR | Relevant Video Views ÷ Impressions |
| CPV | Cost ÷ Relevant Video Views |

Metric definitions depend on the source platform and available conversion or video-view fields. The Skill must not assume that generic conversions represent purchases or that different video-view events are interchangeable.

## Default Reporting Logic

- **Default reporting grain:** Campaign × Ad.
- Campaign and Ad names must be preserved exactly as provided in the source data.
- Additional dimensions such as Date, Ad Set, Placement, Device, and Country can be included when requested.
- Multiple source files are processed separately by default.
- Base metrics are aggregated first, and ratio metrics are recalculated after aggregation.
- Uncertain duplicate records must not be deleted automatically.

## Excel Output Structure

The primary deliverable is a new XLSX workbook.

| Worksheet | Description |
|---|---|
| Raw Data | Preserved original source data |
| Pivot | Genuine, editable native Excel PivotTable, if successfully created and verified |
| Weighted Summary | Formula-based performance summary with correctly recalculated metrics |
| QA | Data quality checks, validation results, and relevant notes |

If native Excel PivotTable creation is not supported or cannot be verified, the Skill must clearly disclose this limitation and provide a Weighted Summary where possible.

A formula-based summary table must never be presented as a native Excel PivotTable.

## Data Quality & Validation

Before delivering the final workbook, the Skill checks:

- Source data preservation and row counts.
- Campaign and Ad reporting dimensions.
- Missing or malformed values.
- Potential duplicate records.
- Currency consistency.
- Conversion event definitions.
- Video-view metric definitions.
- Aggregated totals and calculation accuracy.
- PivotTable creation and editability, when applicable.

All assumptions, unresolved issues, and unsupported workbook features must be disclosed.

## How to Use

1. Provide a Meta Ads or Google Ads report in XLSX, CSV, or TSV format.
2. Specify the reporting dimensions or metrics you want to analyze, if different from the defaults.
3. The Skill inspects the source data and identifies relevant fields and potential data issues.
4. The assistant cleans and aggregates the data, recalculates supported metrics, and validates the results.
5. Receive a new Excel workbook with the processed data and a concise QA summary.

### Example Request

> Please clean this Meta Ads report. Use Campaign × Ad as the reporting level, preserve the original raw data, calculate all available performance and video metrics, and generate a new Excel workbook.

## Repository Structure

```text
ad-data-cleaning/
├── README.md
└── SKILL.md
```

- `README.md`: Project overview, supported platforms, key features, and usage instructions.
- `SKILL.md`: Detailed execution instructions, field mappings, metric definitions, data-cleaning rules, Excel output requirements, and QA procedures.

## Version

**Current Version:** 1.1

**Primary Platforms:** Meta Ads & Google Ads
