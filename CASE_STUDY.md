# FedEx Ground: Automated Delivery Reconciliation System (Excel/VBA)
## Overview
A modernized Excel/VBA system that replaced a multi-hour nightly workflow for validating packages flagged by the system as missing final delivery confirmation (status code 85) and for managing reconciliation tasks. The tool supports Operations Administrators (Ops Admins) through automated data retrieval from multiple internal tracking tools and detail views, intelligent parsing to extract and structure fields with varying formats and optional components, and automated population of standardized delivery form templates and follow-up records.

## Before - Pain Points
- Daily 60–100 packages required manual investigation.
- Ops Admins spent 1–3 hours validating code 85 assignments.
- Tracking details required 3–5 different internal pages per package.
- Handsheets (manual delivery records) were handwritten and time-consuming.
- No reliable way to monitor outstanding unreconciled packages.
- Follow-up tracking required re-pulling old reports and manual comparison.
**Total workflow: 2–4 hours of clerical labor every day across multiple employees.**

## What I Built
- An integrated Excel automation toolkit that streamlined the entire code 85 workflow through:
- Automated data ingestion + validation
- Tracking history retrieval and interpretation
- Template-based delivery form* automation with dynamically generated barcodes
- Workflow management for 30-day follow-up
- Automated reporting and reconciliation lifecycle tracking tools
> \* The handsheet format itself was an official company form; the work here focused on automating population, duplication, and lifecycle handling of these standardized templates rather than designing the form layout.

## Key Automations
### 1. Tracking & Validation Automation
- Automated retrieval of tracking data from internal web-based tools
- Extracts shipment and package details from multiple internal tracking tools and detail views
- Parses complex HTML structures to pull tables and embedded data fields
- Classifies and normalizes address components (recipient, street, suite, city/state, zip) even when lines are missing or merged
- Detects corrected or updated addresses across pages and applies selection logic
- Interprets tracking history to determine whether a missing-scan status (code 85) is valid
- Status tagging for reconciled, suspect, or missing packages
### 2. Digital Handsheet Generator
- Template-based, legal-sized digital handsheet identical to the standardized physical form
- Auto-fill with package data
- Includes inline package details (removing need for stapled pages)
- Automatic overflow handling (new sheets generated as needed)
- Generates 2D barcodes tied to unique record IDs
### 3. Daily Workflow Tools
- "Morning Ops" tab to batch-retrieve up to 15 additional tracking and delivery fields
- Batch validation and research of selected packages via a single select-and-click workflow
- One-click data fetch for any selected range
- Driver lookup tool (with fallback paste-into-form input)
- Automated import + parsing of driver table
### 4. 30-Day Follow-Up & Resolution Tracking
- Outstanding table with auto-research button for each row or batch
- Automatic grouping + printing by driver and contractor
- Resolve-date dropdown + event macro that moves records to Resolved tab
- Organized workflow supporting Ops Admins and management oversight

## Impact
- Reduced nightly validation work for Ops Admins from **1–3 hours to under 10 minutes**
- Eliminated handwriting of **100+ handsheets per week**
- Simplified and centralized multiple workflows across Ops Admins
- Provided management with high-level visibility into reconciliation trends by contractor.

## Technical Stack
- Excel VBA (macros, forms, modules)
- Web data extraction via Excel Power Query (Get & Transform), orchestrated and parameterized through VBA
- Template-based worksheet duplication + automated formatting
- 2D barcode creation
- Event-driven automation (Worksheet_Change, button actions)
- Parsing and transformation logic for variably structured data

