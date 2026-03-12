# Claims Management System

A Google Apps Script that automates the end-to-end claims workflow for a hall residence Finance team — from form submission to document generation.

---

## Why I Built This

After 2 years across Finance CCAs, I kept running into the same problem, a significant chunk of the work was repetitive and mechanical like copying data between forms, filling in the same fields over and over and manually chasing documents. Tasks that felt like they existed just to consume time rather than add any real value.

I always suspected most of it could be automated, but never had the technical background to act on it. Now that I do, and with Claude, I finally built the system I wished had existed all along.

This is not just the workflow that is automated, the entire setup is automated as well. A single script builds every sheet, creates the Google Form, sets up the folder structure, and wires everything together. No manual configuration, no piecing things together step by step.

This is my attempt to make a meaningful, lasting change to how the Finance team operates, so that everyone who comes after spends less time on menial work, and more time on work that actually matters.

---

## What It Does

When a CCA Treasurer submits a claim, the system handles everything automatically:

- Parses form responses and organises claim data into a structured master sheet
- Generates a pre-filled email for the CCA Treasurer to forward to the Finance inbox
- Auto-generates the LOA, Summary, and RFP claim documents from Google Docs/Sheets templates
- Saves all generated documents to the correct Google Drive folders
- Tracks the status of every claim from submission through to reimbursement

---

## Sheets Created on Setup

| Sheet | Purpose |
|---|---|
| **Config** | Store Finance D details, template IDs, and folder/form links |
| **Master Sheet** | Dashboard showing all claims and their current status |
| **Claims Data** | Structured data store for all parsed claim entries (hidden) |
| **Add Claim** | Paste a form response here to add a claim to the system |
| **Form Options** | Manages the portfolio and CCA dropdown options for the Google Form |
| **Identifier Data** | Maps WBS account names to their codes and short forms |
| **Form Responses** | Auto-created and linked to the Google Form |

---

## Menu Options

Once set up, a **Claims Tools** menu appears in the spreadsheet:

- **Add Claim** — Reads the pasted form response in the Add Claim sheet and appends a new row to Claims Data and Master Sheet
- **Generate Emails** — For selected rows in Master Sheet, generates and sends a pre-filled email to the CCA Treasurer
- **Generate Forms** — Auto-generates the LOA, Summary, and RFP documents for selected claims
- **Compile Forms** — Compiles all generated documents for a claim into a single PDF package

---

## One-Time Setup

1. Create a new Google Sheet
2. Go to **Extensions > Apps Script**
3. Delete all default code, paste the entire script, and save
4. Run `setupClaimsSystem` (click **Run** in the toolbar)
5. Grant permissions when prompted
6. Fill in the **Config** sheet:
   - Academic year, Finance D name, matric, and phone number
   - Template IDs for your Summary (Google Sheets) and RFP (Google Docs) templates
7. Enable the Drive API: **Services (+) > Drive API v3**
8. Open the generated Google Form and manually add **2 file upload questions per receipt section** (10 total), placed after each *Amount* field:
   - `Receipt Softcopy/Image [N]`
   - `Bank Transaction Screenshot [N]`
9. Share the form link (saved in Config) with claimers

---

## Key Configuration

All settings live in the **Config** sheet:

| Setting | Description |
|---|---|
| `ACADEMIC_YEAR` | e.g. `2526` |
| `FINANCE_D_NAME` | Finance Director's full name |
| `FINANCE_D_MATRIC` | Finance Director's matric number |
| `FINANCE_D_PHONE` | Finance Director's phone number |
| `SUMMARY_TEMPLATE_ID` | Google Sheets template file ID |
| `RFP_TEMPLATE_ID` | Google Docs template file ID |

Folder IDs and the Form URL are populated automatically during setup.

---

## Updating CCA Options

The Google Form dropdowns do **not** update automatically when the CCA list changes. After editing the **Form Options** sheet, run `recreateClaimsForm()` manually from the Apps Script editor to rebuild the form.

---

## Notes

- Each claim reference code follows the format `AY-PORTFOLIO-CCA-001`
- The system tracks each receipt's description, category, GST, company, date, receipt number, amount, softcopy, and bank screenshot
- The **Claims Data** sheet is hidden but stores all structured claim data
- `Claim Data Template` is also hidden and is used to reset the Add Claim sheet after each submission
