# Claims Management System

A Google Apps Script that automates the end-to-end claims workflow for a hall residence Finance team — from form submission to document generation.

---

## Why I Built This

After 2 years across Finance CCAs, I kept running into the same problem — a significant chunk of the work was repetitive and mechanical: copying data between forms, filling in the same fields over and over, and manually chasing documents. Tasks that felt like they existed just to consume time rather than add any real value.

I always suspected most of it could be automated, but never had the technical background to act on it. Now that I do, and with Claude, I finally built the system I wished had existed all along.

This is not just the workflow that is automated — the entire setup is automated as well. A single script builds every sheet, creates the Google Form, sets up the folder structure, and wires everything together. No manual configuration, no piecing things together step by step.

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
| **Config** | Stores Finance D details, template IDs, and folder/form links |
| **Master Sheet** | Dashboard showing all claims and their current status |
| **Claims Data** | Structured data store for all parsed claim entries (hidden) |
| **Add Claim** | Paste a form response here to add a claim to the system |
| **Identifier Data** | Maps CCA members to their matric, phone, and email for auto-fill |
| **Form Options** | Manages the portfolio and CCA dropdown options for the Google Form |
| **Finance Team** | List of Finance Team members for the Filled By and Processed By fields |
| **CCA Spending** | Auto-calculated spending tracker broken down by CCA and fund |
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

### Create the Spreadsheet

1. Create a new Google Sheet
2. Go to **Extensions > Apps Script**
3. Delete all default code, paste the entire script, and press **Ctrl + S** to save

### Add Drive API

4. In the Apps Script editor, click **Services** on the left panel
5. Find **Drive API** and click **Add**

### Run Setup

6. Click **Run** in the toolbar to run `setupClaimsSystem`
7. When prompted, click **Review Permissions → Advanced → Go to Untitled project (unsafe)**
8. Click **Select all → Continue**
9. Switch back to the Google Sheet tab and click **Yes** when the dialog appears

### Fill in the Config Sheet

10. Open the **Config** sheet and fill in:
    - Academic year (e.g. `2526`)
    - Finance D name, matric number, and phone number
    - **Summary Template ID** — your Google Sheets claims summary template
    - **RFP Template ID** — your Google Docs RFP template

    To find a template ID, open the file and copy the ID from the URL:
    ```
    https://docs.google.com/document/d/COPY_THIS_PART/edit?usp=sharing
    ```
    > ⚠️ Make sure both templates are set to **Anyone with the link can view** before running

11. Hide the **Config** sheet once done

### Set Up the Finance Team Sheet

12. Open the **Finance Team** sheet and enter the names of all Finance Team members in column A, one per row

### Configure the Google Form

13. Open the **Claims Submission Form** (link is saved in the Config sheet under `FORM_URL`)
14. Find the **Filled By** question and replace the placeholder options (Person 1, Person 2) with the actual Finance Team member names
15. For each of the 5 receipt sections, add **2 file upload questions** after the *Amount* field:
    - `Receipt Softcopy/Image [N]`
    - `Bank Transaction Screenshot [N]`

    Set each to allow up to **10 attachments**. If a warning about 1 GB appears, click **Continue**

16. Go to **Settings → Responses** and change the **Total Size Limit** to **10 GB**

### Final Steps

17. Reload the spreadsheet — the **Claims Tools** menu will now appear
18. Share the form link (from the Config sheet) with CCA Treasurers

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

The Google Form dropdowns do **not** update automatically when the CCA list changes. To update them:

1. Edit the `CCA_DEPARTMENTS` constant in the Apps Script code
2. Run `recreateClaimsForm()` from the Apps Script editor to rebuild the form
3. Manually re-add the **Filled By** names and **file upload questions** to the new form, as these cannot be scripted

---

## Notes

- Each claim reference code follows the format `AY-PORTFOLIO-CCA-001`
- The system tracks each receipt's description, category, GST, company, date, receipt number, amount, softcopy, and bank screenshot
- The **Claims Data** sheet is hidden but stores all structured claim data
- `Claim Data Template` is also hidden and is used to reset the Add Claim sheet after each submission
- File uploads from form responses are automatically saved to a Google Drive folder created by Google Forms — this is a Google limitation and cannot be changed
