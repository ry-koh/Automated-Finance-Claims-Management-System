function setupClaimsSystem() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'Setup Claims Management System',
    'This will create all necessary sheets and folder structure.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    ui.alert('Setup cancelled.');
    return;
  }
  
  try {
    // Create all sheets
    createConfigSheet(ss);
    createMasterSheet(ss);
    createClaimsDataSheet(ss);
    createAddClaimTemplate(ss);
    createIdentifierDataSheet(ss);
    createFormOptionsSheet(ss);

    let originalSheet = ss.getSheetByName('Sheet1');
    if (originalSheet) ss.deleteSheet(originalSheet);
    
    // Create folder structure
    const config = loadConfigFromSheet(ss);
    createFolderStructure(ss, config);

    // FIX 3: Only create the form if Form Responses doesn't exist; otherwise ask
    const existingFormSheet = ss.getSheetByName('Form Responses');
    if (!existingFormSheet) {
      createClaimsForm(ss);
    } else {
      const formResponse = ui.alert(
        'Form Already Exists',
        'A Form Responses sheet already exists. Recreate the Google Form and reset responses?\n\n(Click No to keep the existing form)',
        ui.ButtonSet.YES_NO
      );
      if (formResponse === ui.Button.YES) {
        createClaimsForm(ss);
      }
    }
    
    ui.alert(
      'Setup Complete!\n\nNext steps:\n1. Fill in the Config sheet (template IDs, Finance D info, academic year)\n2. The Google Form has been created — URL saved in Config sheet\n3. ⚠️  Open the form and manually add 2 file upload fields per receipt section (10 total):\n   • "Receipt Softcopy/Image [N]"\n   • "Bank Transaction Screenshot [N]"\n   Place each pair after "Amount [N]", before "Are there more receipts?"\n4. Share the form link with claimers',
      ui.ButtonSet.OK
    );
    
    ss.setActiveSheet(ss.getSheetByName('Config'));
    
  } catch (e) {
    ui.alert('Error', `Setup failed: ${e.message}`, ui.ButtonSet.OK);
    console.error('Setup error:', e);
  }
}

// ============================================================================
// FORM OPTIONS SHEET
// ============================================================================

function createFormOptionsSheet(ss) {
  let sheet = ss.getSheetByName('Form Options');
  
  if (sheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Form Options Exists',
      'Reset Form Options sheet? (This erases current portfolio/CCA options)',
      ui.ButtonSet.YES_NO
    );
    if (response === ui.Button.YES) {
      ss.deleteSheet(sheet);
    } else {
      return sheet;
    }
  }
  
  sheet = ss.insertSheet('Form Options', 5);

  // Header row
  sheet.getRange(1, 1, 1, 2)
    .setValues([['Portfolio', 'CCA']])
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // Populate from the CCA_DEPARTMENTS constant
  const portfolios = Object.keys(CCA_DEPARTMENTS);
  let currentRow = 2;

  portfolios.forEach(portfolio => {
    const ccas = CCA_DEPARTMENTS[portfolio];
    ccas.forEach(cca => {
      sheet.getRange(currentRow, 1).setValue(portfolio);
      sheet.getRange(currentRow, 2).setValue(cca);
      currentRow++;
    });
  });

  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 200);
  sheet.deleteRows(currentRow, sheet.getMaxRows() - currentRow + 1);

  // Add note for users
  sheet.getRange('A1').setNote('Add or remove Portfolio/CCA pairs here. These will be used in the Google Form dropdowns. Re-run "Create Claims Form" from the Claims Tools menu to update the form.');

  return sheet;
}

// ============================================================================
// GOOGLE FORM CREATION
// ============================================================================

/**
 * Reads Portfolio and CCA data from the Form Options sheet.
 * Returns: { portfolios: [...], ccaByPortfolio: { Portfolio: [...ccas] } }
 */
function loadFormOptions(ss) {
  const sheet = ss.getSheetByName('Form Options');
  if (!sheet) throw new Error('Form Options sheet not found. Please run setup first.');

  const data = sheet.getDataRange().getValues();
  const portfolioSet = new Set();
  const ccaByPortfolio = {};

  for (let i = 1; i < data.length; i++) {
    const portfolio = (data[i][0] || '').toString().trim();
    const cca = (data[i][1] || '').toString().trim();
    if (!portfolio || !cca) continue;

    portfolioSet.add(portfolio);
    if (!ccaByPortfolio[portfolio]) ccaByPortfolio[portfolio] = [];
    ccaByPortfolio[portfolio].push(cca);
  }

  return {
    portfolios: Array.from(portfolioSet),
    ccaByPortfolio
  };
}

/**
 * Creates the Google Form, sets up sections with branching logic,
 * and links responses to a new 'Form Responses' sheet.
 */
function createClaimsForm(ss) {
  const ui = SpreadsheetApp.getUi();
  const { portfolios, ccaByPortfolio } = loadFormOptions(ss);

  // --- Create the Form ---
  const form = FormApp.create('Claims Submission Form');
  form.setTitle('Claims Submission Form');
  form.setAllowResponseEdits(false);
  form.setLimitOneResponsePerUser(false);

  // ---- SECTION 1: Basic Info ----
  // Portfolio dropdown
  const portfolioItem = form.addListItem();
  portfolioItem.setTitle('Portfolio')
    .setRequired(true)
    .setChoices(portfolios.map(p => portfolioItem.createChoice(p)));

  // CCA dropdown — flat list of all CCAs (Google Forms doesn't support dynamic filtering)
  // We include all CCAs and note that claimers should pick the correct one.
  // A separate "CCA" question per portfolio would require many sections; flat list is the practical approach.
  const allCcas = [];
  portfolios.forEach(p => {
    (ccaByPortfolio[p] || []).forEach(cca => {
      if (!allCcas.includes(cca)) allCcas.push(cca);
    });
  });

  const ccaItem = form.addListItem();
  ccaItem.setTitle('CCA')
    .setRequired(true)
    .setChoices(allCcas.map(c => ccaItem.createChoice(c)));

  // Claim Description
  form.addTextItem()
    .setTitle('Claim Description')
    .setRequired(true);

  // Other people involved
  form.addParagraphTextItem()
    .setTitle('Are there other people involved? If yes, include their emails.')
    .setHelpText('Leave a comma and space between emails\nE.g. test123@gmail.com, test456@gmail.com, test789@gmail.com')
    .setRequired(false);

  // Remarks
  form.addParagraphTextItem()
    .setTitle('Remarks')
    .setRequired(false);

  // ---- RECEIPT SECTIONS (1–5) ----
  // Section indices: Section 1 is index 0.
  // We'll create Sections 2–6 for receipts 1–5.
  // After the last receipt (5), no more branching needed.

  const receiptSections = [];

  for (let i = 1; i <= 5; i++) {
    const section = form.addPageBreakItem();
    section.setTitle(`Receipt ${i}`);
    receiptSections.push(section);

    form.addTextItem()
      .setTitle(`Description of Expense ${i}`)
      .setHelpText('What was purchased?')
      .setRequired(true);

    form.addTextItem()
      .setTitle(`Company ${i}`)
      .setHelpText('Name of vendor / merchant')
      .setRequired(false);

    // FIX 2: Date question with DD/MM/YYYY help text to guide claimers
    form.addDateItem()
      .setTitle(`Date of Receipt/Invoice ${i}`)
      .setHelpText('DD/MM/YYYY')
      .setRequired(true);

    form.addTextItem()
      .setTitle(`Receipt/Invoice No. ${i}`)
      .setHelpText('If no number, write "-"')
      .setRequired(true);

    form.addTextItem()
      .setTitle(`Amount ${i}`)
      .setHelpText('Enter amount in numbers only, e.g. 12.50')
      .setRequired(true);

    // ⚠️ MANUAL STEP REQUIRED after setup — open the form in Google Forms UI
    // and add these two questions after "Amount ${i}", before "Are there more receipts?":
    //   1. "Receipt Softcopy/Image ${i}"     — File upload | Required | Images + PDF | Max 10MB
    //   2. "Bank Transaction Screenshot ${i}" — File upload | Required | Images + PDF | Max 10MB
    // Do this for all 5 receipt sections (10 upload fields total).

    // "More receipts?" branching — only add for receipts 1–4
    if (i < 5) {
      const moreReceiptsItem = form.addMultipleChoiceItem();
      moreReceiptsItem.setTitle('Are there more receipts?')
        .setRequired(true);

      // Choices will be set with branching after all sections are created
      // Store a reference to set branching logic after all sections exist
      receiptSections[i - 1]._moreReceiptsItem = moreReceiptsItem;
    }
  }

  // ---- SET BRANCHING LOGIC ----
  // Now that all sections exist, set the "Yes → next section, No → submit" choices
  for (let i = 0; i < 4; i++) {
    const moreReceiptsItem = receiptSections[i]._moreReceiptsItem;
    const nextSection = receiptSections[i + 1]; // next receipt section

    moreReceiptsItem.setChoices([
      moreReceiptsItem.createChoice('Yes', nextSection),                      // Go to next receipt
      moreReceiptsItem.createChoice('No', FormApp.PageNavigationType.SUBMIT)  // Submit form
    ]);
  }

  // ---- LINK FORM TO SPREADSHEET ----
  // Delete existing Form Responses sheet if present (will be recreated by linkFormToSheet)
  const existingResponseSheet = ss.getSheetByName('Form Responses');
  if (existingResponseSheet) ss.deleteSheet(existingResponseSheet);

  form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  // The above creates a sheet named "Form Responses (1)" or similar — rename it
  SpreadsheetApp.flush();
  Utilities.sleep(2000); // Give Sheets time to create the linked sheet

  const allSheets = ss.getSheets();
  const linkedSheet = allSheets.find(s =>
    s.getName().startsWith('Form Responses') && s.getName() !== 'Form Responses'
  );
  if (linkedSheet) linkedSheet.setName('Form Responses');

  // Insert 'Processed' and 'Error' columns as the two leftmost columns
  const formResponsesSheet = ss.getSheetByName('Form Responses');
  if (formResponsesSheet) {
    formResponsesSheet.insertColumnsBefore(1, 2);
    formResponsesSheet.getRange('A1').setValue('Processed');
    formResponsesSheet.getRange('B1').setValue('Error');
    formResponsesSheet.getRange('A1:B1')
      .setBackground('#4a86e8')
      .setFontColor('#ffffff')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    formResponsesSheet.setColumnWidth(1, 100);
    formResponsesSheet.setColumnWidth(2, 200);

    // FIX 2: Format date columns in Form Responses to DD/MM/YYYY
    // Timestamp is col 3 (C), receipt dates are at cols 11, 17, 23, 29, 35 (1-indexed)
    // after the 2 prepended cols: Processed(1), Error(2), Timestamp(3), Portfolio(4), CCA(5),
    // ClaimDesc(6), OtherEmails(7), Remarks(8), then Receipt blocks start at col 9
    // Each receipt block: Desc(+0), Company(+1), Date(+2), ReceiptNo(+3), Amount(+4), Softcopy(+5), Bank(+6) = 7 cols
    // Receipt 1 date = col 9+2 = 11, Receipt 2 = 11+7 = 18... wait, block is 7 wide so:
    // R1 date=11, R2 date=18, R3 date=25, R4 date=32, R5 date=39
    // But file upload fields add 2 per receipt (softcopy + bank), so each block is actually 9 cols wide
    // with file uploads: Desc, Company, Date, ReceiptNo, Amount, Softcopy, Bank, [MoreReceipts for 1-4] = varies
    // Conservative approach: apply format to the entire sheet date-like columns using number format
    formResponsesSheet.getRange(1, 3, formResponsesSheet.getMaxRows(), 1)
      .setNumberFormat('DD/MM/YYYY HH:mm:ss'); // Timestamp col
    // Receipt date cols: R1=9+2=11, R2=15+2=17, R3=21+2=23, R4=27+2=29, R5=33+2=35
    // Receipt date cols: R1=7+2=9, R2=13+2=15, R3=21, R4=27, R5=33
    // Receipt date cols: R1=7+2=9, R2=13+2=15, R3=19+2=21, R4=25+2=27, R5=31+2=33
    [9, 15, 21, 27, 33].forEach(col => {
      formResponsesSheet.getRange(2, col, formResponsesSheet.getMaxRows() - 1, 1)
        .setNumberFormat('DD/MM/YYYY');
    });

    // Add a trigger to auto-insert checkboxes in A and B for each new form response
    // Also pre-fill checkboxes for any existing data rows
    const lastRow = formResponsesSheet.getLastRow();
    if (lastRow > 1) {
      formResponsesSheet.getRange(2, 1, lastRow - 1, 2).insertCheckboxes();
    }
  }

  // Save Form URL to Config sheet
  const configSheet = ss.getSheetByName('Config');
  if (configSheet) {
    const configData = configSheet.getRange('A:A').getValues();
    let formUrlRow = -1;
    for (let i = 0; i < configData.length; i++) {
      if (configData[i][0] === 'FORM_URL') {
        formUrlRow = i + 1;
        break;
      }
    }
    if (formUrlRow === -1) {
      // Append after last config row
      configSheet.appendRow(['FORM_URL', form.getPublishedUrl(), 'Share this link with claimers']);
    } else {
      configSheet.getRange(formUrlRow, 2).setValue(form.getPublishedUrl());
    }
  }

  // Install trigger to auto-add checkboxes on each new form submission
  installFormSubmitTrigger(ss);

  console.log('Form created: ' + form.getPublishedUrl());
  return form;
}

/**
 * Regenerate the Google Form (e.g. after updating Form Options sheet).
 * Accessible from the Claims Tools menu.
 */
function recreateClaimsForm() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  const response = ui.alert(
    'Recreate Claims Form',
    'This will delete the existing form and create a new one linked to this sheet.\n\nExisting form responses will NOT be deleted from this sheet, but the old form link will stop working.\n\nContinue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('Cancelled.');
    return;
  }

  try {
    createClaimsForm(ss);
    const formUrl = getConfigValue('FORM_URL');
    ui.alert('Form recreated successfully!\n\nNew form URL:\n' + formUrl);
  } catch (e) {
    ui.alert('Error', `Failed to recreate form: ${e.message}`, ui.ButtonSet.OK);
    console.error(e);
  }
}

// ============================================================================
// FORM SUBMIT TRIGGER
// ============================================================================

/**
 * Automatically called when a form response is submitted.
 * Inserts checkboxes in the Processed (col A) and Error (col B) columns
 * for the newly added row.
 */
function onFormSubmitInsertCheckboxes(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Form Responses');
    if (!sheet) return;

    // The new response is always appended at the last row
    const newRow = sheet.getLastRow();
    if (newRow < 2) return;

    sheet.getRange(newRow, 1, 1, 2).insertCheckboxes();
  } catch (err) {
    console.error('onFormSubmitInsertCheckboxes error: ' + err.message);
  }
}

/**
 * Installs the onFormSubmit trigger for the spreadsheet.
 * Called automatically during setup via createClaimsForm.
 */
function installFormSubmitTrigger(ss) {
  // Remove any existing triggers for this function to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onFormSubmitInsertCheckboxes') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Install a fresh trigger
  ScriptApp.newTrigger('onFormSubmitInsertCheckboxes')
    .forSpreadsheet(ss)
    .onFormSubmit()
    .create();
}

// ============================================================================
// PART 2: EXISTING SHEET CREATION FUNCTIONS (UNCHANGED)
// ============================================================================

function createConfigSheet(ss) {
  let sheet = ss.getSheetByName('Config');
  
  if (sheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Config Exists',
      'Reset Config sheet? (This erases current settings)',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      ss.deleteSheet(sheet);
    } else {
      return sheet;
    }
  }
  
  sheet = ss.insertSheet('Config', 0);
  
  const data = [
    ['Setting', 'Value', 'Description'],
    ['ACADEMIC_YEAR', '', '=IF(B2<>"","","REQUIRED: Academic Year (e.g. 2526)")'],
    ['FINANCE_D_NAME', '', '=IF(B3<>"","","REQUIRED: Finance D Name")'],
    ['FINANCE_D_MATRIC', '', '=IF(B4<>"","","REQUIRED: Finance D Matric No.")'],
    ['FINANCE_D_PHONE', '', '=IF(B5<>"","","REQUIRED: Finance D Phone No.")'],
    ['', '', ''],
    ['[TEMPLATE FILE IDs]', '', ''],
    ['SUMMARY_TEMPLATE_ID', '', '=IF(B8<>"","","REQUIRED: Google Sheets template ID")'],
    ['RFP_TEMPLATE_ID', '', '=IF(B9<>"","","REQUIRED: Google Docs template ID")'],
    ['', '', ''],
    ['[FOLDERS - Auto-created]', '', ''],
    ['MAIN_FOLDER_ID', '', 'Main claims folder'],
    ['RFP_FOLDER_ID', '', 'RFPs subfolder'],
    ['', '', ''],
    ['[FORM - Auto-created]', '', ''],
    ['FORM_URL', '', 'Share this link with claimers'],
  ];
  
  const range = sheet.getRange(1, 1, data.length, 3);
  range.setValues(data).setHorizontalAlignment('center');
  
  sheet.getRange(1, 1, 1, 3)
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 300);
  sheet.setColumnWidth(3, 400);

  sheet.deleteRows(data.length + 1, sheet.getMaxRows() - data.length);
  sheet.deleteColumns(4, sheet.getMaxColumns() - 3);

  // Conditional formatting for required fields
  const requiredRange1 = sheet.getRange('B2:B5');
  const requiredRange2 = sheet.getRange('B8:B9');
  
  const emptyRule = SpreadsheetApp.newConditionalFormatRule()
    .whenCellEmpty()
    .setBackground('#f4cccc')
    .setRanges([requiredRange1, requiredRange2])
    .build();
  
  const rules = sheet.getConditionalFormatRules();
  rules.push(emptyRule);
  sheet.setConditionalFormatRules(rules);

  const protection = sheet.protect().setDescription('Protected with exceptions');
  protection.setWarningOnly(true);
  const ranges = sheet.getRangeList(['B2:B5', 'B8:B9']).getRanges();
  protection.setUnprotectedRanges(ranges);
  
  return sheet;
}

function createMasterSheet(ss) {
  let sheet = ss.getSheetByName('Master Sheet');

  if (sheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Master Sheet Exists',
      'Reset Master Sheet? (This erases current settings)',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      ss.deleteSheet(sheet);
    } else {
      return sheet;
    }
  }

  sheet = ss.insertSheet('Master Sheet', 1);
  
  const headers = [
    'No.', 'PORTFOLIO', 'CCA', 'ITEM', 'AMOUNT', 'PERSON CLAIMING', 'MOBILE NO.',
    'REFERENCE CODE', 'WBS ACCOUNT', 'Generate Email', 'Emails Sent', 'Generate Forms',
    'Forms Generated', 'Email Screenshot Added', 'Formatting Remarks', 'Compile Forms',
    'Compiled', 'PROCESSED TO FINANCE DIRECTOR', 'SUBMISSION TO OFFICE',
    'DATE OF REIMBURSEMENT', 'STATUS', 'REMARKS (For my own use)', 'FOR RL USE ONLY'
  ];

  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setWrap(true);

  // Pastel green for editable columns: J(10), L(12), N(14), O(15), P(16)
  const pastelGreen = '#d9ead3';
  [10, 12, 14, 15, 16].forEach(col => {
    sheet.getRange(1, col).setBackground(pastelGreen).setFontColor('#000000');
  });

  // Pastel red for Finance D columns: R(18) to W(23)
  const pastelRed = '#f4cccc';
  for (let col = 18; col <= 23; col++) {
    sheet.getRange(1, col).setBackground(pastelRed).setFontColor('#000000');
  }

  sheet.setColumnWidth(1, 50);
  sheet.setColumnWidth(4, 250);
  sheet.setColumnWidth(8, 200);
  sheet.getRange('E2:E1000').setNumberFormat('"$"#,##0.00');
  sheet.deleteRows(2, sheet.getMaxRows() - 1);
  sheet.deleteColumns(headers.length + 1, sheet.getMaxColumns() - headers.length);

  // Conditional formatting: STATUS col U=21 — green if COMPLETED
  const statusRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('COMPLETED')
    .setBackground('#00b050')
    .setFontColor('#ffffff')
    .setRanges([sheet.getRange('U2:U1000')])
    .build();
  sheet.setConditionalFormatRules([statusRule]);

  const protection = sheet.protect().setDescription('Protected with exceptions');
  protection.setWarningOnly(true);
  const ranges = sheet.getRangeList(['J:J', 'L:L', 'N:P']).getRanges();
  protection.setUnprotectedRanges(ranges);
  
  return sheet;
}

function createClaimsDataSheet(ss) {
  let sheet = ss.getSheetByName('Claims Data');

  if (sheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Claims Data Exists',
      'Reset Claims Data Sheet? (This erases current settings)',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      ss.deleteSheet(sheet);
    } else {
      return sheet;
    }
  }

  sheet = ss.insertSheet('Claims Data', 2);
  
  const baseHeaders = [
    'No.', 'Finance D Name', 'Finance D Matric No.', 'Finance D Phone No.',
    'Claimer Name', 'Claimer Matric No.', 'Claimer Phone No.', 'Email Address',
    'Portfolio', 'CCA', 'Claim Description', 'Total Claim Amount', 'Date',
    'Reference Code', 'WBS Account Name', 'WBS No.', 'Remarks', 'Other Emails Involved'
  ];
  
  const receiptHeaders = [
    'DR/CR', 'Description', 'Category', 'Category Code', 'GST Code',
    'Company', 'Date', 'Receipt No.', 'Amount', 'Softcopy', 'Bank'
  ];
  
  const allHeaders = [...baseHeaders];
  for (let i = 1; i <= 5; i++) {
    receiptHeaders.forEach(h => allHeaders.push(`R${i} ${h}`));
  }
  
  sheet.getRange(1, 1, 1, allHeaders.length)
    .setValues([allHeaders])
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  sheet.setColumnWidth(1, 50);
  sheet.deleteRows(2, sheet.getMaxRows() - 1);

  const protection = sheet.protect().setDescription('Protected with exceptions');
  protection.setWarningOnly(true);
  
  sheet.hideSheet();
  return sheet;
}

function createAddClaimTemplate(ss) {
  let sheet = ss.getSheetByName('Add Claim');

  if (sheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Add Claim Sheet Exists',
      'Reset Add Claim Sheet? (This erases current settings)',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      ss.deleteSheet(sheet);
    } else {
      return sheet;
    }
  }

  sheet = ss.insertSheet('Add Claim', 3);
  
  // Title
  sheet.getRange('A1:L1').merge()
    .setValue('ADD NEW CLAIM')
    .setFontSize(18)
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setBackground('#d9ead3');
  
  // Instructions
  sheet.getRange('A3:L3').merge()
    .setValue('Option 1: Paste Google Form response in row 5 below | Option 2: Fill in manually starting row 8')
    .setFontStyle('italic')
    .setHorizontalAlignment('center');
  
  // Google Forms response row header
  sheet.getRange('A4:L4').merge()
    .setValue('PASTE GOOGLE FORM RESPONSE HERE (row 5) →')
    .setBackground('#fce5cd')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');

  // FIX 2: Format I5 as DD/MM/YYYY (Timestamp cell in the pasted form response row)
  sheet.getRange('I5').setNumberFormat('DD/MM/YYYY');
  
  // Claim info section header
  sheet.getRange('A7:B7').merge()
    .setValue('CLAIM INFORMATION')
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  
  // Manual input fields start at row 8
  const startRow = 8;
  const fields = [
    // B8: Finance D Name — from Config B3
    ['Finance D Name', '=IF($A$5<>"",\'Config\'!$B$3,"")'],
    // B9: Finance D Matric No. — from Config B4
    ['Finance D Matric No.', '=IF($A$5<>"",\'Config\'!$B$4,"")'],
    // B10: Finance D Phone No. — from Config B5
    ['Finance D Phone No.', '=IF($A$5<>"",\'Config\'!$B$5,"")'],
    // B11: Claimer Name — VLOOKUP by CCA (B16) from Identifier Data
    ['Claimer Name', '=IF($A$5<>"",IFNA(VLOOKUP($B$16,\'Identifier Data\'!$B:$C,2,FALSE),"Name not in list"),"")'],
    // B12: Claimer Matric No.
    ['Claimer Matric No.', '=IF($A$5<>"",IFNA(VLOOKUP($B$16,\'Identifier Data\'!$B:$D,3,FALSE),"Not in list"),"")'],
    // B13: Claimer Phone No.
    ['Claimer Phone No.', '=IF($A$5<>"",IFNA(VLOOKUP($B$16,\'Identifier Data\'!$B:$E,4,FALSE),"Not in list"),"")'],
    // B14: Email Address
    ['Email Address', '=IF($A$5<>"",IFNA(VLOOKUP($B$16,\'Identifier Data\'!$B:$F,5,FALSE),"Not in list"),"")'],
    // B15: Portfolio — from form response col B (Timestamp=A, Portfolio=B)
    ['Portfolio', '=IF($A$5<>"",B5,"")'],
    // B16: CCA — from form response col C
    ['CCA', '=IF($A$5<>"",C5,"")'],
    // B17: Claim Description — from form response col D
    ['Claim Description', '=IF($A$5<>"",D5,"")'],
    // B18: Total Claim Amount
    ['Total Claim Amount', '=IF($A$5<>"",SUM(E16,G16,I16,K16,M16),"")'],
    // B19: Date — FIX 2: formatted as DD/MM/YYYY
    ['Date', '=IF($A$5<>"",TEXT(TODAY(),"DD/MM/YYYY"),"")'],
    // FIX 1: Reference Code — AY-Portfolio-CCA-No. (zero-padded 3 digits)
    ['Reference Code', '=IF($A$5<>"",CONCATENATE(\'Config\'!$B$2,"-",UPPER($B$15),"-",UPPER($B$16),"-",TEXT(COUNTIF(\'Master Sheet\'!$C:$C,$B$16)+1,"000")),"")'],
    ['WBS Account Name', ''],
    ['WBS No.', '=IF($B$21<>"",SWITCH($B$21, "Student Activity Fund", "H-404-00-000003", "Managed by Hall Fund", "H-404-00-000004", "Master Fund", "E-404-10-0001-01", "Master Fund (RHMP)", "E-404-10-0001-07"),"")'],
    // B23: Remarks — from form response col F
    ['Remarks', '=IF($A$5<>"",F5,"")'],
    // B24: Other Emails Involved — from form response col E
    ['Other Emails Involved', '=IF($A$5<>"",E5,"")'],
    ['', ''],
    ['WBS Account Short Form', '=IF($B$21<>"",SWITCH($B$21, "Student Activity Fund", "SA", "Managed by Hall Fund", "MBH", "Master Fund", "MF", "Master Fund (RHMP)", "MF (RHMP)"),"")']
  ];

  const wbsAccountOptions = [
    'Student Activity Fund',
    'Managed by Hall Fund',
    'Master Fund',
    'Master Fund (RHMP)'
  ];

  const wbsRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(wbsAccountOptions, true)
    .setAllowInvalid(false)
    .setHelpText('Select a WBS Account')
    .build();

  sheet.getRange('B21').setDataValidation(wbsRule).setHorizontalAlignment('center');

  const categoryOptions = [
    'Office Supplies', 'Consumables', 'Sports & Cultural Materials',
    'Other fees (Others)', 'Professional fees', 'Bank Charges',
    'Licensing/Subscription', 'Postage & Telecommunication Charges',
    'Maintenance (Equipment)', 'Lease expense (premises)',
    'Lease expense (rental of equipment)', 'Furniture', 'Equipment Purchase',
    'Publications', 'Meals & Refreshments', 'Local Travel',
    'Student awards/prizes', 'Donation/Sponsorship', 'Miscellaneous Expense',
    'Other Services', 'Fund Transfer', 'N/A'
  ];

  const categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(categoryOptions, true)
    .setAllowInvalid(false)
    .setHelpText('Select expense category')
    .build();

  const drCrRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['DR', 'CR'], true)
    .setAllowInvalid(false)
    .setHelpText('Select DR (Debit) or CR (Credit)')
    .build();

  const gstCodeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['IE', 'I9', 'L9'], true)
    .setAllowInvalid(false)
    .setHelpText('Select GST code')
    .build();
  
  fields.forEach((field, i) => {
    sheet.getRange(startRow + i, 1)
      .setValue(field[0])
      .setFontWeight('bold');
    if (field[1]) {
      sheet.getRange(startRow + i, 2)
      .setFormula(field[1])
      .setHorizontalAlignment('center');
    }
  });
  
  const cols = ['D', 'F', 'H', 'J', 'L'];
  const receiptLabels = ['RECEIPT 1', 'RECEIPT 2', 'RECEIPT 3', 'RECEIPT 4', 'RECEIPT 5'];
  // Column positions in Form Responses row 5 (1-indexed):
  // Column positions in Form Responses row 5 (1-indexed, NO Processed/Error cols pasted):
  // A=1(Timestamp) B=2(Portfolio) C=3(CCA) D=4(ClaimDesc) E=5(OtherEmails) F=6(Remarks)
  // Receipt blocks (6 cols each: Desc, Company, Date, ReceiptNo, Amount, MoreReceipts?):
  //   R1=7, R2=13, R3=19, R4=25, R5=31 (R5 no MoreReceipts, only 5 cols)
  // File uploads at the end:
  // File uploads at the end:
  //   R1 Softcopy=36, R1 Bank=37, R2 Softcopy=38, R2 Bank=39,
  //   R3 Softcopy=40, R3 Bank=41, R4 Softcopy=42, R4 Bank=43,
  //   R5 Softcopy=44, R5 Bank=45
  //   R3 Softcopy=46, R3 Bank=47, R4 Softcopy=48, R4 Bank=49,
  //   R5 Softcopy=50, R5 Bank=51
  //   R3 Softcopy=48, R3 Bank=49, R4 Softcopy=50, R4 Bank=51,
  //   R5 Softcopy=52, R5 Bank=53
  const formColumnStarts = [7, 13, 19, 25, 31];
  const softcopyColumns = [36, 38, 40, 42, 44];
  const bankColumns     = [37, 39, 41, 43, 45];
  
  cols.forEach((col, i) => {
    sheet.getRange(`${col}7`)
      .setValue(receiptLabels[i])
      .setBackground('#93c47d')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    const descColLetter = String.fromCharCode(cols[i].charCodeAt(0) + 1);
    const descCell = `${descColLetter}${startRow + 1}`;
    const receiptFields = [
      ['DR/CR', `=IF(${descCell}="","","DR")`],
      ['Description', `=IF(ISBLANK(INDIRECT("R5C"&${formColumnStarts[i]},FALSE)),"",INDIRECT("R5C"&${formColumnStarts[i]},FALSE))`],
      ['Category', ''],
      ['Category Code', ''],
      ['GST Code', `=IF(${descCell}="","","IE")`],
      ['Company', `=IF(ISBLANK(INDIRECT("R5C"&${formColumnStarts[i]+1},FALSE)),"",INDIRECT("R5C"&${formColumnStarts[i]+1},FALSE))`],
      // FIX 2: Receipt date fields formatted as DD/MM/YYYY
      ['Date', `=IF(ISBLANK(INDIRECT("R5C"&${formColumnStarts[i]+2},FALSE)),"",TEXT(INDIRECT("R5C"&${formColumnStarts[i]+2},FALSE),"DD/MM/YYYY"))`],
      ['Receipt No.', `=IF(ISBLANK(INDIRECT("R5C"&${formColumnStarts[i]+3},FALSE)),"",INDIRECT("R5C"&${formColumnStarts[i]+3},FALSE))`],
      ['Amount', `=IF(ISBLANK(INDIRECT("R5C"&${formColumnStarts[i]+4},FALSE)),"",INDIRECT("R5C"&${formColumnStarts[i]+4},FALSE))`],
      ['Softcopy Link', `=IF(ISBLANK(INDIRECT("R5C"&${softcopyColumns[i]},FALSE)),"",INDIRECT("R5C"&${softcopyColumns[i]},FALSE))`],
      ['Bank Link', `=IF(ISBLANK(INDIRECT("R5C"&${bankColumns[i]},FALSE)),"",INDIRECT("R5C"&${bankColumns[i]},FALSE))`]
    ];
    
    receiptFields.forEach((field, j) => {
      const cellRef = `${col}${startRow + j}`;
      const colLetter = String.fromCharCode(col.charCodeAt(0) + 1);
      const valueCell = `${colLetter}${startRow + j}`;

      sheet.getRange(cellRef)
        .setValue(field[0])
        .setFontSize(9)
        .setFontWeight('bold');
      
      if (field[1]) {
        if (field[1].startsWith('=')) {
          sheet.getRange(valueCell).setFormula(field[1]).setHorizontalAlignment('center');
        } else {
          sheet.getRange(valueCell).setValue(field[1]).setHorizontalAlignment('center');
        }
      } else {
        sheet.getRange(valueCell).setHorizontalAlignment('center');
      }
      
      if (field[0] === 'DR/CR') {
        sheet.getRange(valueCell).setDataValidation(drCrRule);
      } else if (field[0] === 'Category') {
        sheet.getRange(valueCell).setDataValidation(categoryRule);
        
        const categoryCodeCell = `${colLetter}${startRow + j + 1}`;
        const categoryCell = valueCell;
        sheet.getRange(categoryCodeCell).setFormula(
          `=IF(ISBLANK(${categoryCell}),"",SWITCH(${categoryCell},` +
          `"Office Supplies","7100101",` +
          `"Consumables","7100103",` +
          `"Sports & Cultural Materials","7100104",` +
          `"Other fees (Others)","7200108",` +
          `"Professional fees","7200201",` +
          `"Bank Charges","7200213",` +
          `"Licensing/Subscription","7200402",` +
          `"Postage & Telecommunication Charges","7200412",` +
          `"Maintenance (Equipment)","7400112",` +
          `"Lease expense (premises)","7400301",` +
          `"Lease expense (rental of equipment)","7400301",` +
          `"Furniture","7400401",` +
          `"Equipment Purchase","7400401",` +
          `"Publications","7500104",` +
          `"Meals & Refreshments","7500106",` +
          `"Local Travel","7600105",` +
          `"Student awards/prizes","7650119",` +
          `"Donation/Sponsorship","7700101",` +
          `"Miscellaneous Expense","7700701",` +
          `"Other Services","7700715",` +
          `"Fund Transfer","7800201",` +
          `"N/A"))`
        ).setHorizontalAlignment('center');
      } else if (field[0] === 'GST Code') {
        sheet.getRange(valueCell).setDataValidation(gstCodeRule);
      }
    });
  });
  
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 300);

  let rules = sheet.getConditionalFormatRules();

  const emailRequiredCells = ['B8','B9','B10','B11','B12','B13','B14','B15','B16','B17','B18','B19','B20','B21','B22','B23','B24'];
  emailRequiredCells.forEach(cell => {
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($B$14<>"", ${cell}="")`)
      .setBackground('#f4cccc')
      .setRanges([sheet.getRange(cell)])
      .build();
    rules.push(rule);
  });

  const descriptionColumns = ['E', 'G', 'I', 'K', 'M'];
  descriptionColumns.forEach(col => {
    const requiredCells = [`${col}8`, `${col}10`, `${col}11`, `${col}12`];
    requiredCells.forEach(cell => {
      const rule = SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied(`=AND(${col}$9<>"", ${cell}="")`)
        .setBackground('#f4cccc')
        .setRanges([sheet.getRange(cell)])
        .build();
      rules.push(rule);
    });
  });

  sheet.setConditionalFormatRules(rules);
  
  let templateSheet = ss.getSheetByName('Claim Data Template');

  if (templateSheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Claim Data Template Sheet Exists',
      'Reset Claim Data Template Sheet? (This erases current settings)',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      ss.deleteSheet(templateSheet);
    }
  }

  templateSheet = sheet.copyTo(ss);
  templateSheet.setName('Claim Data Template');
  templateSheet.hideSheet();
  const protection = templateSheet.protect().setDescription('Protected with exceptions');
  protection.setWarningOnly(true);
  
  return sheet;
}

function createIdentifierDataSheet(ss) {
  let sheet = ss.getSheetByName('Identifier Data');
  
  if (sheet) {
    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Identifier Data Exists',
      'Reset Identifier Data sheet? (This erases current settings)',
      ui.ButtonSet.YES_NO
    );
    
    if (response === ui.Button.YES) {
      ss.deleteSheet(sheet);
    } else {
      return sheet;
    }
  }
  
  sheet = ss.insertSheet('Identifier Data', 4);
  
  const data = [
    ['Portfolio', 'CCA', 'Name', 'Matric', 'Phone Number', 'Email Address']
  ];
  
  const range = sheet.getRange(1, 1, 1, 6);
  range.setValues(data).setHorizontalAlignment('center');
  
  sheet.getRange(1, 1, 1, 6)
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');
  
  return sheet;
}

function createFolderStructure(ss, config) {
  const parentFolder = DriveApp.getRootFolder();
  
  const mainFolderName = `Automated Claims`;
  let mainFolder = getOrCreateFolderByName(parentFolder, mainFolderName);
  let rfpFolder = getOrCreateFolderByName(mainFolder, 'RFPs');
  
  const configSheet = ss.getSheetByName('Config');
  const data = configSheet.getRange('A:B').getValues();
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === 'MAIN_FOLDER_ID') {
      configSheet.getRange(i + 1, 2).setValue(mainFolder.getId());
    } else if (data[i][0] === 'RFP_FOLDER_ID') {
      configSheet.getRange(i + 1, 2).setValue(rfpFolder.getId());
    }
  }
  
  return {
    mainFolderId: mainFolder.getId(),
    rfpFolderId: rfpFolder.getId()
  };
}

function getOrCreateFolderByName(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function loadConfigFromSheet(ss) {
  const sheet = ss.getSheetByName('Config');
  if (!sheet) {
    return { ACADEMIC_YEAR: '2526' };
  }
  
  const data = sheet.getRange('A:B').getValues();
  const config = {};
  
  data.forEach(row => {
    if (row[0] && row[0] !== 'Setting' && row[0] !== '' && !row[0].startsWith('[')) {
      config[row[0]] = row[1];
    }
  });
  
  return config;
}

// ============================================================================
// PART 3: MENU & UTILITY FUNCTIONS
// ============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu(CONFIG.MENU.TITLE);
  CONFIG.MENU.ITEMS.forEach(item => menu.addItem(item.label, item.handler));
  menu.addToUi();
}

function getSheetOrThrow(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet not found: ${sheetName}`);
  }
  return sheet;
}

function getCellValues(sheet, ranges) {
  return ranges.map(range => sheet.getRange(range).getValue());
}

function buildColumnRangeList(columnLetter, startRow, endRow) {
  const ranges = [];
  for (let r = startRow; r <= endRow; r++) {
    ranges.push(`${columnLetter}${r}`);
  }
  return ranges;
}

function formatDateIfNeeded(value) {
  return value instanceof Date
    ? Utilities.formatDate(value, CONFIG.TIMEZONE, CONFIG.DATE_FORMAT)
    : value;
}

function formatMoney(value) {
  return typeof value === 'number' ? value.toFixed(2) : value;
}

function extractDriveFileIds(cellValue) {
  if (!cellValue) return [];
  return cellValue
    .toString()
    .split(',')
    .map(url => url.trim())
    .map(url => {
      const match = url.match(/[-\w]{25,}/);
      return match ? match[0] : null;
    })
    .filter(Boolean);
}

function forEachReceipt(row, callback) {
  for (let r = 0; r < CONFIG.RECEIPT.COUNT; r++) {
    const base = CONFIG.CLAIMS_COL.DESC_START + (r * CONFIG.RECEIPT.BLOCK_WIDTH);
    if (base + CONFIG.RECEIPT.BLOCK_WIDTH - 1 >= row.length) break;

    callback({
      index: r + 1,
      base,
      drCr: row[base],
      desc: row[base + 1],
      categoryCode: row[base + 3],
      gstCode: row[base + 4],
      company: row[base + 5],
      date: row[base + 6],
      receiptNo: row[base + 7],
      amount: row[base + 8],
      imgUrlCell: row[base + 9],
      bankUrlCell: row[base + 10]
    });
  }
}

function findClaimRow(claimsData, claimNo) {
  return claimsData.find(r => r[CONFIG.CLAIMS_COL.NO] == claimNo);
}

function confirmBatchAction(title, message) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(title, message, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) {
    ui.alert('Operation cancelled.');
    return false;
  }
  return true;
}

function getPendingRowIndices(masterData, generateCol, statusCol) {
  const pending = [];
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][generateCol] === true && masterData[i][statusCol] !== true) {
      pending.push(i);
    }
  }
  return pending;
}

function markMasterStatus(masterSheet, rowIndex, statusCol) {
  masterSheet.getRange(rowIndex + 1, statusCol + 1).setValue(true);
}

function findFileByNameInFolder(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  return files.hasNext() ? files.next() : null;
}

function deleteExistingFileByName(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  while (files.hasNext()) {
    files.next().setTrashed(true);
  }
}

function exportToPdfBlob(file, outputName) {
  const pdfBlob = file.getBlob().getAs('application/pdf');
  return pdfBlob.setName(outputName);
}

function loadPdfLib() {
  if (typeof setTimeout === 'undefined') {
    this.setTimeout = (fn, ms) => {
      Utilities.sleep(ms || 0);
      fn();
      return 0;
    };
  }
  if (typeof clearTimeout === 'undefined') {
    this.clearTimeout = () => {};
  }
  if (typeof PDFLib !== 'undefined') return PDFLib;
  const code = UrlFetchApp.fetch(CONFIG.PDF.LIB_URL).getContentText();
  const factory = new Function(code + '; return PDFLib;');
  return factory();
}

async function mergePdfBlobs(pdfBlobs, outputName) {
  const PDFLib = loadPdfLib();
  const mergedPdf = await PDFLib.PDFDocument.create();

  for (const blob of pdfBlobs) {
    const srcBytes = Uint8Array.from(blob.getBytes());
    const srcPdf = await PDFLib.PDFDocument.load(srcBytes);
    const copiedPages = await mergedPdf.copyPages(srcPdf, srcPdf.getPageIndices());
    copiedPages.forEach(page => mergedPdf.addPage(page));
  }

  const mergedBytes = await mergedPdf.save();
  return Utilities.newBlob(mergedBytes, 'application/pdf', outputName);
}

function buildReceiptListHtml(row) {
  let receiptCount = 0;
  let receiptListHtml = '';

  forEachReceipt(row, receipt => {
    if (receipt.desc && receipt.amount !== '') {
      receiptCount++;
      const formattedAmt = formatMoney(receipt.amount);
      const formattedDate = formatDateIfNeeded(receipt.date);
      receiptListHtml += `<div>#${receiptCount}: $${formattedAmt}, ${receipt.company}, ${receipt.desc}, ${formattedDate}</div>`;
    }
  });

  return receiptListHtml;
}

function collectReceiptAttachments(row, referenceCode) {
  const attachments = [];
  let totalSizeBytes = 0;
  let warningShown = false;

  forEachReceipt(row, receipt => {
    if (!receipt.desc || receipt.amount === '') return;

    [receipt.imgUrlCell, receipt.bankUrlCell].forEach(cellValue => {
      const fileIds = extractDriveFileIds(cellValue);
      fileIds.forEach(fileId => {
        try {
          const file = DriveApp.getFileById(fileId);
          const blob = file.getBlob();
          const fileSize = blob.getBytes().length;

          if (totalSizeBytes + fileSize > CONFIG.MAX_ATTACHMENTS_BYTES) {
            if (!warningShown) {
              SpreadsheetApp.getUi().alert(
                `Warning: Claim ${referenceCode} exceeds 25MB. It will be sent without some attachments.`
              );
              warningShown = true;
            }
            return;
          }

          attachments.push(blob);
          totalSizeBytes += fileSize;
        } catch (e) {
          console.log(`Drive error for file ID (${fileId}): ${e.message}`);
        }
      });
    });
  });

  return attachments;
}

function buildClaimEmailHtml(payload) {
  const tableStyle = 'border-collapse: collapse; width: 100%; max-width: 500px;';
  const tdLabel = 'padding: 6px 10px; border: 1px solid #ccc; background-color: #f5f5f5; font-weight: bold; width: 40%;';
  const tdValue = 'padding: 6px 10px; border: 1px solid #ccc;';

  // Build CC line for display in email body
  const ccEmails = ['rh.finance@u.nus.edu'];
  if (payload.otherEmails) {
    payload.otherEmails.split(',').map(e => e.trim()).filter(Boolean).forEach(e => ccEmails.push(e));
  }

  return `
    <div style="font-family: Arial, sans-serif; color: #222; font-size: 14px; line-height: 1.6;">

      <p>Hi ${payload.name.split(' ')[0]},</p>

      <p>We have received your claim and after sending the following email, this is a confirmation that your claim is being processed.</p>

      <p>Please copy and paste everything below the line into a new email, and attach all the attachments from this email:</p>

      <p>
        <strong>To:</strong> rh.finance@u.nus.edu<br>
        ${ccEmails.length > 1 ? `<strong>CC:</strong> ${ccEmails.slice(1).join(', ')}<br>` : ''}
        <strong>Subject:</strong> ${payload.referenceCode}
      </p>

      <hr style="border: none; border-top: 2px solid #000; margin: 20px 0;">

      <p>Dear Ryan,</p>
      <p>Attached is the claims for ${payload.event}.</p>
      <p>To whom it may concern,</p>
      <p>I, ${payload.name.toUpperCase()}, ${payload.matric.toUpperCase()}, hereby authorise my treasurer, Ryan, to collect reimbursement on my behalf.</p>

      <strong>Claims Summary</strong><br>
      <table style="${tableStyle}">
        <tr><td style="${tdLabel}">CCA</td><td style="${tdValue}">${payload.ccaName}</td></tr>
        <tr><td style="${tdLabel}">Event</td><td style="${tdValue}">${payload.event.toUpperCase()}</td></tr>
        <tr><td style="${tdLabel}">CCA Treasurer</td><td style="${tdValue}">${payload.name.toUpperCase()}</td></tr>
        <tr><td style="${tdLabel}">Phone Number for PayNow</td><td style="${tdValue}">${payload.phone}</td></tr>
        <tr><td style="${tdLabel}">Total collated amount</td><td style="${tdValue}">$${payload.totalFormatted}</td></tr>
      </table>

      <p><strong>Purpose of Purchase:</strong><br>
      ${payload.receiptListHtml || 'No individual receipts found.'}</p>

      ${payload.remarks ? `<p><strong>Remarks:</strong><br>${payload.remarks}</p>` : ''}

      <p>Thank you.</p>

    </div>`;
}

function resetAddClaimSheet(spreadsheet, templateSheet) {
  const oldSheet = spreadsheet.getSheetByName(CONFIG.SHEETS.ADD_CLAIM);
  if (oldSheet) {
    spreadsheet.deleteSheet(oldSheet);
  }
  const newSheet = templateSheet.copyTo(spreadsheet);
  newSheet.setName(CONFIG.SHEETS.ADD_CLAIM);
  spreadsheet.setActiveSheet(newSheet);
  spreadsheet.moveActiveSheet(3);
}

// ============================================================================
// PART 4: CORE WORKFLOW FUNCTIONS (UNCHANGED)
// ============================================================================

function addClaim() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = getSheetOrThrow(spreadsheet, CONFIG.SHEETS.TEMPLATE);
  const addClaimSheet = getSheetOrThrow(spreadsheet, CONFIG.SHEETS.ADD_CLAIM);
  const claimsDataSheet = getSheetOrThrow(spreadsheet, CONFIG.SHEETS.CLAIMS_DATA);
  const masterSheet = getSheetOrThrow(spreadsheet, CONFIG.SHEETS.MASTER);
  const formResponsesSheet = getSheetOrThrow(spreadsheet, CONFIG.SHEETS.FORM_RESPONSES);

  const receiptRanges = CONFIG.ADD_CLAIM.RECEIPT_COLUMNS.flatMap(column =>
    buildColumnRangeList(column, CONFIG.ADD_CLAIM.RECEIPT_ROWS.START, CONFIG.ADD_CLAIM.RECEIPT_ROWS.END)
  );

  const rowClaimsData = [
    claimsDataSheet.getLastRow(),
    ...getCellValues(addClaimSheet, CONFIG.ADD_CLAIM.CLAIM_RANGES),
    ...getCellValues(addClaimSheet, receiptRanges)
  ];

  const rowMasterData = [
    masterSheet.getLastRow(),
    ...getCellValues(addClaimSheet, CONFIG.ADD_CLAIM.MASTER_RANGES)
  ];

  claimsDataSheet.appendRow(rowClaimsData);
  const newClaimsRow = claimsDataSheet.getLastRow();
  claimsDataSheet.getRange(newClaimsRow, 1, 1, claimsDataSheet.getLastColumn())
    .setBackground(null)
    .setFontColor(null)
    .setFontWeight('normal')
    .setHorizontalAlignment('center')
    .setWrap(false);
  masterSheet.appendRow(rowMasterData);
  // Clear header formatting that bleeds into new rows
  const newMasterRow = rowMasterData[0] + 1;
  masterSheet.getRange(newMasterRow, 1, 1, masterSheet.getLastColumn())
    .setBackground(null)
    .setFontColor(null)
    .setFontWeight('normal')
    .setHorizontalAlignment('center')
    .setWrap(false);
  masterSheet.getRange(rowMasterData[0] + 1, 10, 1, 5).insertCheckboxes();
  masterSheet.getRange(rowMasterData[0] + 1, 16, 1, 2).insertCheckboxes();

  resetAddClaimSheet(spreadsheet, templateSheet);

  spreadsheet.setActiveSheet(formResponsesSheet);

  const ui = SpreadsheetApp.getUi();
  ui.alert('Successfully added claim');
}

function generateEmail() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const claimsDataSheet = getSheetOrThrow(spreadsheet, CONFIG.SHEETS.CLAIMS_DATA);
  const masterSheet = getSheetOrThrow(spreadsheet, CONFIG.SHEETS.MASTER);
  const MASTER_COL = CONFIG.MASTER_COL_EMAIL;

  const masterData = masterSheet.getDataRange().getValues();
  const pendingRows = getPendingRowIndices(masterData, MASTER_COL.GENERATE_EMAIL, MASTER_COL.EMAILS_SENT);

  if (pendingRows.length === 0) {
    SpreadsheetApp.getUi().alert('No rows marked for email generation in Master Sheet.');
    return;
  }

  if (!confirmBatchAction(
    'Confirm Batch Send',
    `You are about to send ${pendingRows.length} email(s). Are you sure you want to proceed?`
  )) {
    return;
  }

  const claimsData = claimsDataSheet.getDataRange().getValues();
  let sentCount = 0;

  pendingRows.forEach(rowIndex => {
    const masterRow = masterData[rowIndex];
    const claimNo = masterRow[MASTER_COL.NO];

    const row = findClaimRow(claimsData, claimNo);
    if (!row) {
      console.log("No matching data found for Claim No: " + claimNo);
      return;
    }

    const payload = {
      recipientEmail: row[CONFIG.CLAIMS_COL.EMAIL],
      name: (row[CONFIG.CLAIMS_COL.CLAIMER_NAME] || '').toString(),
      matric: (row[CONFIG.CLAIMS_COL.CLAIMER_MATRIC] || '').toString(),
      phone: row[CONFIG.CLAIMS_COL.CLAIMER_PHONE],
      event: (row[CONFIG.CLAIMS_COL.CLAIM_DESC] || '').toString(),
      ccaName: (row[CONFIG.CLAIMS_COL.CCA] || '').toString(),
      remarks: (row[CONFIG.CLAIMS_COL.REMARKS] || '').toString().replace(/\n/g, '<br>'),
      totalFormatted: formatMoney(row[CONFIG.CLAIMS_COL.TOTAL]),
      referenceCode: row[CONFIG.CLAIMS_COL.REFERENCE_CODE],
      otherEmails: (row[CONFIG.CLAIMS_COL.OTHER_EMAILS] || '').toString().trim()
    };

    const receiptListHtml = buildReceiptListHtml(row);
    const attachments = collectReceiptAttachments(row, payload.referenceCode);
    const htmlTemplate = buildClaimEmailHtml({ ...payload, receiptListHtml });

    // Build CC list: other emails involved + rh.finance@u.nus.edu
    const ccList = ['rh.finance@u.nus.edu'];
    if (payload.otherEmails) {
      payload.otherEmails.split(',').map(e => e.trim()).filter(Boolean).forEach(e => ccList.push(e));
    }

    // Send to claimer with confirmation message
    GmailApp.sendEmail(payload.recipientEmail, payload.referenceCode, "", {
      htmlBody: htmlTemplate,
      attachments: attachments,
      cc: ccList.join(', ')
    });

    markMasterStatus(masterSheet, rowIndex, MASTER_COL.EMAILS_SENT);
    sentCount++;
  });

  if (sentCount > 0) {
    SpreadsheetApp.getUi().alert(`Success: ${sentCount} email(s) sent.`);
  }
}

function generateForm() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const claimsDataSheet = getSheetOrThrow(spreadsheet, CONFIG.SHEETS.CLAIMS_DATA);
  const masterSheet = getSheetOrThrow(spreadsheet, CONFIG.SHEETS.MASTER);
  const MASTER_COL = CONFIG.MASTER_COL_FORM;

  const masterData = masterSheet.getDataRange().getValues();
  const pendingRows = getPendingRowIndices(masterData, MASTER_COL.GENERATE_FORM, MASTER_COL.FORMS_GENERATED);

  if (pendingRows.length === 0) {
    SpreadsheetApp.getUi().alert('No rows marked for form generation in Master Sheet.');
    return;
  }

  if (!confirmBatchAction(
    'Confirm Batch Generation',
    `You are about to generate forms for ${pendingRows.length} claim(s). Are you sure you want to proceed?`
  )) {
    return;
  }

  const claimsData = claimsDataSheet.getDataRange().getValues();
  const mainFolder = DriveApp.getFolderById(
    getOrCreateFolder(CONFIG.FOLDERS.RFP_ROOT_NAME, DriveApp.getFolderById(CONFIG.FOLDERS.RFP_ROOT_ID))
  );

  let generatedCount = 0;

  pendingRows.forEach(rowIndex => {
    const masterRow = masterData[rowIndex];
    const claimNo = masterRow[MASTER_COL.NO];

    const row = findClaimRow(claimsData, claimNo);
    if (!row) {
      console.log("No matching data found for Claim No: " + claimNo);
      return;
    }

    const referenceCode = (row[CONFIG.CLAIMS_COL.REFERENCE_CODE] || "").toString();
    if (!referenceCode) {
      console.log("No reference code found for Claim No: " + claimNo);
      return;
    }

    const subfolderName = `Claim No. ${claimNo} - (${referenceCode})`;
    const claimFolder = DriveApp.getFolderById(getOrCreateFolder(subfolderName, mainFolder));

    try {
      generateLOA(claimNo, row, claimFolder);
      generateSummary(claimNo, row, claimFolder);
      generateRFP(claimNo, row, claimFolder);

      markMasterStatus(masterSheet, rowIndex, MASTER_COL.FORMS_GENERATED);
      generatedCount++;

      console.log(`Successfully generated forms for ${referenceCode}`);
    } catch (e) {
      console.error(`Error generating forms for ${referenceCode}: ${e.message}`);
      SpreadsheetApp.getUi().alert(`Error generating forms for ${referenceCode}: ${e.message}`);
    }
  });

  if (generatedCount > 0) {
    SpreadsheetApp.getUi().alert(`Success: Generated forms for ${generatedCount} claim(s).`);
  }
}

async function compileForms() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const claimsDataSheet = getSheetOrThrow(spreadsheet, CONFIG.SHEETS.CLAIMS_DATA);
  const masterSheet = getSheetOrThrow(spreadsheet, CONFIG.SHEETS.MASTER);
  const MASTER_COL = CONFIG.MASTER_COL_COMPILE;
  const FORM_COL = CONFIG.MASTER_COL_FORM;

  const masterData = masterSheet.getDataRange().getValues();
  const claimsData = claimsDataSheet.getDataRange().getValues();
  const rowsToCompile = getPendingRowIndices(masterData, MASTER_COL.GENERATE_COMPILE, MASTER_COL.COMPILED);

  if (rowsToCompile.length === 0) {
    SpreadsheetApp.getUi().alert('No rows marked for compile in Master Sheet.');
    return;
  }

  if (!confirmBatchAction(
    'Confirm Compile',
    `You are about to compile ${rowsToCompile.length} claim(s) into PDF. Proceed?`
  )) {
    return;
  }

  const mainFolderId = getOrCreateFolder(CONFIG.FOLDERS.RFP_ROOT_NAME, DriveApp.getFolderById(CONFIG.FOLDERS.RFP_ROOT_ID));
  const mainFolder = DriveApp.getFolderById(mainFolderId);

  let compiledCount = 0;
  let skippedNotGenerated = 0;
  let skippedMissingData = 0;
  let skippedMissingReference = 0;
  let skippedMissingFiles = 0;
  let errorCount = 0;
  const errorMessages = [];

  for (const rowIndex of rowsToCompile) {
    const claimNo = masterData[rowIndex][MASTER_COL.NO];

    if (masterData[rowIndex][FORM_COL.FORMS_GENERATED] !== true) {
      skippedNotGenerated++;
      continue;
    }

    const row = findClaimRow(claimsData, claimNo);
    if (!row) { skippedMissingData++; continue; }

    const referenceCode = (row[CONFIG.CLAIMS_COL.REFERENCE_CODE] || '').toString();
    if (!referenceCode) { skippedMissingReference++; continue; }

    const claimFolderName = `Claim No. ${claimNo} - (${referenceCode})`;
    const claimFolderId = getOrCreateFolder(claimFolderName, mainFolder);
    const claimFolder = DriveApp.getFolderById(claimFolderId);

    const loaName = `LOA No. ${claimNo} - (${referenceCode})`;
    const summaryName = `Summary No. ${claimNo} - (${referenceCode})`;
    const rfpName = `RFP No. ${claimNo} - (${referenceCode})`;

    const loaFile = findFileByNameInFolder(claimFolder, loaName);
    const summaryFile = findFileByNameInFolder(claimFolder, summaryName);
    const rfpFile = findFileByNameInFolder(claimFolder, rfpName);

    if (!loaFile || !summaryFile || !rfpFile) {
      skippedMissingFiles++;
      continue;
    }

    const pdfBlobs = [
      exportToPdfBlob(rfpFile, `${rfpName}.pdf`),
      exportToPdfBlob(loaFile, `${loaName}.pdf`),
      exportToPdfBlob(summaryFile, `${summaryName}.pdf`)
    ];

    const compiledName = `${CONFIG.PDF.COMPILED_PREFIX} ${claimNo} - (${referenceCode}).pdf`;
    deleteExistingFileByName(claimFolder, compiledName);

    try {
      const mergedBlob = await mergePdfBlobs(pdfBlobs, compiledName);
      claimFolder.createFile(mergedBlob);
      compiledCount++;
      markMasterStatus(masterSheet, rowIndex, MASTER_COL.COMPILED);
    } catch (e) {
      errorCount++;
      const message = `Failed to compile PDF for ${referenceCode}: ${e.message}`;
      errorMessages.push(message);
      console.error(message);
    }
  }

  if (compiledCount > 0) {
    SpreadsheetApp.getUi().alert(`Success: Compiled ${compiledCount} PDF(s).`);
  } else {
    const errorDetails = errorMessages.length ? `\n\nErrors:\n${errorMessages.join('\n')}` : '';
    SpreadsheetApp.getUi().alert(
      `No PDFs compiled.\n` +
      `Not generated: ${skippedNotGenerated}\n` +
      `Missing data: ${skippedMissingData}\n` +
      `Missing reference code: ${skippedMissingReference}\n` +
      `Missing files: ${skippedMissingFiles}\n` +
      `Errors: ${errorCount}` +
      errorDetails
    );
  }
}

function getOrCreateFolder(folderName, parentFolder) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next().getId();
  } else {
    return parentFolder.createFolder(folderName).getId();
  }
}

// ============================================================================
// PART 5: DOCUMENT GENERATION FUNCTIONS (UNCHANGED)
// ============================================================================

function generateLOA(claimNo, row, folder) {
  const referenceCode = (row[CONFIG.CLAIMS_COL.REFERENCE_CODE] || "").toString();

  const doc = DocumentApp.create(`LOA No. ${claimNo} - (${referenceCode})`);
  const body = doc.getBody();

  body.appendPageBreak();

  let isFirstImage = true;

  const addFileToDoc = (fileId) => {
    try {
      const file = DriveApp.getFileById(fileId);
      const mimeType = file.getMimeType();
      
      if (mimeType === 'application/pdf') {
        let convertedFileId = null;
        let conversionSuccessful = false;
        let contentElements = [];
        
        try {
          const pdfBlob = file.getBlob();
          
          const fileMetadata = {
            name: file.getName() + '_temp_' + new Date().getTime(),
            mimeType: 'application/vnd.google-apps.document',
            parents: [folder.getId()]
          };
          
          const convertedFile = Drive.Files.create(fileMetadata, pdfBlob, {
            supportsAllDrives: true,
            fields: 'id'
          });
          
          if (convertedFile && convertedFile.id) {
            convertedFileId = convertedFile.id;
            
            const tempDoc = DocumentApp.openById(convertedFile.id);
            const tempBody = tempDoc.getBody();
            const numChildren = tempBody.getNumChildren();
            
            let hasContent = false;
            
            for (let i = 0; i < numChildren; i++) {
              const element = tempBody.getChild(i);
              const elementType = element.getType();
              
              if (elementType === DocumentApp.ElementType.PARAGRAPH) {
                const para = element.asParagraph();
                if (para.getText().trim().length > 0 || para.getNumChildren() > 0) {
                  hasContent = true;
                  contentElements.push(element.copy());
                }
              } else if (elementType === DocumentApp.ElementType.TABLE ||
                         elementType === DocumentApp.ElementType.LIST_ITEM ||
                         elementType === DocumentApp.ElementType.INLINE_IMAGE) {
                hasContent = true;
                contentElements.push(element.copy());
              }
            }
            
            if (hasContent && contentElements.length > 0) {
              conversionSuccessful = true;
            }
          }
          
        } catch (convertError) {
          console.log(`PDF conversion error for ${file.getName()}: ${convertError.message}`);
          conversionSuccessful = false;
        } finally {
          if (convertedFileId) {
            try {
              DriveApp.getFileById(convertedFileId).setTrashed(true);
            } catch (deleteError) {
              console.log(`Could not delete temp file: ${deleteError.message}`);
            }
          }
        }
        
        if (conversionSuccessful && contentElements.length > 0) {
          if (!isFirstImage) body.appendPageBreak();
          isFirstImage = false;
          
          contentElements.forEach(element => {
            const elementType = element.getType();
            if (elementType === DocumentApp.ElementType.PARAGRAPH) {
              body.appendParagraph(element.asParagraph().copy());
            } else if (elementType === DocumentApp.ElementType.TABLE) {
              body.appendTable(element.asTable().copy());
            } else if (elementType === DocumentApp.ElementType.LIST_ITEM) {
              body.appendListItem(element.asListItem().copy());
            }
          });
        } else {
          try {
            file.makeCopy(file.getName(), folder);
          } catch (copyError) {
            console.log(`Could not copy PDF: ${copyError.message}`);
          }
        }
        
      } else {
        if (!isFirstImage) body.appendPageBreak();
        isFirstImage = false;
        
        const blob = file.getBlob();
        const image = body.appendImage(blob);
        
        const width = image.getWidth();
        const height = image.getHeight();
        if (width > 550) {
          const ratio = 550 / width;
          image.setWidth(550);
          image.setHeight(height * ratio);
        }
      }
    } catch (e) {
      console.log(`Error adding file ${fileId}: ${e.message}`);
    }
  };

  forEachReceipt(row, receipt => {
    if (!receipt.desc || receipt.amount === '') return;
    extractDriveFileIds(receipt.imgUrlCell).forEach(addFileToDoc);
    extractDriveFileIds(receipt.bankUrlCell).forEach(addFileToDoc);
  });

  if (!isFirstImage) {
    const paragraphs = body.getParagraphs();
    if (paragraphs.length > 0 && paragraphs[0].getText() === '') {
      paragraphs[0].removeFromParent();
    }
  }

  doc.saveAndClose();
  
  const docFile = DriveApp.getFileById(doc.getId());
  docFile.moveTo(folder);
}

function generateSummary(claimNo, row, folder) {
  const referenceCode = (row[CONFIG.CLAIMS_COL.REFERENCE_CODE] || "").toString();
  const financeDirectorName = (row[CONFIG.CLAIMS_COL.FINANCE_D_NAME] || "").toString();
  const financeDirectorMatric = (row[CONFIG.CLAIMS_COL.FINANCE_D_MATRIC] || "").toString();
  const financeDirectorPhone = row[CONFIG.CLAIMS_COL.FINANCE_D_PHONE];
  const claimDesc = (row[CONFIG.CLAIMS_COL.CLAIM_DESC] || "").toString();
  const total = row[CONFIG.CLAIMS_COL.TOTAL];
  const totalFormatted = formatMoney(total);
  const date = row[CONFIG.CLAIMS_COL.DATE];
  const formattedDate = formatDateIfNeeded(date);
  const wbsNo = (row[CONFIG.CLAIMS_COL.WBS_NO] || "").toString();
  
  const templateFile = DriveApp.getFileById(CONFIG.TEMPLATE_IDS.SUMMARY);
  const copiedFile = templateFile.makeCopy(`Summary No. ${claimNo} - (${referenceCode})`, folder);
  const ss = SpreadsheetApp.openById(copiedFile.getId());
  const sheet = ss.getActiveSheet();
  
  sheet.getRange('B6').setValue(referenceCode);
  sheet.getRange('B8').setValue(formattedDate);
  sheet.getRange('B12').setValue(claimDesc);
  sheet.getRange('C20').setValue(financeDirectorName);
  sheet.getRange('I20').setValue(financeDirectorMatric);
  sheet.getRange('C24').setValue(totalFormatted);
  sheet.getRange('I24').setValue(financeDirectorPhone);
  sheet.getRange('B36').setValue(wbsNo);
  
  let currentRow = 31;
  let receiptNum = 1;
  
  forEachReceipt(row, receipt => {
    if (!receipt.desc || receipt.amount === "") return;

    const formattedAmt = formatMoney(receipt.amount);
    const formattedDate = formatDateIfNeeded(receipt.date);
    
    sheet.getRange(currentRow, 1).setValue(receiptNum || "");
    sheet.getRange(currentRow, 2).setValue(receipt.receiptNo || "");
    sheet.getRange(currentRow, 4).setValue(receipt.desc || "");
    sheet.getRange(currentRow, 9).setValue(formattedAmt || "");
    sheet.getRange(currentRow, 10).setValue(receipt.categoryCode || "");
    sheet.getRange(currentRow, 11).setValue(receipt.gstCode || "");
    
    currentRow++;
    receiptNum++;
  });
  
  sheet.getRange(36, 9).setValue(totalFormatted);
  SpreadsheetApp.flush();
}

function generateRFP(claimNo, row, folder) {
  const referenceCode = (row[CONFIG.CLAIMS_COL.REFERENCE_CODE] || "").toString();
  const financeDirectorName = (row[CONFIG.CLAIMS_COL.FINANCE_D_NAME] || "").toString();
  const financeDirectorMatric = (row[CONFIG.CLAIMS_COL.FINANCE_D_MATRIC] || "").toString();
  const total = row[CONFIG.CLAIMS_COL.TOTAL];
  const totalFormatted = formatMoney(total);
  const wbsNo = (row[CONFIG.CLAIMS_COL.WBS_NO] || "").toString();
  
  const templateFile = DriveApp.getFileById(CONFIG.TEMPLATE_IDS.RFP);
  const copiedFile = templateFile.makeCopy(`RFP No. ${claimNo} - (${referenceCode})`, folder);
  const doc = DocumentApp.openById(copiedFile.getId());
  const body = doc.getBody();
  
  body.replaceText('{{NAME}}', financeDirectorName.toUpperCase());
  body.replaceText('{{MATRIC}}', financeDirectorMatric.toUpperCase());
  body.replaceText('{{TOTAL_AMOUNT}}', totalFormatted);
  body.replaceText('{{REFERENCE_CODE}}', referenceCode);
  body.replaceText('{{WBS_NO}}', wbsNo);

  forEachReceipt(row, receipt => {
    const index = receipt.index;
    const amt = receipt.amount;

    if (amt !== "" && amt !== null && !isNaN(amt)) {
      const dollars = Math.floor(amt);
      const cents = Math.round((amt - dollars) * 100).toString().padStart(2, '0');

      body.replaceText(`{{DR_CR_${index}}}`, receipt.drCr || "DR");
      body.replaceText(`{{GL_${index}}}`, receipt.categoryCode || "");
      body.replaceText(`{{DOLLAR_${index}}}`, dollars.toString());
      body.replaceText(`{{CENTS_${index}}}`, cents);
      body.replaceText(`{{GST_${index}}}`, receipt.gstCode || "");
      body.replaceText(`{{WBS_${index}}}`, wbsNo || "");
    } else {
      body.replaceText(`{{DR_CR_${index}}}`, "");
      body.replaceText(`{{GL_${index}}}`, "");
      body.replaceText(`{{DOLLAR_${index}}}`, "");
      body.replaceText(`{{CENTS_${index}}}`, "");
      body.replaceText(`{{GST_${index}}}`, "");
      body.replaceText(`{{WBS_${index}}}`, "");
    }
  });
  
  doc.saveAndClose();
}

function getConfigValue(key) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  if (!configSheet) return null;
  
  const data = configSheet.getRange('A:B').getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      return data[i][1];
    }
  }
  return null;
}

// ============================================================================
// PART 6: CONFIGURATION CONSTANTS
// ============================================================================

/**
 * Portfolio → CCA mapping.
 * Edit this object OR edit the "Form Options" sheet directly
 * and use "Recreate Claims Form" from the menu to update the form.
 */
const CCA_DEPARTMENTS = {
  "Culture": [
    "Culture", "RHockerfellas", "RH Unplugged", "RH Dance",
    "RHebels", "RHythm", "RH Voices", "Culture Comm"
  ],
  "Welfare": [
    "Welfare Comm", "RVC SP", "RVC Children", "RVC Pioneers",
    "RVC Special Needs", "HeaRHtfelt", "Green Comm"
  ],
  "Sports": [
    "Badminton M", "Basketball M", "Floorball M", "Handball M",
    "Soccer M", "Swimming M", "Squash M", "Sepak Takraw M",
    "Tennis M", "Touch Rugby M", "Table Tennis M", "Volleyball M",
    "SMC", "Softball", "Track", "Road Relay", "Frisbee",
    "Netball F", "Badminton F", "Basketball F", "Floorball F",
    "Handball F", "Soccer F", "Swimming F", "Squash F",
    "Tennis F", "Touch Rugby F", "Table Tennis F", "Volleyball F"
  ],
  "Social": [
    "Block 2 Comm", "Block 3 Comm", "Block 4 Comm", "Block 5 Comm",
    "Block 6 Comm", "Block 7 Comm", "Block 8 Comm", "Social Comm", "RHSafe"
  ],
  "VPI": ["Bash", "DND", "AEAC"],
  "RHMP": [
    "RHMP Producers", "RHMP Directors", "RHMP Ensemble", "RHMP Stage Managers",
    "RHMP Sets", "RHMP Costumes", "RHMP Relations", "RHMP Publicity", "RHMP EM",
    "RHMP Graphic Design", "RHMP Musicians", "RHMP Composers"
  ],
  "Media": [
    "BOP", "Phoenix Studios", "Phoenix Press", "AnG",
    "Tech Crew", "ComMotion", "RH Devs"
  ],
  "HGS": ["Vacation Storage", "Auditor", "Finance", "Secretariat"],
  "VPE": ["HPB", "RHOC", "RHAG", "RFLAG"]
};

const CONFIG = {
  MENU: {
    TITLE: 'Claims Tools',
    ITEMS: [
      { label: 'Add Claim', handler: 'addClaim' },
      { label: 'Generate Emails', handler: 'generateEmail' },
      { label: 'Generate Forms', handler: 'generateForm' },
      { label: 'Compile Forms', handler: 'compileForms' },
    ]
  },
  SHEETS: {
    TEMPLATE: 'Claim Data Template',
    ADD_CLAIM: 'Add Claim',
    CLAIMS_DATA: 'Claims Data',
    MASTER: 'Master Sheet',
    FORM_RESPONSES: 'Form Responses'
  },
  ADD_CLAIM: {
    CLAIM_RANGES: [
      'B8','B9','B10','B11','B12','B13','B14','B15','B16','B17','B18','B19','B20','B21','B22','B23','B24'
    ],
    RECEIPT_COLUMNS: ['E','G','I','K','M'],
    RECEIPT_ROWS: { START: 8, END: 18 },
    MASTER_RANGES: ['B15','B16','B17','B18','B11','B13','B20','B26']
  },
  MASTER_COL_EMAIL: {
    NO: 0,
    GENERATE_EMAIL: 9,
    EMAILS_SENT: 10
  },
  MASTER_COL_FORM: {
    NO: 0,
    GENERATE_FORM: 11,
    FORMS_GENERATED: 12
  },
  MASTER_COL_COMPILE: {
    NO: 0,
    GENERATE_COMPILE: 15,
    COMPILED: 16
  },
  CLAIMS_COL: {
    NO: 0,
    FINANCE_D_NAME: 1,
    FINANCE_D_MATRIC: 2,
    FINANCE_D_PHONE: 3,
    CLAIMER_NAME: 4,
    CLAIMER_MATRIC: 5,
    CLAIMER_PHONE: 6,
    EMAIL: 7,
    PORTFOLIO: 8,
    CCA: 9,
    CLAIM_DESC: 10,
    TOTAL: 11,
    DATE: 12,
    REFERENCE_CODE: 13,
    WBS_ACCOUNT_NAME: 14,
    WBS_NO: 15,
    REMARKS: 16,
    OTHER_EMAILS: 17,
    DESC_START: 18
  },
  RECEIPT: {
    COUNT: 5,
    BLOCK_WIDTH: 11
  },
  get TEMPLATE_IDS() {
    return {
      SUMMARY: getConfigValue('SUMMARY_TEMPLATE_ID') || '',
      RFP: getConfigValue('RFP_TEMPLATE_ID') || ''
    };
  },
  get FOLDERS() {
    return {
      RFP_ROOT_ID: getConfigValue('RFP_FOLDER_ID') || '',
      RFP_ROOT_NAME: 'RFPs'
    };
  },
  PDF: {
    LIB_URL: 'https://unpkg.com/pdf-lib@1.17.1/dist/pdf-lib.min.js',
    COMPILED_PREFIX: 'Compiled No.'
  },
  TIMEZONE: 'GMT+8',
  DATE_FORMAT: 'dd/MM/yyyy',
  MAX_ATTACHMENTS_BYTES: 25 * 1024 * 1024
};
