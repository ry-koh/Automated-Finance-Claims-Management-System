/**
 * ====================================================================================
 * CLAIMS MANAGEMENT SYSTEM
 * ====================================================================================
 * 
 * INSTALLATION (One-Time Setup):
 * ------------------------------
 * 1. Create a brand new Google Sheet
 * 2. Extensions > Apps Script
 * 3. Delete all default code
 * 4. Paste THIS ENTIRE FILE
 * 5. Save (Ctrl/Cmd + S)
 * 6. Click Run > select "setupClaimsSystem"
 * 7. Grant permissions when prompted
 * 8. Fill in the Config sheet that appears
 * 9. Enable Drive API: Click + next to Services > Add "Drive API" v3
 * 
 * 10. Move the Google Sheets into the Automated Claims folder
 * 11. Copy the Template RFP and Summary into the Automated Claims folder
 * 12. Set them to 'Anyone with the link can view'
 * 13. Copy their links and extract the ID and input into the Config sheet
 * Example 1: https://docs.google.com/spreadsheets/d/1LDyJGAGqTKjra-2ErZ885EEU2N-wVJetvL82kNQYS9Y/edit?usp=drive_link
 * ID: 1LDyJGAGqTKjra-2ErZ885EEU2N-wVJetvL82kNQYS9Y1LDyJGAGqTKjra-2ErZ885EEU2N-wVJetvL82kNQYS9Y
 * Example 2: https://docs.google.com/document/d/1qvpOijRMO5chlJIsCYLLpvVVmLrawgHPn0c7uRhPBUo/edit?usp=drive_link
 * ID: 1qvpOijRMO5chlJIsCYLLpvVVmLrawgHPn0c7uRhPBUo
 * 
 * 14. Hide the Config sheet
 * 15. Copy the Finance Form into the Automated Claims folder
 * 16. Link the Finance Form to the Google Sheets (Link to existing spreadsheet and select the Google Sheets created)
 * 
 * 17. Scroll to the bottom of the Form Responses 1 Sheet and add 3000 more rows
 * 18. Add 2 columns to the left of the Form Responses 1 Sheet
 * 19. Select the 2 columns except the headers and click Insert -> Checkboxes
 * 20. Rename the headers Processed and Error
 * 
 * 21. You are done :D
 *
 * ====================================================================================
 */

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

    let originalSheet = ss.getSheetByName('Sheet1');
    ss.deleteSheet(originalSheet);
    
    // Create folder structure
    const config = loadConfigFromSheet(ss);
    createFolderStructure(ss, config);
    
    ui.alert(
      'Setup Complete!',
      ui.ButtonSet.OK
    );
    
    ss.setActiveSheet(ss.getSheetByName('Config'));
    
  } catch (e) {
    ui.alert('Error', `Setup failed: ${e.message}`, ui.ButtonSet.OK);
    console.error('Setup error:', e);
  }
}

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
    ['RFP_FOLDER_ID', '', 'RFPs subfolder']
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

  sheet.deleteRows(14, 987);
  sheet.deleteColumns(4, 23);

  // Add conditional formatting for required fields
  const requiredRange1 = sheet.getRange('B2:B5');  // Academic year and Finance D info
  const requiredRange2 = sheet.getRange('B8:B9');  // Template IDs
  
  const emptyRule = SpreadsheetApp.newConditionalFormatRule()
    .whenCellEmpty()
    .setBackground('#f4cccc')  // Light red
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
    'Forms Generated', 'Email Screenshot Added', ' Formatting Remarks', 'Compile Forms',
    'Compiled', 'EMAIL SUBMISSION DATE', 'PROCESSED TO FINANCE DIRECTOR', 'SUBMISSION TO OFFICE',
    'DATE OF REIMBURSEMENT', 'STATUS', 'REMARKS (For my own use)', 'FOR RL USE ONLY'
  ];
  
  sheet.getRange(1, 1, 1, headers.length)
    .setValues([headers])
    .setHorizontalAlignment('center')
    .setWrap(true);
  
  sheet.setColumnWidth(1, 50);
  sheet.deleteRows(2, 999);

  const protection = sheet.protect().setDescription('Protected with exceptions');
  protection.setWarningOnly(true);
  const ranges = sheet.getRangeList(['J:J', 'L:L', 'N:P']).getRanges();
  protection.setUnprotectedRanges(ranges);

  sheet.getRangeList(["E:E"]).setNumberFormat("$#,##0.00");
  
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
    'Reference Code', 'WBS Account Name', 'WBS No.', 'Remarks'
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
    .setHorizontalAlignment('center');
  
  sheet.setColumnWidth(1, 50);
  sheet.deleteRows(2, 999);

  const protection = sheet.protect().setDescription('Protected with exceptions');
  protection.setWarningOnly(true);
  
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
  
  // Claim info section header
  sheet.getRange('A7:B7').merge()
    .setValue('CLAIM INFORMATION')
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setFontWeight('bold');
  
  // Manual input fields start at row 8
  const startRow = 8;
  const fields = [
    ['Finance D Name', '=IF($B$14<>"",\'Config\'!$B$3,"")'],
    ['Finance D Matric No.', '=IF($B$14<>"",\'Config\'!$B$4,"")'],
    ['Finance D Phone No.', '=IF($B$14<>"",\'Config\'!$B$5,"")'],
    ['Claimer Name', '=IF($B$14<>"",IF(H5="",IFNA(VLOOKUP(B14,\'Identifier Data\'!$A:$F,4,FALSE),"Name not in list"),H5),"")'],
    ['Claimer Matric No.', '=IF($B$14<>"",IF(I5="",IFNA(VLOOKUP(B14,\'Identifier Data\'!$A:$F,5,FALSE),"Name not in list"),I5),"")'],
    ['Claimer Phone No.', '=IF($B$14<>"",IF(J5="",IFNA(VLOOKUP(B14,\'Identifier Data\'!$A:$F,6,FALSE),"Name not in list"),J5),"")'],
    ['Email Address', '=IF(ISBLANK($B$5),"",B5)'],
    ['Portfolio', '=IF(ISBLANK($B$14),"",C5)'],
    ['CCA', '=IF(ISBLANK($B$14),"",D5)'],
    ['Claim Description', '=IF(ISBLANK($B$14), "",E5)'],
    ['Total Claim Amount', '=IF($B$14<>"",SUM(E16,G16,I16,K16,M16),"")'],  // Sum receipt amounts in row 16
    ['Date', '=IF($B$14<>"",TEXT(TODAY(),"DD/MM/YYYY"),"")'],  // From timestamp
    ['Reference Code', '=IF($B$14<>"",CONCATENATE(Config!$B$2,"-",UPPER(TEXT(TODAY(),"MMM")),"-",UPPER($B$15),"-",UPPER($B$16),"-",UPPER($B$17)),"")'],
    ['WBS Account Name', ''],  // Manual entry
    ['WBS No.', '=IF($B$21<>"",SWITCH($B$21, "Student Activity Fund", "H-404-00-000003", "Managed by Hall Fund", "H-404-00-000004", "Master Fund", "E-404-10-0001-01", "Master Fund (RHMP)", "E-404-10-0001-07"),"")'],  // Manual entry or you can add SWITCH formula if needed
    ['Remarks', '=IF(ISBLANK(F5),"",F5)'],
    ['', ''],
    ['WBS Account Short Form', '=IF($B$21<>"",SWITCH($B$21, "Student Activity Fund", "SA", "Managed by Hall Fund", "MBH", "Master Fund", "MF", "Master Fund (RHMP)", "MF (RHMP)"),"")']
  ];

  // Add dropdown validation for WBS Account Name (cell B21)
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

  sheet.getRange('B21').setDataValidation(wbsRule);

  // Category options for dropdown
  const categoryOptions = [
    'Office Supplies',
    'Consumables',
    'Sports & Cultural Materials',
    'Other fees (Others)',
    'Professional fees',
    'Bank Charges',
    'Licensing/Subscription',
    'Postage & Telecommunication Charges',
    'Maintenance (Equipment)',
    'Lease expense (premises)',
    'Lease expense (rental of equipment)',
    'Furniture',
    'Equipment Purchase',
    'Publications',
    'Meals & Refreshments',
    'Local Travel',
    'Student awards/prizes',
    'Donation/Sponsorship',
    'Miscellaneous Expense',
    'Other Services',
    'Fund Transfer',
    'N/A'
  ];

  const categoryRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(categoryOptions, true)
    .setAllowInvalid(false)
    .setHelpText('Select expense category')
    .build();

  // DR/CR validation
  const drCrRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['DR', 'CR'], true)
    .setAllowInvalid(false)
    .setHelpText('Select DR (Debit) or CR (Credit)')
    .build();

  // GST Code validation
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
  
  // Receipt sections with formulas to extract from Google Forms
  const cols = ['D', 'F', 'H', 'J', 'L'];
  const receiptLabels = ['RECEIPT 1', 'RECEIPT 2', 'RECEIPT 3', 'RECEIPT 4', 'RECEIPT 5'];
  
  // Google Forms column mapping (starting from column K in row 5)
  // K=Description1, L=Company1, M=Date1, N=ReceiptNo1, O=Amount1, P=Softcopy1, Q=Bank1, R=MoreReceipts1
  // S=Description2, etc.
  const formColumnStarts = [11, 19, 27, 35, 43]; // K, S, AA, AI, AQ (columns 11, 19, 27, 35, 43)
  
  cols.forEach((col, i) => {
    sheet.getRange(`${col}7`)
      .setValue(receiptLabels[i])
      .setBackground('#93c47d')
      .setFontWeight('bold')
      .setHorizontalAlignment('center');
    
    const receiptFields = [
      ['DR/CR', ''],
      ['Description', `=IF(ISBLANK(INDIRECT("R5C"&${formColumnStarts[i]},FALSE)),"",INDIRECT("R5C"&${formColumnStarts[i]},FALSE))`],
      ['Category', ''],
      ['Category Code', ''],
      ['GST Code', ''],
      ['Company', `=IF(ISBLANK(INDIRECT("R5C"&${formColumnStarts[i]+1},FALSE)),"",INDIRECT("R5C"&${formColumnStarts[i]+1},FALSE))`],
      ['Date', `=IF(ISBLANK(INDIRECT("R5C"&${formColumnStarts[i]+2},FALSE)),"",INDIRECT("R5C"&${formColumnStarts[i]+2},FALSE))`],
      ['Receipt No.', `=IF(ISBLANK(INDIRECT("R5C"&${formColumnStarts[i]+3},FALSE)),"",INDIRECT("R5C"&${formColumnStarts[i]+3},FALSE))`],
      ['Amount', `=IF(ISBLANK(INDIRECT("R5C"&${formColumnStarts[i]+4},FALSE)),"",INDIRECT("R5C"&${formColumnStarts[i]+4},FALSE))`],
      ['Softcopy Link', `=IF(ISBLANK(INDIRECT("R5C"&${formColumnStarts[i]+5},FALSE)),"",INDIRECT("R5C"&${formColumnStarts[i]+5},FALSE))`],
      ['Bank Link', `=IF(ISBLANK(INDIRECT("R5C"&${formColumnStarts[i]+6},FALSE)),"",INDIRECT("R5C"&${formColumnStarts[i]+6},FALSE))`]
    ];
    
    receiptFields.forEach((field, j) => {
      const cellRef = `${col}${startRow + j}`;
      const colLetter = String.fromCharCode(col.charCodeAt(0) + 1);
      const valueCell = `${colLetter}${startRow + j}`;

      sheet.getRange(cellRef)
        .setValue(field[0])
        .setFontSize(9)
        .setFontWeight('bold');
      
      // Add formula in the value column
      if (field[1]) {
        sheet.getRange(valueCell)
        .setFormula(field[1])
        .setHorizontalAlignment('center');
      } else {
        sheet.getRange(valueCell).setHorizontalAlignment('center');
      }
      
      // Add data validation for specific fields
      if (field[0] === 'DR/CR') {
        sheet.getRange(valueCell).setDataValidation(drCrRule);
      } else if (field[0] === 'Category') {
        sheet.getRange(valueCell).setDataValidation(categoryRule);
        
        // Add Category Code formula in the next column (Category Code row)
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

  // Get existing rules first
  let rules = sheet.getConditionalFormatRules();

  // 1) Check if Email (B14) is filled, then highlight empty cells in B8:B22
  const emailRequiredCells = ['B8','B9','B10','B11','B12','B13','B14','B15','B16','B17','B18','B19','B20','B21','B22'];
  emailRequiredCells.forEach(cell => {
    const rule = SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied(`=AND($B$14<>"", ${cell}="")`)
      .setBackground('#f4cccc')
      .setRanges([sheet.getRange(cell)])
      .build();
    rules.push(rule);
  });

  // 2-6) Check if Description columns are filled, then highlight empty required cells
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

  // Apply all conditional formatting rules
  sheet.setConditionalFormatRules(rules);
  
  sheet.getRangeList(["B18", "E16", "G16", "I16", "K16", "M16"]).setNumberFormat("$#,##0.00");

  // Create template copy
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
    ['Email Address', 'Portfolio', 'CCA', 'Full Name', 'Matric No.', 'Phone No.']
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

/**
 * Adds custom menu on open.
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const menu = ui.createMenu(CONFIG.MENU.TITLE);
  CONFIG.MENU.ITEMS.forEach(item => menu.addItem(item.label, item.handler));
  menu.addToUi();
}

/**
 * Get a sheet by name and fail fast if it doesn't exist.
 */
function getSheetOrThrow(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet not found: ${sheetName}`);
  }
  return sheet;
}

/**
 * Read a list of single-cell ranges into an array of values.
 */
function getCellValues(sheet, ranges) {
  return ranges.map(range => sheet.getRange(range).getValue());
}

/**
 * Build a list of A1 ranges for a column and row span.
 */
function buildColumnRangeList(columnLetter, startRow, endRow) {
  const ranges = [];
  for (let r = startRow; r <= endRow; r++) {
    ranges.push(`${columnLetter}${r}`);
  }
  return ranges;
}

/**
 * Format date cells into the configured timezone and date format.
 */
function formatDateIfNeeded(value) {
  return value instanceof Date
    ? Utilities.formatDate(value, CONFIG.TIMEZONE, CONFIG.DATE_FORMAT)
    : value;
}

/**
 * Format numeric values into 2-decimal currency strings.
 */
function formatMoney(value) {
  return typeof value === 'number' ? value.toFixed(2) : value;
}

/**
 * Extract Drive file IDs from comma-separated URLs in a cell.
 */
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

/**
 * Iterate through each receipt block in the claims row.
 */
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

/**
 * Find a claim row in the Claims Data sheet by claim number.
 */
function findClaimRow(claimsData, claimNo) {
  return claimsData.find(r => r[CONFIG.CLAIMS_COL.NO] == claimNo);
}

/**
 * Confirm a potentially destructive batch action.
 */
function confirmBatchAction(title, message) {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert(title, message, ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) {
    ui.alert('Operation cancelled.');
    return false;
  }
  return true;
}

/**
 * Get master sheet rows that are marked for processing and not yet completed.
 */
function getPendingRowIndices(masterData, generateCol, statusCol) {
  const pending = [];
  for (let i = 1; i < masterData.length; i++) {
    if (masterData[i][generateCol] === true && masterData[i][statusCol] !== true) {
      pending.push(i);
    }
  }
  return pending;
}

/**
 * Mark a master sheet row as completed.
 */
function markMasterStatus(masterSheet, rowIndex, statusCol) {
  masterSheet.getRange(rowIndex + 1, statusCol + 1).setValue(true);
}

/**
 * Find a single file in a folder by exact name.
 */
function findFileByNameInFolder(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  return files.hasNext() ? files.next() : null;
}

/**
 * Ensure only one compiled PDF exists by deleting older copies with the same name.
 */
function deleteExistingFileByName(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  while (files.hasNext()) {
    files.next().setTrashed(true);
  }
}

/**
 * Export a Google file (Doc/Sheet) to a PDF blob.
 */
function exportToPdfBlob(file, outputName) {
  const pdfBlob = file.getBlob().getAs('application/pdf');
  return pdfBlob.setName(outputName);
}

/**
 * Load pdf-lib at runtime (required to merge PDFs).
 */
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

/**
 * Merge multiple PDF blobs into a single PDF blob.
 */
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

/**
 * Build the HTML list of receipt summaries for the email body.
 */
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

/**
 * Collect Drive file blobs from receipt and bank URL cells.
 */
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

/**
 * Build the email HTML body for a claim.
 */
function buildClaimEmailHtml(payload) {
  return `
    <div style="font-family: Arial, sans-serif; color: #000;">
      <p>Dear Ryan,</p>
      <p>Attached is the claims for ${payload.event}.</p>
      <p>To whom it may concern,</p>
      <p>I, ${payload.name.toUpperCase()}, ${payload.matric.toUpperCase()}, hereby authorise my treasurer, Ryan, to collect reimbursement on my behalf.</p>
      
      <strong>Claims Summary</strong><br>
      <table style="border-collapse: collapse; width: 100%; max-width: 450px; border: 1px solid black;">
        <tr><td style="padding: 5px; border: 1px solid black; background-color: #f9f9f9;">CCA</td><td style="padding: 5px; border: 1px solid black;">${payload.ccaName}</td></tr>
        <tr><td style="padding: 5px; border: 1px solid black; background-color: #f9f9f9;">Event</td><td style="padding: 5px; border: 1px solid black;">${payload.event.toUpperCase()}</td></tr>
        <tr><td style="padding: 5px; border: 1px solid black; background-color: #f9f9f9;">CCA Treasurer</td><td style="padding: 5px; border: 1px solid black;">${payload.name.toUpperCase()}</td></tr>
        <tr><td style="padding: 5px; border: 1px solid black; background-color: #f9f9f9;">Phone Number for PayNow</td><td style="padding: 5px; border: 1px solid black;">${payload.phone}</td></tr>
        <tr><td style="padding: 5px; border: 1px solid black; background-color: #f9f9f9;">Total collated amount</td><td style="padding: 5px; border: 1px solid black;">$${payload.totalFormatted}</td></tr>
      </table>
      
      <p><strong>Purpose of Purchase:</strong><br>
      ${payload.receiptListHtml || 'No individual receipts found.'}</p>
      
      <p><strong>Remarks:</strong><br>
      ${payload.remarks || 'No additional remarks.'}</p>
      
      <p>Thank you.</p>
    </div>`;
}

/**
 * Reset the Add Claim sheet by copying the template.
 */
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

function addClaim() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const templateSheet = getSheetOrThrow(spreadsheet, CONFIG.SHEETS.TEMPLATE);
  const addClaimSheet = getSheetOrThrow(spreadsheet, CONFIG.SHEETS.ADD_CLAIM);
  const claimsDataSheet = getSheetOrThrow(spreadsheet, CONFIG.SHEETS.CLAIMS_DATA);
  const masterSheet = getSheetOrThrow(spreadsheet, CONFIG.SHEETS.MASTER);
  const formResponsesSheet = getSheetOrThrow(spreadsheet, CONFIG.SHEETS.FORM_RESPONSES);

  // Build ranges for receipt blocks
  const receiptRanges = CONFIG.ADD_CLAIM.RECEIPT_COLUMNS.flatMap(column =>
    buildColumnRangeList(column, CONFIG.ADD_CLAIM.RECEIPT_ROWS.START, CONFIG.ADD_CLAIM.RECEIPT_ROWS.END)
  );

  // Claims Data row
  const rowClaimsData = [
    claimsDataSheet.getLastRow(), // No.
    ...getCellValues(addClaimSheet, CONFIG.ADD_CLAIM.CLAIM_RANGES),
    ...getCellValues(addClaimSheet, receiptRanges)
  ];

  // Master Sheet row
  const rowMasterData = [
    masterSheet.getLastRow(),
    ...getCellValues(addClaimSheet, CONFIG.ADD_CLAIM.MASTER_RANGES)
  ];

  claimsDataSheet.appendRow(rowClaimsData);
  masterSheet.appendRow(rowMasterData);
  masterSheet.getRange(rowMasterData[0] + 1, 10, 1, 5).insertCheckboxes();
  masterSheet.getRange(rowMasterData[0] + 1, 16, 1, 2).insertCheckboxes();

  // Reset "Add Claim" sheet
  resetAddClaimSheet(spreadsheet, templateSheet);

  // Switch to Master Sheet after adding a claim
  spreadsheet.setActiveSheet(formResponsesSheet);

  const ui = SpreadsheetApp.getUi();
  ui.alert('Successfully added claim');
}

function generateEmail() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const claimsDataSheet = getSheetOrThrow(spreadsheet, CONFIG.SHEETS.CLAIMS_DATA);
  const masterSheet = getSheetOrThrow(spreadsheet, CONFIG.SHEETS.MASTER);
  const MASTER_COL = CONFIG.MASTER_COL_EMAIL;

  // Read master and claims data once for performance
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

    // Find matching claim data
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
      referenceCode: row[CONFIG.CLAIMS_COL.REFERENCE_CODE]
    };

    const receiptListHtml = buildReceiptListHtml(row);
    const attachments = collectReceiptAttachments(row, payload.referenceCode);
    const htmlTemplate = buildClaimEmailHtml({
      ...payload,
      receiptListHtml
    });

    GmailApp.sendEmail(payload.recipientEmail, payload.referenceCode, "", {
      htmlBody: htmlTemplate,
      attachments: attachments
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

  // Read master and claims data once for performance
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
  
  // Get or create the main RFPs folder
  const mainFolderId = getOrCreateFolder(CONFIG.FOLDERS.RFP_ROOT_NAME, DriveApp.getFolderById(CONFIG.FOLDERS.RFP_ROOT_ID));
  const mainFolder = DriveApp.getFolderById(mainFolderId);

  let generatedCount = 0;

  pendingRows.forEach(rowIndex => {
    const masterRow = masterData[rowIndex];
    const claimNo = masterRow[MASTER_COL.NO];

    // Find the corresponding data row in Claims Data sheet
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

    // Create subfolder for this claim
    const subfolderName = `Claim No. ${claimNo} - (${referenceCode})`;
    const claimFolderId = getOrCreateFolder(subfolderName, mainFolder);
    const claimFolder = DriveApp.getFolderById(claimFolderId);

    try {
      // Generate all three forms
      generateLOA(claimNo, row, claimFolder);
      generateSummary(claimNo, row, claimFolder);
      generateRFP(claimNo, row, claimFolder);

      // Mark as generated
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

/**
 * Compile LOA, Summary, and RFP into a single PDF per claim.
 * Run only after confirming all three documents are formatted correctly.
 */
async function compileForms() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const claimsDataSheet = getSheetOrThrow(spreadsheet, CONFIG.SHEETS.CLAIMS_DATA);
  const masterSheet = getSheetOrThrow(spreadsheet, CONFIG.SHEETS.MASTER);
  const MASTER_COL = CONFIG.MASTER_COL_COMPILE;
  const FORM_COL = CONFIG.MASTER_COL_FORM;

  // Read master and claims data once for performance
  const masterData = masterSheet.getDataRange().getValues();
  const claimsData = claimsDataSheet.getDataRange().getValues();

  // Only compile for claims checked in the master sheet
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

  // Get or create the main RFPs folder
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
      console.log('Forms not generated for Claim No: ' + claimNo + '. Skipping compile.');
      skippedNotGenerated++;
      continue;
    }

    const row = findClaimRow(claimsData, claimNo);
    if (!row) {
      console.log('No matching data found for Claim No: ' + claimNo);
      skippedMissingData++;
      continue;
    }

    const referenceCode = (row[CONFIG.CLAIMS_COL.REFERENCE_CODE] || '').toString();
    if (!referenceCode) {
      console.log('No reference code found for Claim No: ' + claimNo);
      skippedMissingReference++;
      continue;
    }

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
      console.log(`Missing LOA/Summary/RFP for ${referenceCode}. Skipping compile.`);
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
    const errorDetails = errorMessages.length
      ? `\n\nErrors:\n${errorMessages.join('\n')}`
      : '';
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

/**
 * Get a folder by name within a parent, or create it if missing.
 */
function getOrCreateFolder(folderName, parentFolder) {
  const folders = parentFolder.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next().getId();
  } else {
    return parentFolder.createFolder(folderName).getId();
  }
}

/**
 * Generate LOA (List of Attachments) Google Doc
 * Alternates between receipt and bank transaction images/PDFs
 * One image per page, no text labels
 * PDFs are converted to images (one image per PDF page)
 */
function generateLOA(claimNo, row, folder) {
  const referenceCode = (row[CONFIG.CLAIMS_COL.REFERENCE_CODE] || "").toString();

  // Create a temporary Google Doc to build the LOA
  const doc = DocumentApp.create(`LOA No. ${claimNo} - (${referenceCode})`);
  const body = doc.getBody();

  body.appendPageBreak();

  let isFirstImage = true;

  // Helper: add a file (image or PDF) to the document
  const addFileToDoc = (fileId) => {
    if (!fileId) return;
    try {
      const file = DriveApp.getFileById(fileId);
      const mimeType = file.getMimeType();
      
      // Check if it's a PDF
      if (mimeType === 'application/pdf') {
        // For PDFs, convert to a Google Doc and copy contents
        let convertedFileId = null;
        let conversionSuccessful = false;
        let contentElements = []; // Store all content elements
        
        try {
          const pdfBlob = file.getBlob();
          
          const fileMetadata = {
            name: file.getName() + '_temp_' + new Date().getTime(),
            mimeType: 'application/vnd.google-apps.document',
            parents: [folder.getId()]
          };
          
          // Use Drive API v3 to upload and convert
          const convertedFile = Drive.Files.create(fileMetadata, pdfBlob, {
            supportsAllDrives: true,
            fields: 'id'
          });
          
          if (convertedFile && convertedFile.id) {
            convertedFileId = convertedFile.id;
            
            // Open the converted doc and get all content
            const tempDoc = DocumentApp.openById(convertedFile.id);
            const tempBody = tempDoc.getBody();
            
            // Get total number of child elements (paragraphs, images, tables, etc.)
            const numChildren = tempBody.getNumChildren();
            
            // Check if there's actual content (not just empty paragraphs)
            let hasContent = false;
            
            for (let i = 0; i < numChildren; i++) {
              const element = tempBody.getChild(i);
              const elementType = element.getType();
              
              // Check if element has meaningful content
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
              console.log(`Successfully converted PDF ${file.getName()} with ${contentElements.length} elements`);
            } else {
              console.log(`No meaningful content extracted from ${file.getName()}`);
            }
          }
          
        } catch (convertError) {
          console.log(`PDF conversion error for ${file.getName()}: ${convertError.message}`);
          conversionSuccessful = false;
        } finally {
          // Always delete the temporary converted doc if it was created
          if (convertedFileId) {
            try {
              DriveApp.getFileById(convertedFileId).setTrashed(true);
              console.log(`Cleaned up temp file for ${file.getName()}`);
            } catch (deleteError) {
              console.log(`Could not delete temp file: ${deleteError.message}`);
            }
          }
        }
        
        // Decide what to do based on conversion success
        if (conversionSuccessful && contentElements.length > 0) {
          // Conversion was successful - add all content to the document
          // Add page break before first element if this isn't the first item
          if (!isFirstImage) {
            body.appendPageBreak();
          }
          isFirstImage = false;
          
          // Add all content elements
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
          
          console.log(`Added ${contentElements.length} elements from ${file.getName()} to LOA`);
        } else {
          // Conversion failed or no meaningful content - copy original PDF to folder
          try {
            file.makeCopy(file.getName(), folder);
            console.log(`Copied original PDF ${file.getName()} to folder (conversion unsuccessful)`);
          } catch (copyError) {
            console.log(`Could not copy PDF: ${copyError.message}`);
          }
        }
        
      } else {
        // It's an image file - add it directly
        if (!isFirstImage) {
          body.appendPageBreak();
        }
        isFirstImage = false;
        
        const blob = file.getBlob();
        const image = body.appendImage(blob);
        
        // Scale image to fit page width (max 550px to account for margins)
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

  // Process each receipt block
  forEachReceipt(row, receipt => {
    if (!receipt.desc || receipt.amount === '') return;

    extractDriveFileIds(receipt.imgUrlCell).forEach(addFileToDoc);
    extractDriveFileIds(receipt.bankUrlCell).forEach(addFileToDoc);
  });

  // Remove the default empty paragraph at the beginning if we added images
  if (!isFirstImage) {
    const paragraphs = body.getParagraphs();
    if (paragraphs.length > 1 && paragraphs[0].getText() === '') {
      paragraphs[0].removeFromParent();
    }
  }

  // Save and close the document
  doc.saveAndClose();
  
  // Move the Google Doc to the claim folder
  const docFile = DriveApp.getFileById(doc.getId());
  docFile.moveTo(folder);
}

/**
 * Generate Summary spreadsheet by copying template and filling in data
 */
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
  
  console.log("attempting to copy template");

  // Copy the template to the claim folder
  const templateFile = DriveApp.getFileById(CONFIG.TEMPLATE_IDS.SUMMARY);
  const copiedFile = templateFile.makeCopy(`Summary No. ${claimNo} - (${referenceCode})`, folder);
  const ss = SpreadsheetApp.openById(copiedFile.getId());
  const sheet = ss.getActiveSheet();

  // Fill in the basic information fields
  sheet.getRange('B6').setValue(referenceCode);              // Ref
  sheet.getRange('B8').setValue(formattedDate);              // Date
  sheet.getRange('B12').setValue(claimDesc);                 // Event
  sheet.getRange('C20').setValue(financeDirectorName);       // Payee's Name (Finance Director)
  sheet.getRange('I20').setValue(financeDirectorMatric);     // Matric No (Finance Director)
  sheet.getRange('C24').setValue(totalFormatted);            // Total Amount
  sheet.getRange('I24').setValue(financeDirectorPhone);      // Contact No (Finance Director)
  sheet.getRange('B36').setValue(wbsNo);                     // WBS Account No.
  
  // Fill in receipt details (starting at row 31)
  let currentRow = 31;
  let receiptNum = 1;
  
  forEachReceipt(row, receipt => {
    if (!receipt.desc || receipt.amount === "") return;

    const formattedAmt = formatMoney(receipt.amount);
    
    // S/No (Column A)
    sheet.getRange(currentRow, 1).setValue(receiptNum || "");
    
    // Receipt/Inv No (Column B)
    sheet.getRange(currentRow, 2).setValue(receipt.receiptNo || "");
    
    // Description (Column D)
    sheet.getRange(currentRow, 4).setValue(receipt.desc || "");
    
    // Amount (Column I)
    sheet.getRange(currentRow, 9).setValue(formattedAmt || "");
    
    // GL (Column J)
    sheet.getRange(currentRow, 10).setValue(receipt.categoryCode || "");
    
    // GST Code (Column K)
    sheet.getRange(currentRow, 11).setValue(receipt.gstCode || "");
    
    currentRow++;
    receiptNum++;
  });
  
  // Update total in the total row (row 36, column I)
  sheet.getRange(36, 9).setValue(totalFormatted);
  
  // Flush changes to ensure they're saved
  SpreadsheetApp.flush();
}

/**
 * Generate RFP document with split Dollars and Cents placeholders
 */
function generateRFP(claimNo, row, folder) {
  const referenceCode = (row[CONFIG.CLAIMS_COL.REFERENCE_CODE] || "").toString();
  const financeDirectorName = (row[CONFIG.CLAIMS_COL.FINANCE_D_NAME] || "").toString();
  const financeDirectorMatric = (row[CONFIG.CLAIMS_COL.FINANCE_D_MATRIC] || "").toString();
  const total = row[CONFIG.CLAIMS_COL.TOTAL];
  const totalFormatted = formatMoney(total);
  const wbsNo = (row[CONFIG.CLAIMS_COL.WBS_NO] || "").toString();
  
  // Copy the template to the claim folder
  const templateFile = DriveApp.getFileById(CONFIG.TEMPLATE_IDS.RFP);
  const copiedFile = templateFile.makeCopy(`RFP No. ${claimNo} - (${referenceCode})`, folder);
  const doc = DocumentApp.openById(copiedFile.getId());
  const body = doc.getBody();
  
  // Static Replacements
  body.replaceText('{{NAME}}', financeDirectorName.toUpperCase());
  body.replaceText('{{MATRIC}}', financeDirectorMatric.toUpperCase());
  body.replaceText('{{TOTAL_AMOUNT}}', totalFormatted);
  body.replaceText('{{REFERENCE_CODE}}', referenceCode);
  body.replaceText('{{WBS_NO}}', wbsNo);

  // Loop through receipt placeholders
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
      // Clear out unused placeholders
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


// Helper function to get config value from Config sheet
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

// Centralized settings for easy customization.
const CONFIG = {
  MENU: {
    TITLE: 'Claims Tools',
    ITEMS: [
      { label: 'Add Claim', handler: 'addClaim' },
      { label: 'Generate Emails', handler: 'generateEmail' },
      { label: 'Generate Forms', handler: 'generateForm' },
      { label: 'Compile Forms', handler: 'compileForms' }
    ]
  },
  SHEETS: {
    TEMPLATE: 'Claim Data Template',
    ADD_CLAIM: 'Add Claim',
    CLAIMS_DATA: 'Claims Data',
    MASTER: 'Master Sheet',
    FORM_RESPONSES: 'Form Responses 1'
  },
  ADD_CLAIM: {
    CLAIM_RANGES: [
      'B8','B9','B10','B11','B12','B13','B14','B15','B16','B17','B18','B19','B20','B21','B22','B23'
    ],
    RECEIPT_COLUMNS: ['E','G','I','K','M'],  // Updated to match new layout
    RECEIPT_ROWS: { START: 8, END: 18 },
    MASTER_RANGES: ['B15','B16','B17','B18','B11','B13','B20','B25']  // Updated mapping
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
    DESC_START: 17
  },
  RECEIPT: {
    COUNT: 5,
    BLOCK_WIDTH: 11
  },
  // Template IDs now loaded from Config sheet
  get TEMPLATE_IDS() {
    return {
      SUMMARY: getConfigValue('SUMMARY_TEMPLATE_ID') || '',
      RFP: getConfigValue('RFP_TEMPLATE_ID') || ''
    };
  },
  // Folder IDs now loaded from Config sheet
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
