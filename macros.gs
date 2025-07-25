function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp, SlidesApp or FormApp.
  ui.createMenu('Macro Menu')
      .addItem('1 - Editing Format', 'Format')
      .addSeparator()
      .addItem('2 - Post-Editing Format', 'Cleanup')
      .addSeparator()
      .addItem('3 - Download for Rocksolid', 'downloadXLS_GUI')
      .addToUi();
}

function Format() {
  var spreadsheet = SpreadsheetApp.getActive();
  var final = spreadsheet.getDataRange().getLastRow();

  // Move Columns
  spreadsheet.getRange('K:K').activate();
  spreadsheet.getActiveSheet().moveColumns(spreadsheet.getRange('K:K'), 2);
  spreadsheet.getRange('E:E').activate();
  spreadsheet.getActiveSheet().moveColumns(spreadsheet.getRange('E:E'), 3);
  spreadsheet.getRange('K:K').activate();
  spreadsheet.getActiveSheet().moveColumns(spreadsheet.getRange('K:K'), 4);
  spreadsheet.getRange('E:I').activate();
  spreadsheet.getActiveSheet().deleteColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());

  // create new Columns
  spreadsheet.getRange('H1').activate();
  spreadsheet.getCurrentCell().setValue('Margin');
  spreadsheet.getRange('H2').activate();
  spreadsheet.getCurrentCell().setFormula('=ArrayFormula(((F2-D2)/F2))');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('H2:H'+final), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('I1').activate();
  spreadsheet.getCurrentCell().setValue('$ Dif');
  spreadsheet.getRange('I2').activate();
  spreadsheet.getCurrentCell().setFormula('=ArrayFormula(F2-E2)');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('I2:I'+final), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('H:H').activate();
  spreadsheet.getActiveSheet().moveColumns(spreadsheet.getRange('H:H'), 7);
  spreadsheet.getRange('I:I').activate();
  spreadsheet.getActiveSheet().moveColumns(spreadsheet.getRange('I:I'), 7);

  // making extended cost into a formula
  spreadsheet.getRange('I2').activate();
  spreadsheet.getCurrentCell().setFormula('=ARRAYFORMULA(B2*D2)');
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('I2:I'+final), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  
  // changing the 0 member retial stuff to suggested retail price because Do It Bess catlog puts those as nothing if you modified their prices in the last 30 days
  var storedValue = 0;
  for (let i = 2; i < final; i++) {
    if (spreadsheet.getRange("F"+i).isBlank()) {
      storedValue = spreadsheet.getRange('E'+i).getValue();
      spreadsheet.getRange("F"+i).setValue(storedValue);
    }
  }

  // symbols adder making them dollars and percents
  spreadsheet.getRangeList(['D1:G', 'I:I', 'J2', 'J5', 'J8']).activate()
  .setNumberFormat('"$"#,##0.00');
  spreadsheet.getRange('H:H').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('0.00%');  

  // Conditional colors  
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('H:H').activate();
  var conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(0, 1, SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('H1:H'+final)])
  .whenNumberGreaterThanOrEqualTo(0.35)
  .setBackground('#00FF00')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(1, 1, SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('H1:H'+final)])
  .whenNumberBetween(0.35, 0.05)
  .setBackground('#FFFF00')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(2, 1, SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('H1:H'+final)])
  .whenNumberLessThan(0.05)
  .setBackground('#FF0000')
  .setFontColor('#FFFFFF')
  .build());

  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  spreadsheet.getRange('G:G').activate();
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(3, 1, SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('G1:G'+final)])
  .whenNumberGreaterThan(0)
  .setBackground('#00FF00')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);
  conditionalFormatRules = spreadsheet.getActiveSheet().getConditionalFormatRules();
  conditionalFormatRules.splice(4, 1, SpreadsheetApp.newConditionalFormatRule()
  .setRanges([spreadsheet.getRange('G1:G'+final)])
  .whenNumberLessThan(0)
  .setBackground('#FF0000')
  .setFontColor('#FFFFFF')
  .build());
  spreadsheet.getActiveSheet().setConditionalFormatRules(conditionalFormatRules);

  // creating totals
  spreadsheet.getRange('K2').activate();
  spreadsheet.getCurrentCell().setValue('Total Extended Cost');
  spreadsheet.getRange('K3').activate();
  spreadsheet.getCurrentCell().setFormula('=Sum(I:I)');
  spreadsheet.getRange('K5').activate();
  spreadsheet.getCurrentCell().setValue('Gross Sales');
  spreadsheet.getRange('K6').activate();
  spreadsheet.getCurrentCell().setFormula('=SUMPRODUCT(B2:B, F2:F)');
  spreadsheet.getRange('K8').activate();
  spreadsheet.getCurrentCell().setValue('Net Profit');
  spreadsheet.getRange('K9').activate();
  spreadsheet.getCurrentCell().setFormula('=Sum(J6-J3)');
  spreadsheet.getRange('K11').activate();
  spreadsheet.getCurrentCell().setValue('Total SKUs');
  spreadsheet.getRange('K12').activate();
  spreadsheet.getCurrentCell().setFormula('=COUNTA(B2:B) - COUNTIF(B2:B, 0)');
  
  // total/sum modification
  spreadsheet.getRange('K2:K3').activate();
  spreadsheet.getActiveRangeList().setBackground('#ffff00');
  spreadsheet.getRange('K5:K6').activate();
  spreadsheet.getActiveRangeList().setBackground('#6d9eeb');
  spreadsheet.getRange('K8:K9').activate();
  spreadsheet.getActiveRangeList().setBackground('#00ff00');
  spreadsheet.getRange('K11:K12').activate();
  spreadsheet.getActiveRangeList().setBackground('#fbbc04');

  spreadsheet.getRangeList(['K2:K3', 'K5:K6', 'K8:K9', 'K11:K12']).activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  spreadsheet.getRangeList(['K3', 'K6', 'K9', 'K12']).activate()
  .setFontWeight('bold');

  // borders on all cells and changing color
  spreadsheet.getRange('A1:I'+final).activate();
  spreadsheet.getRange('A1:I'+final).applyRowBanding(SpreadsheetApp.BandingTheme.CYAN);
  var banding = spreadsheet.getRange('A1:I' + final).getBandings()[0];
  banding.setFirstRowColor('#ffffff')
  .setSecondRowColor('#e0f7fa')
  .setFooterRowColor(null);
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)

  // adding green background colors to the important columns
  spreadsheet.getRangeList(['A1:B'+final,'F1:F'+final]).setBackground('#b6d7a8');
  spreadsheet.getRange('A1:J'+final).createFilter();

  // create saved version of member retail
  spreadsheet.getRange('J:J').activate();
  spreadsheet.getRange('F:F').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  spreadsheet.getRange('J1').setValue("Original Retail")
  spreadsheet.getActiveRange().setHorizontalAlignment('center');
  spreadsheet.getActiveRange().setBackground(null);  

  // create the top section for the google sheet
  spreadsheet.getRange('1:1').activate();
  spreadsheet.getActiveSheet().insertRowsBefore(spreadsheet.getActiveRange().getRow(), 1);
  spreadsheet.getActiveRange().offset(0, 0, 1, spreadsheet.getActiveRange().getNumColumns()).activate();
  spreadsheet.getRangeList(['B1', 'F1']).setValue('EDIT');
  spreadsheet.getRange('G1:I1').mergeAcross();
  spreadsheet.getRange('G1:I1').setValue('Sort A-Z & Z-A');
  spreadsheet.getRangeList(['B1','F1', 'G1'])
  .setBackground('#e69138')
  .setFontColor('#ff0000');

  // all column changes
  spreadsheet.getRange('A:K').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getActiveRangeList().setFontSize(12);
  spreadsheet.getActiveSheet().autoResizeColumns(1, 11);
  spreadsheet.getActiveSheet().setColumnWidth(10, 15);

  // protect this area from mistakes
  var protection = spreadsheet.getRange('A1:K2').protect();
  protection.setDescription('Do Not Change').setWarningOnly(true);
  protection = spreadsheet.getRange('A3:A').protect();
  protection.setDescription('SKU Do Not Change').setWarningOnly(true);
  protection = spreadsheet.getRange('G3:K').protect();
  protection.setDescription('Formulas Do Not Change').setWarningOnly(true);
  protection = spreadsheet.getRange('C3:E').protect();
  protection.setDescription('Formulas Do Not Change').setWarningOnly(true);

  //freeze rows
  spreadsheet.getActiveSheet().setFrozenRows(2);
  
  var name = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  ExportFormat_();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(name), true);
  spreadsheet.getRange('A:A').activate();
}

function ExportFormat_() {
  var spreadsheet = SpreadsheetApp.getActive();
  var name = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  var final = spreadsheet.getDataRange().getLastRow();
  if(!spreadsheet.getSheetByName("SendSheet - .CSV"))
  {
    spreadsheet.insertSheet('SendSheet - .CSV');
  }

  // make the data refrence the original file
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('SendSheet - .CSV'), true);
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().setFormula("='" + name + "'!A2");
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('A1:A'+final), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('B1').activate();
  spreadsheet.getCurrentCell().setValue('OrderQuantity');
  spreadsheet.getRange('B2').activate();
  spreadsheet.getCurrentCell().setFormula("='" + name + "'!B3");
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('B2:B'+final), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('C1').activate();
  spreadsheet.getCurrentCell().setValue('MemberRetail');
  spreadsheet.getRange('C2').activate();
  spreadsheet.getCurrentCell().setFormula("='" + name + "'!F3");
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('C2:C'+final), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  // resize the columns
  spreadsheet.getActiveSheet().autoResizeColumns(1, 3);

  // get rid of formating
  spreadsheet.getRange('A:C').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('General');

  // protecting from deletion or modification in case people accidently do it
  var protection = spreadsheet.getRange('A:C').protect();
  protection.setDescription('Do Not Change').setWarningOnly(true);
};

function Cleanup(){
  var name = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  if(name !== "SendSheet - .CSV" && name !== 'RockSolidSendSheet - .XLSX'){
  Remove0QTY_();
  RockSolidSendSheet_();
  }else{
    SpreadsheetApp.getUi().alert("Please use this on the original document");
  }
};

function Remove0QTY_(){
  var name = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  var spreadsheet = SpreadsheetApp.getActive();
  var final = spreadsheet.getDataRange().getLastRow();

  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('SendSheet - .CSV'), true);

  // delete rows with 0 quatity
  const values = spreadsheet.getRange("B1:B" + final).getValues();
  let rowsToDelete = [];
  for (let i = 0; i < values.length; i++) {
    if (values[i][0] <= 0) {
      rowsToDelete.push(i + 1); // Store row number (1-indexed)
    }
  }
  for (let j = rowsToDelete.length - 1; j >= 0; j--) {
    spreadsheet.deleteRow(rowsToDelete[j]); 
  }

  // sorting the skus to get rid of options that haven't changed
  spreadsheet.getRange('A:C').createFilter();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getActiveSheet().getFilter().sort(1, true);
  spreadsheet.getRange('A:C').activate();
  spreadsheet.getActiveSheet().getFilter().remove();

  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(name), true);
}

function RockSolidSendSheet_(){
  var spreadsheet = SpreadsheetApp.getActive();
  var name = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  var final = spreadsheet.getDataRange().getLastRow();
  if(!spreadsheet.getSheetByName("RockSolidSendSheet - .XLSX"))
  {
    spreadsheet.insertSheet('RockSolidSendSheet - .XLSX');
  }

  // make the data refrence the original file
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('RockSolidSendSheet - .XLSX'), true);

  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().setValue("SKU");
  spreadsheet.getRange('B1').activate();
  spreadsheet.getCurrentCell().setValue('MemberRetail');

  spreadsheet.getRange('A2').activate();
  spreadsheet.getCurrentCell().setFormula("=IF(EXACT('"+ name + "'!F3, '"+ name +"'!J3),"+ '" "' + ", '" + name+"'!A3)");
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('A2:A'+final), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  spreadsheet.getRange('B2').activate();
  spreadsheet.getCurrentCell().setFormula("=IF(EXACT('"+ name + "'!F3, '"+ name +"'!J3),"+ '" "' + ", '" + name+"'!F3)");
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('B2:B'+final), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  // get rid of formating
  spreadsheet.getRange('A:B').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('General');
  
  // sorting the skus to get rid of options that haven't changed
  spreadsheet.getRange('A:B').createFilter();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getActiveSheet().getFilter().sort(1, true);
  spreadsheet.getRange('A:B').activate();
  spreadsheet.getActiveSheet().getFilter().remove();

  // resize the columns
  spreadsheet.getActiveSheet().autoResizeColumns(1, 2);

  // protecting the sheet in case people accidently delete
  var protection = spreadsheet.getRange('A:B').protect();
  protection.setDescription('Do Not Change').setWarningOnly(true);

  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(name), true);
};

function downloadXLS_GUI() {
  var name = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  if(name === 'RockSolidSendSheet - .XLSX'){
    var ssID = SpreadsheetApp.getActive().getId();
    var gid = SpreadsheetApp.getActive().getSheetId();
    var url = 'https://docs.google.com/spreadsheets/d/'+ssID+'/export?format=xlsx&gid='+ gid;
    var html = '<a href="' + url + '" target="_blank" download>Download XLSX</a>';
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html), 'Download');
  }else{
    SpreadsheetApp.getUi().alert("Please go to RockSolidSendSheet - .XLSX to download");
  }
  
};
