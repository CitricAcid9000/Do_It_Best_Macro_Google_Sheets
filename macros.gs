function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp, SlidesApp or FormApp.
  ui.createMenu('Macro Menu')
      .addItem('1 - Format', 'Format')
      .addSeparator()
      .addItem('2 - DownloadFormat', 'Cleanup')
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
  spreadsheet.getRange('J2').activate();
  spreadsheet.getCurrentCell().setValue('Total Extended Cost');
  spreadsheet.getRange('J3').activate();
  spreadsheet.getCurrentCell().setFormula('=Sum(I:I)');
  spreadsheet.getRange('J5').activate();
  spreadsheet.getCurrentCell().setValue('Gross Sales');
  spreadsheet.getRange('J6').activate();
  spreadsheet.getCurrentCell().setFormula('=SUMPRODUCT(B2:B, F2:F)');
  spreadsheet.getRange('J8').activate();
  spreadsheet.getCurrentCell().setValue('Net Profit');
  spreadsheet.getRange('J9').activate();
  spreadsheet.getCurrentCell().setFormula('=Sum(J6-J3)');
  spreadsheet.getRange('J11').activate();
  spreadsheet.getCurrentCell().setValue('Total SKUs');
  spreadsheet.getRange('J12').activate();
  spreadsheet.getCurrentCell().setFormula('=COUNTA(B2:B) - COUNTIF(B2:B, 0)');
  
  // total/sum modification
  spreadsheet.getRange('J2:J3').activate();
  spreadsheet.getActiveRangeList().setBackground('#ffff00');
  spreadsheet.getRange('J5:J6').activate();
  spreadsheet.getActiveRangeList().setBackground('#6d9eeb');
  spreadsheet.getRange('J8:J9').activate();
  spreadsheet.getActiveRangeList().setBackground('#00ff00');
  spreadsheet.getRange('J11:J12').activate();
  spreadsheet.getActiveRangeList().setBackground('#fbbc04');

  spreadsheet.getRangeList(['J2:J3', 'J5:J6', 'J8:J9', 'J11:J12']).activate();
  spreadsheet.getActiveRangeList().setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
  spreadsheet.getRangeList(['J3', 'J6', 'J9', 'J12']).activate()
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
  spreadsheet.getRangeList(['A1:B'+final,'F1:F'+final]).activate();
  spreadsheet.getActiveRangeList().setBackground('#b6d7a8');
  spreadsheet.getRange('A1:I'+final).createFilter();

  // all column changes
  spreadsheet.getRange('A:J').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  spreadsheet.getActiveRangeList().setFontSize(12)
  spreadsheet.getActiveSheet().autoResizeColumns(1, 10);

  // protect this area from mistakes
  var protection = spreadsheet.getRange('A1:J1').protect();
  protection.setDescription('Do Not Change').setWarningOnly(true);
  protection = spreadsheet.getRange('A2:A').protect();
  protection.setDescription('SKU Do Not Change').setWarningOnly(true);
  protection = spreadsheet.getRange('G2:J').protect();
  protection.setDescription('Formulas Do Not Change').setWarningOnly(true);
  protection = spreadsheet.getRange('C2:E').protect();
  protection.setDescription('Formulas Do Not Change').setWarningOnly(true);

  //freeze rows
  spreadsheet.getActiveSheet().setFrozenRows(1);

  // create saved version of member retail
  spreadsheet.getRange('U:U').activate();
  spreadsheet.getRange('F:F').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  
  var name = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  ExportFormat_();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName(name), true);
  spreadsheet.getRange('A:A').activate();
}

function ExportFormat_() {
  var spreadsheet = SpreadsheetApp.getActive();
  var name = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  var final = spreadsheet.getDataRange().getLastRow();
  spreadsheet.getRange('F:F').activate();
  if(!spreadsheet.getSheetByName("SendSheet - .CSV"))
  {
    spreadsheet.insertSheet('SendSheet - .CSV');
  }

  // make the data refrence the original file
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('SendSheet - .CSV'), true);
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().setFormula("='" + name + "'!A1");
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('A1:A'+final), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('B1').activate();
  spreadsheet.getCurrentCell().setValue('OrderQuantity');
  spreadsheet.getRange('B2').activate();
  spreadsheet.getCurrentCell().setFormula("='" + name + "'!B2");
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('B2:B'+final), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  spreadsheet.getRange('C1').activate();
  spreadsheet.getCurrentCell().setValue('MemberRetail');
  spreadsheet.getRange('C2').activate();
  spreadsheet.getCurrentCell().setFormula("='" + name + "'!F2");
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
  Remove0QTY_();
  RockSolidSendSheet_();
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
  spreadsheet.getCurrentCell().setFormula("=IF(EXACT('"+ name + "'!F2, '"+ name +"'!U2),"+ '" "' + ", '" + name+"'!A2)");
  spreadsheet.getActiveRange().autoFill(spreadsheet.getRange('A2:A'+final), SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);

  spreadsheet.getRange('B2').activate();
  spreadsheet.getCurrentCell().setFormula("=IF(EXACT('"+ name + "'!F2, '"+ name +"'!U2),"+ '" "' + ", '" + name+"'!F2)");
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
