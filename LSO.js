function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Incident Report')
    .addItem('Generate Report', 'generateReport')
    .addItem('Clear Fields', 'clearFields')
    .addToUi();
}

// Checkboxes

function onEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;
  
  // Check if the edited cell is C17, C24, or C27
  if (range.getA1Notation() === 'C17' || range.getA1Notation() === 'C24' || range.getA1Notation() === 'C27') {
    checkCheckboxes();
  }
}

function checkCheckboxes() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // Check the value of C17
  var c17Value = sheet.getRange('C17').getValue();
  if (!c17Value) {
    sheet.hideRows(18, 5);  // Hide rows 18:21
  } else {
    sheet.showRows(18, 5);  // Show rows 18:21
  }

  // Check the value of C24
  var c24Value = sheet.getRange('C24').getValue();
  if (!c24Value) {
    sheet.hideRows(25, 2);  // Hide rows 25:26
    sheet.hideRows(23, 1);  // Hide rows 25:26
  } else {
    sheet.showRows(25, 2);  // Show rows 25:26
    sheet.showRows(23, 1);  // Hide rows 25:26
  }

  // Check the value of C27
  var c27Value = sheet.getRange('C27').getValue();
  if (!c27Value) {
    sheet.hideRows(28, 2);  // Hide rows 28:29
  } else {
    sheet.showRows(28, 2);  // Show rows 28:29
  }
}

// Clear Fields

function clearFields() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var pcodeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('pcode');

  // Clear specific cells
  sheet.getRange('C4').clearContent();
  sheet.getRange('C6').clearContent();
  sheet.getRange('C10:C22').clearContent();
  sheet.getRange('C24').clearContent();
  sheet.getRange('C27:28').clearContent();
  sheet.getRange('C28').clearContent();
  sheet.getRange('H10').clearContent();
  
  // Uncheck checkboxes if checked
  var checkboxCells = ['C24', 'C27'];
  for (var i = 0; i < checkboxCells.length; i++) {
    var cell = sheet.getRange(checkboxCells[i]);
    if (cell.getValue() === true) {
      cell.setValue(false);
    }
  }

  // Clear range in pcode sheet
  pcodeSheet.getRange('G5:G108').clearContent();

  // Recheck the state of checkboxes and hide/show rows accordingly
  checkCheckboxes();
}



// Generate Report

function generateReport() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
   var pcodeSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('pcode');
   
  var location = sheet.getRange('C4').getValue();
  var dateN = sheet.getRange('C5').getValue();
  // var formattedDate = Utilities.formatDate(new Date(date), "GMT", "MM/dd/yyyy"); // Change "GMT" to your desired time zone if needed
  var timeR = sheet.getRange('C6').getValue();
  var deputyR = sheet.getRange('C7').getValue();
  var dRank = sheet.getRange('C8').getValue();
  var deputy1 = sheet.getRange('C10').getValue();
  var deputy2 = sheet.getRange('C11').getValue();
  var deputy3 = sheet.getRange('C12').getValue();
  var deputy4 = sheet.getRange('C13').getValue();
  var deputy5 = sheet.getRange('C14').getValue();
  var deputy6 = sheet.getRange('C15').getValue();
  var deputy7 = sheet.getRange('C16').getValue();
  var c17Value = sheet.getRange('C17').getValue();

  var numCrims = sheet.getRange('C18').getValue();
  var numHostages = sheet.getRange('C19').getValue();
  var countdown = sheet.getRange('C20').getValue();
  var direction = sheet.getRange('C21').getValue();
  var chaseDescription = sheet.getRange('C22').getValue();
  var chargesChecked = sheet.getRange('C24').getValue();
  var itemsChecked = sheet.getRange('C27').getValue();
  var charges = chargesChecked ? sheet.getRange('C25').getValue() : '';
  var items = itemsChecked ? sheet.getRange('C28').getValue() : '';
  var timeFine = sheet.getRange('C26').getValue();

  var deputy1Text = c17Value ? `${deputy1} - Negotiator` : deputy1;

    var incidentReport = "";
  if (c17Value) {
    incidentReport = `Deputies above arrived at "${location}". ${numCrims} criminals had ${numHostages} hostages. "${deputy1}" negotiated for ${countdown} seconds, and the criminals agreed to head "${direction}".

CHASE REPORT:\n 
${chaseDescription}`;
  }

  var pcodeG105 = pcodeSheet.getRange('G105').getValue();
  var lemoyneDrugDirectionAct = pcodeG105 ? "\n$20 for The Lemoyne Drug-Direction Act" : "";

  var report = "```autohotkey\n" + `---INCIDENT REPORT---
PLACE OF INCIDENT: ${location}

DATE: ${dateN}

TIME:: ${timeR}

DEPUTY REPORTING: ${deputyR}

DEPUTY RANK: ${dRank}

DEPUTIES INVOLVED:
${deputy1Text}
${deputy2}
${deputy3}
${deputy4}
${deputy5}
${deputy6}
${deputy7}

INCIDENT REPORT:

${incidentReport}

ADDITIONAL NOTES:

${charges ? 'CHARGES: ' : ''}

${charges ?  charges : ''}

${charges ?  timeFine : ''}

${items ? 'EVIDENCE LOGGED:\n' + items : ''} ${lemoyneDrugDirectionAct}

SIGNED: ${dRank} ${deputyR}` + "\n```";

  report = report.replace(/\n{2,}/g, '\n\n'); // Replace multiple newlines with two newlines

  sheet.getRange('H10').setValue(report);
}
