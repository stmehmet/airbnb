// @ts-nocheck
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp or FormApp.
  ui.createMenu('QA Tool')
      .addItem('Send Emails', 'emailSender')
      .addItem('Import Samples', 'doStuff')
      .addToUi();
};

function doStuff(){
  determineLastRow();
  importer();
  linker();
};

function determineLastRow() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QA Tool");
  var lastRow = sheet.getLastRow();
  sheet.getRange("C2").setValue(lastRow);
};

function importer() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("QA Tool");
  var ssSource = SpreadsheetApp.openById("sheet id");
  var sheetSource = ssSource.getSheetByName("qa_weekly_sample");
  var previousLastRowRangeSource = sheetSource.getRange("G1")
  var previousLastRowSource = previousLastRowRangeSource.getValue();
  var lastRowSource = sheetSource.getLastRow();
  var startRowSource = previousLastRowSource + 1;
  var numRowsSource = lastRowSource - previousLastRowSource;
  var copyRangeSource = sheetSource.getRange(startRowSource,2,numRowsSource,4);
  var valuesSource = copyRangeSource.getValues();
  var lastRow = sheet.getLastRow();
  var startRow = lastRow + 1;
  var rangeDestination = sheet.getRange(startRow,2, numRowsSource,4);
  rangeDestination.setValues(valuesSource);
  var finalLastRow = sheet.getLastRow();
  var numRows = finalLastRow - 1;
  var previousLastRow = sheet.getRange("C2").getValue();
  var numRowsToolid = numRows - previousLastRow + 2;
  var toolidSource = sheet.getRange(previousLastRow,1);
  var toolidDestination = sheet.getRange(previousLastRow,1,numRowsToolid,1);
  toolidSource.autoFill(toolidDestination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  var escalationSource = sheet.getRange("F2");
  var escalationDestination = sheet.getRange(2,6,numRows,1);
  escalationSource.autoFill(escalationDestination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  var appealSource = sheet.getRange ("T2:Y2");
  var appealDestination = sheet.getRange(2,20,numRows,6);
  appealSource.autoFill(appealDestination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  var locationSource = sheet.getRange ("AD2");
  var locationDestination = sheet.getRange(2,30,numRows,1);
  locationSource.autoFill(locationDestination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}

function linker(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QA Tool");
  var startRow = sheet.getRange("C2").getValue() + 1; // Picking up from last existing row
  var numRows = sheet.getLastRow() - startRow + 1; // Number of rows to process
  // Fetch the range of cells
  var dataRange = sheet.getRange(startRow, 2, numRows, 1);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var ticket_id = row[0];
    var ticket_link = "\"internal link" + ticket_id + "\"";
    var conditionCheck = ticket_id.split("-")[0];
    var linkFormula = "=HYPERLINK(" + ticket_link + ", \"" + ticket_id + "\")";
    if (conditionCheck === "T49") {
    sheet.getRange(startRow + i, 2).setValue(linkFormula); // Set each cell in the range to native hyperlink() formula
    SpreadsheetApp.flush();
    }
  }
};

function protector(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("QA Tool");
  var range = sheet.getRange('A:F');
  range.protect().setDescription('No Touchy');
}

function emailSender() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert('Confirm','Are you sure you want to send email(s)?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    Logger.log('The user clicked "Yes."');
    pigeon();
  } else {
    Logger.log('The user clicked "No" or the close button in the dialog\'s title bar.');
  }   
}

function pigeon(){
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QA Tool");
  var startRow = 2; // Row of data to process
  var numRows = sheet.getLastRow()-1; // Number of rows to process
  // Fetch the range of cells
  var dataRange = sheet.getRange(startRow, 1, numRows, 30);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  for (var i = 0; i < data.length; ++i) {
    var row = data[i];
    var tool_id = row[0];
    var ticket_id = row[1];
    var formLink = `${prefilled1}${tool_id}${prefilled2}${ticket_id}`;
    var ticket_link = "internal link" + ticket_id;
    var ticket_hyperlink = ticket_id.link(ticket_link)
    var ldap = row[3]; // ldap column
    var email_address = ldap + '@company.com';
    var action_date_micro = row[4];
    var action_date_IST = Utilities.formatDate(new Date(action_date_micro), "IST", "yyyy-MM-dd HH:mm:ss 'IST'");
    var action_date_PST = Utilities.formatDate(new Date(action_date_micro), "PST", "yyyy-MM-dd HH:mm:ss 'PST'");
    var tied_to_escalation = row[5];
    var initial_email_sent_on = row[6];
    var reviewed = row[9];
    var decision_correct = row[10];
    var caused_escalation = row[11];
    var error_category = row[12];
    var error_subcategory = row[13];
    var error = row[14];
    var phase = row[15];
    var initial_feedback = row[16];
    var initial_ready_to_send = row[17];
    var initial_email_sent = row[18];
    var appealed = row[19];
    var appeal = row[20];
    var appeal_past_due = row[24];
    var appeal_accepted = row[25];
    var final_feedback = row[26];
    var final_ready_to_send = row[27];
    var final_email_sent = row[28];
    var location = row[29];

    var initial_table_IST = 
      "<table border='1',cellpadding='10',cellspacing ='0', width ='100%'>"
      +"<tr>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"QA Tool ID"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Ticket ID"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Ticket Closed on"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Error Category"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Error Subcategory"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Error"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Phase"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Feedback from Reviewer"+"</th>"
      +"</tr>"
      +"<tr>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+tool_id+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+ticket_hyperlink+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+action_date_IST+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+error_category+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+error_subcategory+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+error+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+phase+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+initial_feedback+"</td>"
      +"</tr>" 
      +"</table>";
    var final_table_IST = 
      "<table border='1',cellpadding='10',cellspacing ='0', width ='1000'>"
      +"<tr>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"QA Tool ID"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Ticket ID"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Ticket Closed on"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Error Category"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Error Subcategory"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Error"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Phase"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Initial Feedback from Reviewer"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Appeal"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Appeal Accepted"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Final Feedback from Reviewer"+"</th>"
      +"</tr>"
      +"<tr height= 200px, overflow-y= scroll, overflow-x= hidden>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+tool_id+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+ticket_hyperlink+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+action_date_IST+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+error_category+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+error_subcategory+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+error+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+phase+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+initial_feedback+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+appeal+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+appeal_accepted+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+final_feedback+"</td>"
      +"</tr>" 
      +"</table>";


    var initial_table_PST = 
      "<table border='1',cellpadding='10',cellspacing ='0', width ='100%'>"
      +"<tr>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"QA Tool ID"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Ticket ID"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Ticket Closed on"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Error Category"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Error Subcategory"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Error"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Phase"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Feedback from Reviewer"+"</th>"
      +"</tr>"
      +"<tr>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+tool_id+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+ticket_hyperlink+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+action_date_PST+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+error_category+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+error_subcategory+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+error+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+phase+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+initial_feedback+"</td>"
      +"</tr>" 
      +"</table>";
    var final_table_PST = 
      "<table border='1',cellpadding='10',cellspacing ='0', width ='1000'>"
      +"<tr>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"QA Tool ID"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Ticket ID"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Ticket Closed on"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Error Category"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Error Subcategory"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Error"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Phase"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Initial Feedback from Reviewer"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Appeal"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Appeal Accepted"+"</th>"
      +"<th bgcolor = '#CCCCCC', Align = 'center'>"+"Final Feedback from Reviewer"+"</th>"
      +"</tr>"
      +"<tr height= 200px, overflow-y= scroll, overflow-x= hidden>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+tool_id+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+ticket_hyperlink+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+action_date_PST+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+error_category+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+error_subcategory+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+error+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+phase+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+initial_feedback+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+appeal+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+appeal_accepted+"</td>"
      +"<td bgcolor = '#ffffff', Align = 'center'>"+final_feedback+"</td>"
      +"</tr>" 
      +"</table>";  
    
    var initial_message_IST =  `${greeting}<br><br>${initial_intro}<br><br>${initial_table_IST}<br><br>${initial_finale1} <a href=${formLink}>${linkLabel}</a><br><br>${initial_finale2}<br><br>${signature}`;
    var final_message_IST = `${greeting}<br><br>${final_intro1} ${ticket_id}${final_intro2}<br><br>${final_table_IST}<br><br>${final_finale}<br><br>${signature}`;
    var initial_message_PST =  `${greeting}<br><br>${initial_intro}<br><br>${initial_table_PST}<br><br>${initial_finale1} <a href=${formLink}>${linkLabel}</a><br><br>${initial_finale2}<br><br>${signature}`;
    var final_message_PST = `${greeting}<br><br>${final_intro1} ${ticket_id}${final_intro2}<br><br>${final_table_PST}<br><br>${final_finale}<br><br>${signature}`;

    if (decision_correct === "No" && reviewed === true && initial_email_sent !== EMAIL_SENT && initial_ready_to_send === true && location == "HYD") { // Prevents sending duplicates and makes sure email is ready to be sent
    var initial_subject = 'QA Feedback Notification for ' + ticket_id;
    GmailApp.sendEmail(email_address,initial_subject,"",{
    name: "QA Tool", 
    noReply: true,
    htmlBody: initial_message_IST,
    cc:"x@email.com",
    bcc: "y@email.com"
    });
    sheet.getRange(startRow + i, 7).setValue(current_date);
    sheet.getRange(startRow + i, 19).setValue(EMAIL_SENT);
    // Make sure the cell is updated right away in case the script is interrupted
    SpreadsheetApp.flush();
    } else if (decision_correct === "No" && reviewed === true && initial_email_sent !== EMAIL_SENT && initial_ready_to_send === true && location == "SVL") { // Prevents sending duplicates and makes sure email is ready to be sent
    var initial_subject = 'QA Feedback Notification for ' + ticket_id;
    GmailApp.sendEmail(email_address,initial_subject,"",{
    name: "QA Tool", 
    noReply: true,
    htmlBody: initial_message_PST,
    cc:"x@company.com",
    bcc: "y@company.com"
    });
    sheet.getRange(startRow + i, 7).setValue(current_date);
    sheet.getRange(startRow + i, 19).setValue(EMAIL_SENT);
    // Make sure the cell is updated right away in case the script is interrupted
    SpreadsheetApp.flush();
    }

    if(initial_email_sent === EMAIL_SENT && appealed ==="Yes" && initial_ready_to_send === true && final_email_sent !==EMAIL_SENT && final_ready_to_send === true && location == "HYD"){
    var final_subject = 'QA Appeal Feedback Notification for ' + ticket_id;
    GmailApp.sendEmail(email_address,final_subject,"",{
    name: "QA Tool", 
    noReply: true,
    htmlBody: final_message_IST,
    cc:"x@company.com",
    bcc: "y@company.com"
    });
    sheet.getRange(startRow + i, 29).setValue(EMAIL_SENT);
    // Make sure the cell is updated right away in case the script is interrupted
    SpreadsheetApp.flush();
    } else if(initial_email_sent === EMAIL_SENT && appealed ==="Yes" && initial_ready_to_send === true && final_email_sent !==EMAIL_SENT && final_ready_to_send === true && location == "SVL"){
    var final_subject = 'QA Appeal Feedback Notification for ' + ticket_id;
    GmailApp.sendEmail(email_address,final_subject,"",{
    name: "QA Tool", 
    noReply: true,
    htmlBody: final_message_PST,
    cc:"x@company.com",
    bcc: "y@company.com"
    });
    sheet.getRange(startRow + i, 29).setValue(EMAIL_SENT);
    // Make sure the cell is updated right away in case the script is interrupted
    SpreadsheetApp.flush();
    }  
  }
};

function sorter(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("QA Form Responses");
  var lastRow = sheet.getLastRow();
  var numRows = lastRow-1
  var lastColumn = sheet.getLastColumn();
  var numColumns = lastColumn-1;
  var range = sheet.getRange(2,1, numRows, lastColumn);
  range.sort(1);
};

function onEdit(e) {
  var oldValue;
  var newValue;
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  var activeCell = ss.getActiveCell();
  if((activeCell.getColumn() == 14 || activeCell.getColumn() == 15) && ss.getActiveSheet().getName()=="QA Tool") {
    newValue=e.value;
    oldValue=e.oldValue;
    if(!e.value) {
      activeCell.setValue("");
    }
  else {
    if (!e.oldValue) {
      activeCell.setValue(newValue);
      }
    else {
      activeCell.setValue(oldValue+', '+newValue);
      }
    }
  }
}

function qa_ticket_linker() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("QA Form Responses");
  var numRows = sheet.getLastRow() - 1
  var locationSource = sheet.getRange ("G2");
  var locationDestination = sheet.getRange(2,7,numRows,1);
  locationSource.autoFill(locationDestination, SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
}

function testFunction(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("QA Tool");
  var destination = sheet.getRange("B431");
  var storage = destination.getValue().split("-");
  if (storage[0] == "T49"){
    console.log(false)
  };
  //console.log(destination.getValue().split("-"));
};
