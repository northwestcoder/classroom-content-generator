// CHANGE ME: a calendar ID, we will use this to post events using DueDate
// you can find a calendar ID in the detailed settings are for a Google Calendar
// if left blank, this part of the tutorial will be ignored
// a calID looks like the following (which you need to change/update):
// var calID = 'c_9fo8l4s9c987js452024njn6o0@group.calendar.google.com';

var calID = '';



// for this tutorial, you do not have to change anything below this line

// the column names - in any order - that we expect from the Google Sheet
// we map columns names to a dictionary so that even if a Google Sheets user
// reorders the columns, this script will still function
var requiredColumns = [
  'ID',
  'AssignmentID',
  'StudentEmail',
  'StudentName',
  'AssignmentName',
  'DocID',
  'OutputFolder',
  'OutputDocID',
  'DueDate',
  'Description'
];

// main() is what you need to add to a new Apps Script onChange trigger
function main(e) {

  var theSource = e.source;
  var theSheet = theSource.getActiveSheet();
  var theActiveRange = theSheet.getActiveRange();
  var theActiveRow = theActiveRange.getRow();

  for (i in requiredColumns) {
    columnMap[i] = getColumnNumberByName(theSheet, requiredColumns[i]);
  }

  var thisID = theSheet.getRange(theActiveRow, columnMap["ID"]).getValue();
  var thisOutputDocID = theSheet.getRange(theActiveRow, columnMap["OutputDocID"]).getValue();

  // if this was a row deletion then ID will be empty, thus no work to do here
  // also, if this was *any* other type of change but the data already existed, 
  // then IGNORE and do not create duplicate content!
  if (thisID != "" && thisOutputDocID == '') {

    var thisAssignmentID = theSheet.getRange(theActiveRow, columnMap["AssignmentID"]).getValue();
    var thisStudentEmail = theSheet.getRange(theActiveRow, columnMap["StudentEmail"]).getValue();
    var thisStudentName = theSheet.getRange(theActiveRow, columnMap["StudentName"]).getValue();
    var thisAssignmentName = theSheet.getRange(theActiveRow, columnMap["AssignmentName"]).getValue();
    var thisDocID = theSheet.getRange(theActiveRow, columnMap["DocID"]).getValue();
    var thisOutputFolder = theSheet.getRange(theActiveRow, columnMap["OutputFolder"]).getValue();
    var thisDueDate = theSheet.getRange(theActiveRow, columnMap["DueDate"]).getValue();
    var thisDescription = theSheet.getRange(theActiveRow, columnMap["Description"]).getValue();

    // pretty print our new file name like this: "AssignmentID: AssignmentName - StudentName"
    // change as desired
    var newDocumenttName = thisAssignmentID + ": " + thisAssignmentName + " - " + thisStudentName;

    // copy the template file to our OutputFolder using our newly created newDocumentName
    var newCopy = DriveApp.getFileById(thisDocID).makeCopy(newDocumenttName, DriveApp.getFolderById(thisOutputFolder));
    var newMimeType = newCopy.getMimeType();
    Logger.log(newMimeType);

    // update our appsheet sheet with the new DocID and also add the student as an Google Drive Editor
    // also lock down the file so students cannot share, cheat, hack, etc.
    theSheet.getRange(theActiveRow, columnMap["OutputDocID"]).setValue(newCopy.getId());
    newCopy.addEditor(thisStudentEmail);
    newCopy.setShareableByEditors(false);
    newCopy.setSharing(DriveApp.Access.ANYONE, DriveApp.Permission.NONE);

    // generate a Google Calendar Event using the description, and add the student to this event
    if (calID != '') {

      // for our calendar event, generate a useful link to the student's homework or assignment
      var eventDocLink = renderMimeTypeURL(newCopy.getId(), newMimeType);

      var newEvent = CalendarApp.getCalendarById(calID).createEvent(newDocumenttName, thisDueDate, thisDueDate);
      newEvent.addGuest(thisStudentEmail);
      newEvent.setDescription(eventDocLink + thisDescription);
      newEvent.addEmailReminder(1440); // remind one day prior
      newEvent.addEmailReminder(10080); // remind one week prior      

    }

  } // done with our work: 1. copy doc, 2. editor assignment, 3. create calendar event
}

var columnMap = {};
function getColumnNumberByName(sheet, name) {
  var range = sheet.getRange(1, 1, 1, sheet.getMaxColumns());
  var values = range.getValues();
  for (var row in values) {
    for (var col in values[row]) {
      if (values[row][col] == name) {
        columnMap[name] = parseInt(col) + 1;
      }
    }
  }
}

function renderMimeTypeURL(docID, mimeType) {

  // IFS through the five supported Google Drive document types, else just return the docID
  if        (mimeType == 'application/vnd.google-apps.document') {
    return 'Your assignment:<br/><br/>https://docs.google.com/document/d/' + docID + '<br/><br/>';

  } else if (mimeType == 'application/vnd.google-apps.presentation') {
    return 'Your assignment:<br/><br/>https://docs.google.com/presentation/d/' + docID + '<br/><br/>';

  } else if (mimeType == 'application/vnd.google-apps.drawing') {
    return 'Your assignment:<br/><br/>https://docs.google.com/drawings/d/' + docID + '<br/><br/>';

  } else if (mimeType == 'application/vnd.google-apps.spreadsheet') {
    return 'Your assignment:<br/><br/>https://docs.google.com/spreadsheets/d/' + docID + '<br/><br/>';

  } else if (mimeType == 'application/vnd.google-apps.jam') {
    return 'Your assignment:<br/><br/>https://jamboard.google.com/d/' + docID + '<br/><br/>';

  } else {
    return docID
  }

}
