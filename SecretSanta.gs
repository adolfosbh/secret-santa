// Some constants
var DEBUG = false;
var NAME_COL = 0;
var EMAIL_COL = 1;


/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Invisible friend')
      .addItem('Do the draw', 'doTheDraw')
      .addToUi();
}

function getSelectionValues() {
  return SpreadsheetApp.getActiveRange().getValues();
}

function doTheDraw() {
  var ui = SpreadsheetApp.getUi();
  var selectionValues = getSelectionValues();
  var text = '';
  
  for (var i = 0; i < selectionValues.length; i++) {
    text += '* ' + selectionValues[i][NAME_COL] + '\n';
  }
  
  var result = ui.alert(
     'Please confirm',
     'Are you sure you want to start the draw with the following participants? (Emails will automatically be sent)\n' + text,
      ui.ButtonSet.YES_NO);

  // Process the user's response.
  if (result == ui.Button.YES) {
    var matches = getRandomMatch(range(selectionValues));
    sendEmails(matches, selectionValues);
    ui.alert('The email has been sent to the participants');
  } 
}

function getRandomMatch(fromValues) {
  var result = [];
  var toValues = fromValues.slice(); // We copy the array
  var numElements = fromValues.length;
  
  for (var i = 0; i < numElements; i++) { 
    var from = fromValues.splice(Math.floor(Math.random() * fromValues.length),1);    
    
    if (fromValues.length == 1) { // When there is one participant left, we must ensure it won't be assigned with herself
      var lastFrom = fromValues.splice(0,1);
      var lastFromInToValues = remove(toValues, lastFrom[0]);
      fromValues.push(lastFrom[0]);
      if (lastFromInToValues.length == 1) {  
        result.push({'from' : from[0], 'to' : lastFromInToValues[0]});
        continue;
      }
    }
    var fromInToValues = remove(toValues, from[0]);
    var to = toValues.splice(Math.floor(Math.random() * toValues.length),1);
    if (fromInToValues.length == 1) {
      toValues.push(fromInToValues[0]);
    }
    result.push({'from' : from[0], 'to' : to[0]});
  }
  return result;
}

function remove(array, value) {
  var idx = array.indexOf(value);
  if (idx > -1) {
    return array.splice(idx,1);
  } else {
    return [];
  }
}

function range(array) {
  var result = [];
  for (var i = 0; i < array.length; i++) {
     result.push(i);
  }
  return result;
}

function getEmailSubject(from) {
  // return "$name, welcome to the Invisible Friend game".replace('$name',from);
  return "[ESSA] $name, welcome to the 'Invisible Friend' game".replace('$name',from);
}

function getEmailContent(from, to) {
  return "Hello $name,\n\n\
Welcome to the Invisible Friend game (Also known as Secret Santa)\n\n\
You are now the invisible friend of **$assignmentName**\n\n\
Just some reminders:\n\
- The budget is around £10.\n\
- Don't forget to write the name of the gift's recipient on a visible part of it.\n\
- All the presents will be collected in a bag at the dinner place.\n\n\
Let's socialize!!\n\n\
Cheers,\n\
Adolfo.".replace('$name', from).replace('$assignmentName', to);
}

function sendEmails(matches, matrix) {
   
  if (DEBUG) {
     var ui = SpreadsheetApp.getUi();
     var text = '';
     for (var i = 0; i < matrix.length; i++) {
       var from = matches[i]['from'];
       var to = matches[i]['to'];
       text += matrix[from][NAME_COL] + ' <-> ' + matrix[to][NAME_COL] + '\n';
     }
    
    ui.alert('Assignments are:\n' + text+ 'Email content example:\n'+getEmailContent(matrix[0][NAME_COL], matrix[1][NAME_COL]));
  } else {
    for (var i = 0; i < matrix.length; i++) {
      var from = matches[i]['from'];
      var to = matches[i]['to'];
   
      MailApp.sendEmail(matrix[from][EMAIL_COL], 
                  getEmailSubject(matrix[from][NAME_COL]), 
                  getEmailContent(matrix[from][NAME_COL], matrix[to][NAME_COL]));
    }
  }
}