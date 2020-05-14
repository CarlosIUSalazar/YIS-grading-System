// FINAL VER 24 NOV 2019 18:20
function onOpen(e) {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Report Generator')
      .addItem('GenerateReport KG KC', 'generateReportKinderGartenKC')
      .addItem('GenerateReport KG KI', 'generateReportKinderGartenKI')

      .addSeparator()

      .addItem('GenerateReport G1 1S', 'generateReportGrade11S')
      .addItem('GenerateReport G1 1L', 'generateReportGrade11L')

      .addSeparator()

      .addItem('GenerateReport G2 2F', 'generateReportGrade22F')
      .addItem('GenerateReport G2 2M', 'generateReportGrade22M')

      .addSeparator()

      .addItem('GenerateReport G3 3B', 'generateReportGrade33B')
      .addItem('GenerateReport G3 3C', 'generateReportGrade33C')
      .addItem('GenerateReport G3 3M', 'generateReportGrade33M')

      .addSeparator()

      .addItem('GenerateReport G4 4H', 'generateReportGrade4H')
      .addItem('GenerateReport G4 4N', 'generateReportGrade4N')
      .addItem('GenerateReport G4 4W', 'generateReportGrade4W')

      .addSeparator()

      .addItem('GenerateReport G5 5G', 'generateReportGrade55G')
      .addItem('GenerateReport G5 5M', 'generateReportGrade55M')
      .addItem('GenerateReport G5 5R', 'generateReportGrade55R')


      .addSeparator()

      .addItem('About YIS Reporting System', 'showInfo')
      // .addItem('Highlight','highlightActiveCell')
      .addToUi();

}

function showInfo() {
  //ui.alert('Yokohama International School. Perfoming Arts School Report System. Ver. 1.0');

  Browser.msgBox("Yokohama International School. Perfoming Arts School Report System. Ver. 1.0.\n May this save Sam time so we can hang out more together. \n Love Carlos");
}


let documentsYearId = [{
      "KC": "1elRrdQiEy3ty9uGguCKrlsbNhlNCCjXaquWS7ksYlCA"
  },
  {
      "KI": "1aYnWE59kypJ9ddoww9h5gyM-6v9Xb_YbfseV9ms9ibs"
  },
  {
      "1S": "17FoCS8i12wCqcfV_FY-_SvdqGWWt6uMpt6uvoH7Qj30"
  },
  {
      "1L": "1D24BXQgnCOJ5KBlAuRPszZPDPb-0LDZ0jDXPR-574mo"
  },
  {
      "2F": "1NraUW_T404-w1Sr9pt1OU6hkPtnGdOGypvI2n7lRQos"
  },
  {
      "2M": "11I_IpEpAHqvva6G5ixynNnIdNkHGX_VRbep_RoDNwh8"
  },
  {
      "3B": "1ZQUD6-pX8EFIEeVrCUaKmWFEv8xYTDK3NrdCJgZx4rQ"
  },
  {
      "3C": "1OfXJgWNbxJoS8btBYXaXNGClovgVJLt0Bss8Lq0a1Wk"
  },
  {
      "3M": "1sEAFVng5QHLN_0ETrHclN-sy6T12K8xHUf-i98ABa8M"
  },
  {
      "4H": "1oGeeV4bXN_8zDGnGxuYh70OkcYIPAcviMYTqZvDZjrA"
  },
  {
      "4N": "1XrgBYdX7swPXCR56PyxN8673WLxTA0OtxgENNykcwF4"
  },
  {
      "4W": "1WAcRtIzRAaa6MvAgwhyqTK6nDrdST9H2IkXVDVvzD6E"
  },
  {
      "5G": "10yAvr44G1EO0WfOuzQ6DH3yYN8-rwECBt-22DO7X7_Y"
  },
  {
      "5M": "1acY_I2g0dIk6FQkE7RoPMRg1i2j-D12mHG4mUhZD6kA"
  },
  {
      "5R": "12NjA36qh_y4xmaJZ_cEjvzqwf9yyTfrtgeof1ZqPEAM"
  }
]

///////////////KINDER GARDEN KC //////////
function generateReportKinderGartenKC() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  // Returns the active cell
  let sheet = ss.getSheets()[0];
  let cell = ss.getActiveCell();
  //cell.setBackground("#ffceee");
  var positionRow = cell.getRow();
  var positionColumn = cell.getColumn()
  var firstName = sheet.getRange(positionRow, positionColumn).getValue();
  var lastName = sheet.getRange(positionRow, positionColumn + 1).getValue();
  var gender = sheet.getRange(positionRow, positionColumn + 3).getValue();
  // var classGroupName = "M1N";
  // var tutorialTeacherName = sheet.getRange('E1').getValue();
  var fullName = (firstName + ' ' + lastName);
  // Open Document to putthe comments
  var doc = DocumentApp.openById(documentsYearId[0]['KC']);
  //var reportDocId = doc.getId();
  //var reportDocUrl = doc.getUrl();
  var body = doc.getBody();
  //body.clear();
  var text = body.editAsText();
  //Looping through comments
  let count = 3;
  for (var j = 1; j <= 100; j++) { // This outer for loops through the ROWS of students
      firstName = sheet.getRange(positionRow, positionColumn).getValue(); //Update name on each outer Forloop iteration
      if ((sheet.getRange(positionRow, positionColumn + 2).getValue()) === 'KC') { //Checks only for KC students
          text.appendText(firstName + ' ' + sheet.getRange(positionRow, positionColumn + 4).getValue()); //Get the comment and append to report
          text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 5).getValue()); //Get the comment and append to report on comment 2 that starts with "This" 
          for (var i = 6; i <= 88; i++) { //This for loops through the COLUMNS for each student
              if ((sheet.getRange(positionRow, positionColumn + i).getValue()) !== 'NA') { //If it says NA skip it
                  if (count % 2 !== 0) {
                      text.appendText(' ' + firstName + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); //Get the name and comment and append to report on odd number columns after column 3     
                  }
                  if (count % 2 === 0) {
                      text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 3).getValue() + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); // Appends He or She depending the gender
                  }
              }
              count++;
          }
          text.appendText('\n');
          text.appendText('\n');
      } //if KC
      positionRow++;
  } // ouuter for
}



/////////////////////   KINDER GARTEN KI   //////////////////////////
function generateReportKinderGartenKI() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  // Returns the active cell
  let sheet = ss.getSheets()[0];
  let cell = ss.getActiveCell();
  //cell.setBackground("#ffceee");
  var positionRow = cell.getRow();
  var positionColumn = cell.getColumn()
  var firstName = sheet.getRange(positionRow, positionColumn).getValue();
  var lastName = sheet.getRange(positionRow, positionColumn + 1).getValue();
  var gender = sheet.getRange(positionRow, positionColumn + 3).getValue();
  // var classGroupName = "M1N";
  // var tutorialTeacherName = sheet.getRange('E1').getValue();
  var fullName = (firstName + ' ' + lastName);
  // Open Document to putthe comments
  var doc = DocumentApp.openById(documentsYearId[1]['KI']);
  //var reportDocId = doc.getId();
  //var reportDocUrl = doc.getUrl();
  var body = doc.getBody();
  //body.clear();
  var text = body.editAsText();
  //Looping through comments
  let count = 3;
  for (var j = 1; j <= 100; j++) { // This outer for loops through the ROWS of students
      firstName = sheet.getRange(positionRow, positionColumn).getValue(); //Update name on each outer Forloop iteration
      if ((sheet.getRange(positionRow, positionColumn + 2).getValue()) === 'KI') { //Checks only for KC students
          text.appendText(firstName + ' ' + sheet.getRange(positionRow, positionColumn + 4).getValue()); //Get the comment and append to report
          text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 5).getValue()); //Get the comment and append to report on comment 2 that starts with "This" 
          for (var i = 6; i <= 88; i++) { //This for loops through the COLUMNS for each student
              if ((sheet.getRange(positionRow, positionColumn + i).getValue()) !== 'NA') { //If it says NA skip it
                  if (count % 2 !== 0) {
                      text.appendText(' ' + firstName + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); //Get the name and comment and append to report on odd number columns after column 3     
                  }
                  if (count % 2 === 0) {
                      text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 3).getValue() + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); // Appends He or She depending the gender
                  }
              }
              count++;
          }
          text.appendText('\n');
          text.appendText('\n');
      } //if KC
      positionRow++;
  } // ouuter for
}



/////////////////////  YEAR 1 1S   //////////////////////////
function generateReportGrade11S() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  // Returns the active cell
  let sheet = ss.getSheets()[1]; //THIS ONE GRABS THE RIGHT SHEET WITHIN THE DOCUMENT
  let cell = ss.getActiveCell();
  //cell.setBackground("#ffceee");
  var positionRow = cell.getRow();
  var positionColumn = cell.getColumn()
  var firstName = sheet.getRange(positionRow, positionColumn).getValue();
  var lastName = sheet.getRange(positionRow, positionColumn + 1).getValue();
  var gender = sheet.getRange(positionRow, positionColumn + 3).getValue();
  // var classGroupName = "M1N";
  // var tutorialTeacherName = sheet.getRange('E1').getValue();
  var fullName = (firstName + ' ' + lastName);
  // Open Document to putthe comments
  var doc = DocumentApp.openById(documentsYearId[2]['1S']);
  //var reportDocId = doc.getId();
  //var reportDocUrl = doc.getUrl();
  var body = doc.getBody();
  //body.clear();
  var text = body.editAsText();
  //Looping through comments
  let count = 3;
  for (var j = 1; j <= 100; j++) { // This outer for loops through the ROWS of students
      firstName = sheet.getRange(positionRow, positionColumn).getValue(); //Update name on each outer Forloop iteration
      if ((sheet.getRange(positionRow, positionColumn + 2).getValue()) === '1S') { //Checks only for KC students
          text.appendText(firstName + ' ' + sheet.getRange(positionRow, positionColumn + 4).getValue()); //Get the comment and append to report
          text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 5).getValue()); //Get the comment and append to report on comment 2 that starts with "This" 
          for (var i = 6; i <= 88; i++) { //This for loops through the COLUMNS for each student
              if ((sheet.getRange(positionRow, positionColumn + i).getValue()) !== 'NA') { //If it says NA skip it
                  if (count % 2 !== 0) {
                      text.appendText(' ' + firstName + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); //Get the name and comment and append to report on odd number columns after column 3     
                  }
                  if (count % 2 === 0) {
                      text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 3).getValue() + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); // Appends He or She depending the gender
                  }
              }
              count++;
          }
          text.appendText('\n');
          text.appendText('\n');
      } //if KC
      positionRow++;
  } // ouuter for
}


/////////////////////   YEAR 1 1L   //////////////////////////
function generateReportGrade11L() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  // Returns the active cell
  let sheet = ss.getSheets()[1]; //THIS ONE GRABS THE RIGHT SHEET WITHIN THE DOCUMENT
  let cell = ss.getActiveCell();
  //cell.setBackground("#ffceee");
  var positionRow = cell.getRow();
  var positionColumn = cell.getColumn()
  var firstName = sheet.getRange(positionRow, positionColumn).getValue();
  var lastName = sheet.getRange(positionRow, positionColumn + 1).getValue();
  var gender = sheet.getRange(positionRow, positionColumn + 3).getValue();
  // var classGroupName = "M1N";
  // var tutorialTeacherName = sheet.getRange('E1').getValue();
  var fullName = (firstName + ' ' + lastName);
  // Open Document to putthe comments
  var doc = DocumentApp.openById(documentsYearId[3]['1L']);
  //var reportDocId = doc.getId();
  //var reportDocUrl = doc.getUrl();
  var body = doc.getBody();
  //body.clear();
  var text = body.editAsText();
  //Looping through comments
  let count = 3;
  for (var j = 1; j <= 100; j++) { // This outer for loops through the ROWS of students
      firstName = sheet.getRange(positionRow, positionColumn).getValue(); //Update name on each outer Forloop iteration
      if ((sheet.getRange(positionRow, positionColumn + 2).getValue()) === '1L') { //Checks only for KC students
          text.appendText(firstName + ' ' + sheet.getRange(positionRow, positionColumn + 4).getValue()); //Get the comment and append to report
          text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 5).getValue()); //Get the comment and append to report on comment 2 that starts with "This" 
          for (var i = 6; i <= 88; i++) { //This for loops through the COLUMNS for each student
              if ((sheet.getRange(positionRow, positionColumn + i).getValue()) !== 'NA') { //If it says NA skip it
                  if (count % 2 !== 0) {
                      text.appendText(' ' + firstName + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); //Get the name and comment and append to report on odd number columns after column 3     
                  }
                  if (count % 2 === 0) {
                      text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 3).getValue() + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); // Appends He or She depending the gender
                  }
              }
              count++;
          }
          text.appendText('\n');
          text.appendText('\n');
      } //if KC
      positionRow++;
  } // ouuter for
}


/////////////////////   YEAR 2 Group 1  2F   //////////////////////////
function generateReportGrade22F() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  // Returns the active cell
  let sheet = ss.getSheets()[2]; //THIS ONE GRABS THE RIGHT SHEET WITHIN THE DOCUMENT
  let cell = ss.getActiveCell();
  //cell.setBackground("#ffceee");
  var positionRow = cell.getRow();
  var positionColumn = cell.getColumn()
  var firstName = sheet.getRange(positionRow, positionColumn).getValue();
  var lastName = sheet.getRange(positionRow, positionColumn + 1).getValue();
  var gender = sheet.getRange(positionRow, positionColumn + 3).getValue();
  // var classGroupName = "M1N";
  // var tutorialTeacherName = sheet.getRange('E1').getValue();
  var fullName = (firstName + ' ' + lastName);
  // Open Document to putthe comments
  var doc = DocumentApp.openById(documentsYearId[4]['2F']);
  //var reportDocId = doc.getId();
  //var reportDocUrl = doc.getUrl();
  var body = doc.getBody();
  //body.clear();
  var text = body.editAsText();
  //Looping through comments
  let count = 3;
  for (var j = 1; j <= 100; j++) { // This outer for loops through the ROWS of students
      firstName = sheet.getRange(positionRow, positionColumn).getValue(); //Update name on each outer Forloop iteration
      if ((sheet.getRange(positionRow, positionColumn + 2).getValue()) === '2F') { //Checks only for KC students
          text.appendText(firstName + ' ' + sheet.getRange(positionRow, positionColumn + 4).getValue()); //Get the comment and append to report
          text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 5).getValue()); //Get the comment and append to report on comment 2 that starts with "This" 
          for (var i = 6; i <= 88; i++) { //This for loops through the COLUMNS for each student
              if ((sheet.getRange(positionRow, positionColumn + i).getValue()) !== 'NA') { //If it says NA skip it
                  if (count % 2 !== 0) {
                      text.appendText(' ' + firstName + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); //Get the name and comment and append to report on odd number columns after column 3     
                  }
                  if (count % 2 === 0) {
                      text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 3).getValue() + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); // Appends He or She depending the gender
                  }
              }
              count++;
          }
          text.appendText('\n');
          text.appendText('\n');
      } //if KC
      positionRow++;
  } // ouuter for
}


/////////////////////   YEAR 2 Group 2  2M   //////////////////////////
function generateReportGrade22M() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  // Returns the active cell
  let sheet = ss.getSheets()[2]; //THIS ONE GRABS THE RIGHT SHEET WITHIN THE DOCUMENT
  let cell = ss.getActiveCell();
  //cell.setBackground("#ffceee");
  var positionRow = cell.getRow();
  var positionColumn = cell.getColumn()
  var firstName = sheet.getRange(positionRow, positionColumn).getValue();
  var lastName = sheet.getRange(positionRow, positionColumn + 1).getValue();
  var gender = sheet.getRange(positionRow, positionColumn + 3).getValue();
  // var classGroupName = "M1N";
  // var tutorialTeacherName = sheet.getRange('E1').getValue();
  var fullName = (firstName + ' ' + lastName);
  // Open Document to putthe comments
  var doc = DocumentApp.openById(documentsYearId[5]['2M']);
  //var reportDocId = doc.getId();
  //var reportDocUrl = doc.getUrl();
  var body = doc.getBody();
  //body.clear();
  var text = body.editAsText();
  //Looping through comments
  let count = 3;
  for (var j = 1; j <= 100; j++) { // This outer for loops through the ROWS of students
      firstName = sheet.getRange(positionRow, positionColumn).getValue(); //Update name on each outer Forloop iteration
      if ((sheet.getRange(positionRow, positionColumn + 2).getValue()) === '2M') { //Checks only for KC students
          text.appendText(firstName + ' ' + sheet.getRange(positionRow, positionColumn + 4).getValue()); //Get the comment and append to report
          text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 5).getValue()); //Get the comment and append to report on comment 2 that starts with "This" 
          for (var i = 6; i <= 88; i++) { //This for loops through the COLUMNS for each student
              if ((sheet.getRange(positionRow, positionColumn + i).getValue()) !== 'NA') { //If it says NA skip it
                  if (count % 2 !== 0) {
                      text.appendText(' ' + firstName + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); //Get the name and comment and append to report on odd number columns after column 3     
                  }
                  if (count % 2 === 0) {
                      text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 3).getValue() + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); // Appends He or She depending the gender
                  }
              }
              count++;
          }
          text.appendText('\n');
          text.appendText('\n');
      } //if KC
      positionRow++;
  } // ouuter for
}



/////////////////////   YEAR 3 Group 3  3B   //////////////////////////
function generateReportGrade33B() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  // Returns the active cell
  let sheet = ss.getSheets()[3]; //THIS ONE GRABS THE RIGHT SHEET WITHIN THE DOCUMENT
  let cell = ss.getActiveCell();
  //cell.setBackground("#ffceee");
  var positionRow = cell.getRow();
  var positionColumn = cell.getColumn()
  var firstName = sheet.getRange(positionRow, positionColumn).getValue();
  var lastName = sheet.getRange(positionRow, positionColumn + 1).getValue();
  var gender = sheet.getRange(positionRow, positionColumn + 3).getValue();
  // var classGroupName = "M1N";
  // var tutorialTeacherName = sheet.getRange('E1').getValue();
  var fullName = (firstName + ' ' + lastName);
  // Open Document to putthe comments
  var doc = DocumentApp.openById(documentsYearId[6]['3B']);
  //var reportDocId = doc.getId();
  //var reportDocUrl = doc.getUrl();
  var body = doc.getBody();
  //body.clear();
  var text = body.editAsText();
  //Looping through comments
  let count = 3;
  for (var j = 1; j <= 100; j++) { // This outer for loops through the ROWS of students
      firstName = sheet.getRange(positionRow, positionColumn).getValue(); //Update name on each outer Forloop iteration
      if ((sheet.getRange(positionRow, positionColumn + 2).getValue()) === '3B') { //Checks only for KC students
          text.appendText(firstName + ' ' + sheet.getRange(positionRow, positionColumn + 4).getValue()); //Get the comment and append to report
          text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 5).getValue()); //Get the comment and append to report on comment 2 that starts with "This" 
          for (var i = 6; i <= 88; i++) { //This for loops through the COLUMNS for each student
              if ((sheet.getRange(positionRow, positionColumn + i).getValue()) !== 'NA') { //If it says NA skip it
                  if (count % 2 !== 0) {
                      text.appendText(' ' + firstName + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); //Get the name and comment and append to report on odd number columns after column 3     
                  }
                  if (count % 2 === 0) {
                      text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 3).getValue() + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); // Appends He or She depending the gender
                  }
              }
              count++;
          }
          text.appendText('\n');
          text.appendText('\n');
      } //if KC
      positionRow++;
  } // ouuter for
}


/////////////////////   YEAR 3 Group 2 3C  //////////////////////////
function generateReportGrade33C() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  // Returns the active cell
  let sheet = ss.getSheets()[3]; //THIS ONE GRABS THE RIGHT SHEET WITHIN THE DOCUMENT
  let cell = ss.getActiveCell();
  //cell.setBackground("#ffceee");
  var positionRow = cell.getRow();
  var positionColumn = cell.getColumn()
  var firstName = sheet.getRange(positionRow, positionColumn).getValue();
  var lastName = sheet.getRange(positionRow, positionColumn + 1).getValue();
  var gender = sheet.getRange(positionRow, positionColumn + 3).getValue();
  // var classGroupName = "M1N";
  // var tutorialTeacherName = sheet.getRange('E1').getValue();
  var fullName = (firstName + ' ' + lastName);
  // Open Document to putthe comments
  var doc = DocumentApp.openById(documentsYearId[7]['3C']);
  //var reportDocId = doc.getId();
  //var reportDocUrl = doc.getUrl();
  var body = doc.getBody();
  //body.clear();
  var text = body.editAsText();
  //Looping through comments
  let count = 3;
  for (var j = 1; j <= 100; j++) { // This outer for loops through the ROWS of students
      firstName = sheet.getRange(positionRow, positionColumn).getValue(); //Update name on each outer Forloop iteration
      if ((sheet.getRange(positionRow, positionColumn + 2).getValue()) === '3C') { //Checks only for KC students
          text.appendText(firstName + ' ' + sheet.getRange(positionRow, positionColumn + 4).getValue()); //Get the comment and append to report
          text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 5).getValue()); //Get the comment and append to report on comment 2 that starts with "This" 
          for (var i = 6; i <= 88; i++) { //This for loops through the COLUMNS for each student
              if ((sheet.getRange(positionRow, positionColumn + i).getValue()) !== 'NA') { //If it says NA skip it
                  if (count % 2 !== 0) {
                      text.appendText(' ' + firstName + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); //Get the name and comment and append to report on odd number columns after column 3     
                  }
                  if (count % 2 === 0) {
                      text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 3).getValue() + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); // Appends He or She depending the gender
                  }
              }
              count++;
          }
          text.appendText('\n');
          text.appendText('\n');
      } //if KC
      positionRow++;
  } // ouuter for
}


/////////////////////   YEAR 3 Group 1 3M  //////////////////////////
function generateReportGrade33M() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  // Returns the active cell
  let sheet = ss.getSheets()[3]; //THIS ONE GRABS THE RIGHT SHEET WITHIN THE DOCUMENT
  let cell = ss.getActiveCell();
  //cell.setBackground("#ffceee");
  var positionRow = cell.getRow();
  var positionColumn = cell.getColumn()
  var firstName = sheet.getRange(positionRow, positionColumn).getValue();
  var lastName = sheet.getRange(positionRow, positionColumn + 1).getValue();
  var gender = sheet.getRange(positionRow, positionColumn + 3).getValue();
  // var classGroupName = "M1N";
  // var tutorialTeacherName = sheet.getRange('E1').getValue();
  var fullName = (firstName + ' ' + lastName);
  // Open Document to putthe comments
  var doc = DocumentApp.openById(documentsYearId[8]['3M']);
  //var reportDocId = doc.getId();
  //var reportDocUrl = doc.getUrl();
  var body = doc.getBody();
  //body.clear();
  var text = body.editAsText();
  //Looping through comments
  let count = 3;
  for (var j = 1; j <= 100; j++) { // This outer for loops through the ROWS of students
      firstName = sheet.getRange(positionRow, positionColumn).getValue(); //Update name on each outer Forloop iteration
      if ((sheet.getRange(positionRow, positionColumn + 2).getValue()) === '3M') { //Checks only for KC students
          text.appendText(firstName + ' ' + sheet.getRange(positionRow, positionColumn + 4).getValue()); //Get the comment and append to report
          text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 5).getValue()); //Get the comment and append to report on comment 2 that starts with "This" 
          for (var i = 6; i <= 88; i++) { //This for loops through the COLUMNS for each student
              if ((sheet.getRange(positionRow, positionColumn + i).getValue()) !== 'NA') { //If it says NA skip it
                  if (count % 2 !== 0) {
                      text.appendText(' ' + firstName + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); //Get the name and comment and append to report on odd number columns after column 3     
                  }
                  if (count % 2 === 0) {
                      text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 3).getValue() + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); // Appends He or She depending the gender
                  }
              }
              count++;
          }
          text.appendText('\n');
          text.appendText('\n');
      } //if KC
      positionRow++;
  } // ouuter for
}



///////////////////   YEAR 4 Group 2  4H //////////////////////////
function generateReportGrade4H() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  // Returns the active cell
  let sheet = ss.getSheets()[4]; //THIS ONE GRABS THE RIGHT SHEET WITHIN THE DOCUMENT
  let cell = ss.getActiveCell();
  //cell.setBackground("#ffceee");
  var positionRow = cell.getRow();
  var positionColumn = cell.getColumn()
  var firstName = sheet.getRange(positionRow, positionColumn).getValue();
  var lastName = sheet.getRange(positionRow, positionColumn + 1).getValue();
  var gender = sheet.getRange(positionRow, positionColumn + 3).getValue();
  // var classGroupName = "M1N";
  // var tutorialTeacherName = sheet.getRange('E1').getValue();
  var fullName = (firstName + ' ' + lastName);
  // Open Document to putthe comments
  var doc = DocumentApp.openById(documentsYearId[9]['4H']);
  //var reportDocId = doc.getId();
  //var reportDocUrl = doc.getUrl();
  var body = doc.getBody();
  //body.clear();
  var text = body.editAsText();
  //Looping through comments
  let count = 3;
  for (var j = 1; j <= 100; j++) { // This outer for loops through the ROWS of students
      firstName = sheet.getRange(positionRow, positionColumn).getValue(); //Update name on each outer Forloop iteration
      if ((sheet.getRange(positionRow, positionColumn + 2).getValue()) === '4H') { //Checks only for KC students
          text.appendText(firstName + ' ' + sheet.getRange(positionRow, positionColumn + 4).getValue()); //Get the comment and append to report
          text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 5).getValue()); //Get the comment and append to report on comment 2 that starts with "This" 
          for (var i = 6; i <= 88; i++) { //This for loops through the COLUMNS for each student
              if ((sheet.getRange(positionRow, positionColumn + i).getValue()) !== 'NA') { //If it says NA skip it
                  if (count % 2 !== 0) {
                      text.appendText(' ' + firstName + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); //Get the name and comment and append to report on odd number columns after column 3     
                  }
                  if (count % 2 === 0) {
                      text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 3).getValue() + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); // Appends He or She depending the gender
                  }
              }
              count++;
          }
          text.appendText('\n');
          text.appendText('\n');
      } //if KC
      positionRow++;
  } // ouuter for
}


////////////////////   YEAR 4 Group 1  4N    //////////////////////////
function generateReportGrade4N() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  // Returns the active cell
  let sheet = ss.getSheets()[4]; //THIS ONE GRABS THE RIGHT SHEET WITHIN THE DOCUMENT
  let cell = ss.getActiveCell();
  //cell.setBackground("#ffceee");
  var positionRow = cell.getRow();
  var positionColumn = cell.getColumn()
  var firstName = sheet.getRange(positionRow, positionColumn).getValue();
  var lastName = sheet.getRange(positionRow, positionColumn + 1).getValue();
  var gender = sheet.getRange(positionRow, positionColumn + 3).getValue();
  // var classGroupName = "M1N";
  // var tutorialTeacherName = sheet.getRange('E1').getValue();
  var fullName = (firstName + ' ' + lastName);
  // Open Document to putthe comments
  var doc = DocumentApp.openById(documentsYearId[10]['4N']);
  //var reportDocId = doc.getId();
  //var reportDocUrl = doc.getUrl();
  var body = doc.getBody();
  //body.clear();
  var text = body.editAsText();
  //Looping through comments
  let count = 3;
  for (var j = 1; j <= 100; j++) { // This outer for loops through the ROWS of students
      firstName = sheet.getRange(positionRow, positionColumn).getValue(); //Update name on each outer Forloop iteration
      if ((sheet.getRange(positionRow, positionColumn + 2).getValue()) === '4N') { //Checks only for KC students
          text.appendText(firstName + ' ' + sheet.getRange(positionRow, positionColumn + 4).getValue()); //Get the comment and append to report
          text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 5).getValue()); //Get the comment and append to report on comment 2 that starts with "This" 
          for (var i = 6; i <= 88; i++) { //This for loops through the COLUMNS for each student
              if ((sheet.getRange(positionRow, positionColumn + i).getValue()) !== 'NA') { //If it says NA skip it
                  if (count % 2 !== 0) {
                      text.appendText(' ' + firstName + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); //Get the name and comment and append to report on odd number columns after column 3     
                  }
                  if (count % 2 === 0) {
                      text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 3).getValue() + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); // Appends He or She depending the gender
                  }
              }
              count++;
          }
          text.appendText('\n');
          text.appendText('\n');
      } //if KC
      positionRow++;
  } // ouuter for
}


////////////////////   YEAR 4 Group 3  4W   //////////////////////////
function generateReportGrade4W() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  // Returns the active cell
  let sheet = ss.getSheets()[4]; //THIS ONE GRABS THE RIGHT SHEET WITHIN THE DOCUMENT
  let cell = ss.getActiveCell();
  //cell.setBackground("#ffceee");
  var positionRow = cell.getRow();
  var positionColumn = cell.getColumn()
  var firstName = sheet.getRange(positionRow, positionColumn).getValue();
  var lastName = sheet.getRange(positionRow, positionColumn + 1).getValue();
  var gender = sheet.getRange(positionRow, positionColumn + 3).getValue();
  // var classGroupName = "M1N";
  // var tutorialTeacherName = sheet.getRange('E1').getValue();
  var fullName = (firstName + ' ' + lastName);
  // Open Document to putthe comments
  var doc = DocumentApp.openById(documentsYearId[11]['4W']);
  //var reportDocId = doc.getId();
  //var reportDocUrl = doc.getUrl();
  var body = doc.getBody();
  //body.clear();
  var text = body.editAsText();
  //Looping through comments
  let count = 3;
  for (var j = 1; j <= 100; j++) { // This outer for loops through the ROWS of students
      firstName = sheet.getRange(positionRow, positionColumn).getValue(); //Update name on each outer Forloop iteration
      if ((sheet.getRange(positionRow, positionColumn + 2).getValue()) === '4W') { //Checks only for KC students
          text.appendText(firstName + ' ' + sheet.getRange(positionRow, positionColumn + 4).getValue()); //Get the comment and append to report
          text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 5).getValue()); //Get the comment and append to report on comment 2 that starts with "This" 
          for (var i = 6; i <= 88; i++) { //This for loops through the COLUMNS for each student
              if ((sheet.getRange(positionRow, positionColumn + i).getValue()) !== 'NA') { //If it says NA skip it
                  if (count % 2 !== 0) {
                      text.appendText(' ' + firstName + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); //Get the name and comment and append to report on odd number columns after column 3     
                  }
                  if (count % 2 === 0) {
                      text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 3).getValue() + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); // Appends He or She depending the gender
                  }
              }
              count++;
          }
          text.appendText('\n');
          text.appendText('\n');
      } //if KC
      positionRow++;
  } // ouuter for
}


////////////////////   YEAR 5 Group 3   5G   //////////////////////////
function generateReportGrade55G() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Returns the active cell
  var sheet = ss.getSheets()[5];
  let cell = ss.getActiveCell();
  //cell.setBackground("#ffceee");
  var positionRow = cell.getRow();
  var positionColumn = cell.getColumn()
  var firstName = sheet.getRange(positionRow, positionColumn).getValue();
  var lastName = sheet.getRange(positionRow, positionColumn + 1).getValue();
  var gender = sheet.getRange(positionRow, positionColumn + 3).getValue();
  // var classGroupName = "M1N";
  // var tutorialTeacherName = sheet.getRange('E1').getValue();
  var fullName = (firstName + ' ' + lastName);
  // Open Document to putthe comments
  var doc = DocumentApp.openById(documentsYearId[12]['5G']);
  //var reportDocId = doc.getId();
  //var reportDocUrl = doc.getUrl();
  var body = doc.getBody();
  //body.clear();
  var text = body.editAsText();
  //Looping through comments
  let count = 3;
  for (var j = 1; j <= 100; j++) { // This outer for loops through the ROWS of students
      firstName = sheet.getRange(positionRow, positionColumn).getValue(); //Update name on each outer Forloop iteration
      if ((sheet.getRange(positionRow, positionColumn + 2).getValue()) === '5G') { //Checks only for KC students
          text.appendText(firstName + ' ' + sheet.getRange(positionRow, positionColumn + 4).getValue()); //Get the comment and append to report
          text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 5).getValue()); //Get the comment and append to report on comment 2 that starts with "This" 
          for (var i = 6; i <= 88; i++) { //This for loops through the COLUMNS for each student
              if ((sheet.getRange(positionRow, positionColumn + i).getValue()) !== 'NA') { //If it says NA skip it
                  if (count % 2 !== 0) {
                      text.appendText(' ' + firstName + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); //Get the name and comment and append to report on odd number columns after column 3     
                  }
                  if (count % 2 === 0) {
                      text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 3).getValue() + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); // Appends He or She depending the gender
                  }
              }
              count++;
          }
          text.appendText('\n');
          text.appendText('\n');
      } //if KC
      positionRow++;
  } // ouuter for
}

////////////////////   YEAR 5 Group 1   5M  //////////////////////////
function generateReportGrade55M() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Returns the active cell
  var sheet = ss.getSheets()[5];
  let cell = ss.getActiveCell();
  //cell.setBackground("#ffceee");
  var positionRow = cell.getRow();
  var positionColumn = cell.getColumn()
  var firstName = sheet.getRange(positionRow, positionColumn).getValue();
  var lastName = sheet.getRange(positionRow, positionColumn + 1).getValue();
  var gender = sheet.getRange(positionRow, positionColumn + 3).getValue();
  // var classGroupName = "M1N";
  // var tutorialTeacherName = sheet.getRange('E1').getValue();
  var fullName = (firstName + ' ' + lastName);
  // Open Document to putthe comments
  var doc = DocumentApp.openById(documentsYearId[13]['5M']);
  //var reportDocId = doc.getId();
  //var reportDocUrl = doc.getUrl();
  var body = doc.getBody();
  //body.clear();
  var text = body.editAsText();
  //Looping through comments
  let count = 3;
  for (var j = 1; j <= 100; j++) { // This outer for loops through the ROWS of students
      firstName = sheet.getRange(positionRow, positionColumn).getValue(); //Update name on each outer Forloop iteration
      if ((sheet.getRange(positionRow, positionColumn + 2).getValue()) === '5M') { //Checks only for KC students
          text.appendText(firstName + ' ' + sheet.getRange(positionRow, positionColumn + 4).getValue()); //Get the comment and append to report
          text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 5).getValue()); //Get the comment and append to report on comment 2 that starts with "This" 
          for (var i = 6; i <= 88; i++) { //This for loops through the COLUMNS for each student
              if ((sheet.getRange(positionRow, positionColumn + i).getValue()) !== 'NA') { //If it says NA skip it
                  if (count % 2 !== 0) {
                      text.appendText(' ' + firstName + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); //Get the name and comment and append to report on odd number columns after column 3     
                  }
                  if (count % 2 === 0) {
                      text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 3).getValue() + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); // Appends He or She depending the gender
                  }
              }
              count++;
          }
          text.appendText('\n');
          text.appendText('\n');
      } //if KC
      positionRow++;
  } // ouuter for
}

////////////////////   YEAR 5 Group 2   5R   //////////////////////////
function generateReportGrade55R() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Returns the active cell
  var sheet = ss.getSheets()[5];
  let cell = ss.getActiveCell();
  //cell.setBackground("#ffceee");
  var positionRow = cell.getRow();
  var positionColumn = cell.getColumn()
  var firstName = sheet.getRange(positionRow, positionColumn).getValue();
  var lastName = sheet.getRange(positionRow, positionColumn + 1).getValue();
  var gender = sheet.getRange(positionRow, positionColumn + 3).getValue();
  // var classGroupName = "M1N";
  // var tutorialTeacherName = sheet.getRange('E1').getValue();
  var fullName = (firstName + ' ' + lastName);
  // Open Document to putthe comments
  var doc = DocumentApp.openById(documentsYearId[14]['5R']);
  //var reportDocId = doc.getId();
  //var reportDocUrl = doc.getUrl();
  var body = doc.getBody();
  //body.clear();
  var text = body.editAsText();
  //Looping through comments
  let count = 3;
  for (var j = 1; j <= 100; j++) { // This outer for loops through the ROWS of students
      firstName = sheet.getRange(positionRow, positionColumn).getValue(); //Update name on each outer Forloop iteration
      if ((sheet.getRange(positionRow, positionColumn + 2).getValue()) === '5R') { //Checks only for KC students
          text.appendText(firstName + ' ' + sheet.getRange(positionRow, positionColumn + 4).getValue()); //Get the comment and append to report
          text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 5).getValue()); //Get the comment and append to report on comment 2 that starts with "This" 
          for (var i = 6; i <= 88; i++) { //This for loops through the COLUMNS for each student
              if ((sheet.getRange(positionRow, positionColumn + i).getValue()) !== 'NA') { //If it says NA skip it
                  if (count % 2 !== 0) {
                      text.appendText(' ' + firstName + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); //Get the name and comment and append to report on odd number columns after column 3     
                  }
                  if (count % 2 === 0) {
                      text.appendText(' ' + sheet.getRange(positionRow, positionColumn + 3).getValue() + ' ' + sheet.getRange(positionRow, positionColumn + i).getValue()); // Appends He or She depending the gender
                  }
              }
              count++;
          }
          text.appendText('\n');
          text.appendText('\n');
      } //if KC
      positionRow++;
  } // ouuter for
}