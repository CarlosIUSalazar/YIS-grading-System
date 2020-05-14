// VER 14 May 2020 

function onOpen(e){  
    var ui = SpreadsheetApp.getUi();
    ui.createMenu('Report Generator')
    .addItem('GenerateReport KG KC','generateReportKinderGartenKC')
    .addItem('GenerateReport KG KI','generateReportKinderGartenKI')
  
    .addSeparator()
  
    .addItem('GenerateReport G1 1S','generateReportGrade11S')
    .addItem('GenerateReport G1 1L','generateReportGrade11L')
  
    .addSeparator()
    
    .addItem('GenerateReport G2 2F','generateReportGrade22F')
    .addItem('GenerateReport G2 2M','generateReportGrade22M')
  
    .addSeparator()
    
    .addItem('GenerateReport G3 3M','generateReportGrade33M')
    .addItem('GenerateReport G3 3C','generateReportGrade33C')
    .addItem('GenerateReport G3 3B','generateReportGrade33B')
    
    .addSeparator()
    
    .addItem('GenerateReport G4 4N','generateReportGrade4N')
    .addItem('GenerateReport G4 4H','generateReportGrade4H')
    .addItem('GenerateReport G4 4W','generateReportGrade4W')
  
    .addSeparator()
    
    .addItem('GenerateReport G5 5M','generateReportGrade55M')
    .addItem('GenerateReport G5 5R','generateReportGrade55R')
    .addItem('GenerateReport G5 5G','generateReportGrade55G')
  
    .addSeparator()
    
    .addItem('About YIS Reporting System','showInfo')
   // .addItem('Highlight','highlightActiveCell')
    .addToUi();
    
  }
  
  
  function showInfo() {
   //ui.alert('Yokohama International School. Perfoming Arts School Report System. Ver. 1.0');
    
   Browser.msgBox("Yokohama International School. Perfoming Arts School Report System. Ver. 1.0.\n May this save Sam time so we can hang out more together. \n Love Carlos"); 
  }
  
  
  ///////////////KINDER GARDEN KC //////////
  function generateReportKinderGartenKC() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
   
    // Returns the active cell
    let sheet = ss.getSheets()[0];
    let cell = ss.getActiveCell();
   //cell.setBackground("#ffceee");
    var positionRow = cell.getRow();
    var positionColumn = cell.getColumn()
   
    var firstName = sheet.getRange(positionRow,positionColumn).getValue();
    var lastName = sheet.getRange(positionRow,positionColumn+1).getValue();
    var gender = sheet.getRange(positionRow,positionColumn+3).getValue();
   // var classGroupName = "M1N";
   // var tutorialTeacherName = sheet.getRange('E1').getValue();
    var fullName = firstName; // + ' ' + lastName + ' ';
    
    
    // Open Document to putthe comments
    var doc = DocumentApp.openById('1elRrdQiEy3ty9uGguCKrlsbNhlNCCjXaquWS7ksYlCA');  // KI  1hWhbbFsyShKNuhF0Baqr-4olDCUfFjRbygAfJcmu6NE
  
    //var reportDocId = doc.getId();
    //var reportDocUrl = doc.getUrl();
    var body = doc.getBody();
    
    //body.clear();
    var text = body.editAsText();
    //body.setText(fullName);
    
    //Looping through comments
  let count = 3; 
  for (var j = 1; j<=100; j++){  // This outer for loops through the ROWS of students
   firstName = sheet.getRange(positionRow,positionColumn).getValue(); //Update name on each outer Forloop iteration
    
   if ((sheet.getRange(positionRow,positionColumn+2).getValue()) === 'KC'){  //Checks only for KC students
  
     text.appendText(firstName + ' ' + sheet.getRange(positionRow,positionColumn+4).getValue());  //Get the comment and append to report
  
     text.appendText(' ' + sheet.getRange(positionRow,positionColumn+5).getValue());  //Get the comment and append to report on comment 2 that starts with "This" 
  
    for (var i = 6; i <= 88; i++) {  //This for loops through the COLUMNS for each student
      if ((sheet.getRange(positionRow,positionColumn+i).getValue()) !== 'NA') {  //If it says NA skip it
        if (count % 2 !== 0){
          text.appendText(' ' + firstName + ' ' + sheet.getRange(positionRow,positionColumn+i).getValue());  //Get the name and comment and append to report on odd number columns after column 3     
        }  
        if (count % 2 === 0) {
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn+3).getValue() + ' ' + sheet.getRange(positionRow,positionColumn+i).getValue());  // Appends He or She depending the gender
        }
      }
      count++;
    }
    text.appendText('\n');
    text.appendText('\n');
    }  //if KC
  
    positionRow++;
  } // ouuter for
  }
      
  
  
  /////////////////////   KINDER GARTEN KI   //////////////////////////
  
  function generateReportKinderGartenKI() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
   
    // Returns the active cell
    var sheet = ss.getSheets()[0];
    var cell = ss.getActiveCell();
   //cell.setBackground("#ffceee");
    var positionRow = cell.getRow();
    var positionColumn = cell.getColumn()
   
    var firstName = sheet.getRange(positionRow,positionColumn).getValue();
    var lastName = sheet.getRange(positionRow,positionColumn+1).getValue();
    var gender = sheet.getRange(positionRow,positionColumn+3).getValue();
   // var classGroupName = "M1N";
   // var tutorialTeacherName = sheet.getRange('E1').getValue();
    var fullName = firstName + ' ' + lastName + ' ';
    
    
    // Open Document to putthe comments
    var doc = DocumentApp.openById('1aYnWE59kypJ9ddoww9h5gyM-6v9Xb_YbfseV9ms9ibs');  // KI  1hWhbbFsyShKNuhF0Baqr-4olDCUfFjRbygAfJcmu6NE
    //var reportDocId = doc.getId();
    //var reportDocUrl = doc.getUrl();
    var body = doc.getBody();
    
    //body.clear();
    var text = body.editAsText();
    //body.setText(fullName);
  
    //Looping through comments
   for (var j = 1; j<=100; j++){  // This outer for loops through the ROWS of students
    
   if ((sheet.getRange(positionRow,positionColumn+2).getValue()) === 'KI'){  //Checks only for KI students
    
    text.appendText(sheet.getRange(positionRow,positionColumn).getValue() + ' '); //Appends Name spaces
    for (var i = 4; i <= 92; i++){  //This for loops through the COLUMNS for each student
      
      if (sheet.getRange(positionRow,positionColumn+i).getValue() != 'NA') {  //If it says NA skip it
          
        text.appendText(sheet.getRange(positionRow,positionColumn+i).getValue());  //Get the comment and append to report
        
        var random = Math.floor(Math.random() *100);   //The chances of appending the name rather than he/she is 89%
        
        if (random < 89) {   
          
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn+3).getValue() + ' ');  // Appends He or She depending the gender
        
        } else if (random >=90) {   
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
          
        };
       }
     };
    text.appendText('\n');
    text.appendText('\n');
    }  //if KC
  
    positionRow++;
  } // ouuter for
    
  }     
        
  
  
  
  
  /////////////////////  YEAR 1 1S   //////////////////////////
  
  function generateReportGrade11S() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
   
    // Returns the active cell
    var sheet = ss.getSheets()[1];
    var cell = ss.getActiveCell();
   //cell.setBackground("#ffceee");
    var positionRow = cell.getRow();
    var positionColumn = cell.getColumn()
   
    var firstName = sheet.getRange(positionRow,positionColumn).getValue();
    var lastName = sheet.getRange(positionRow,positionColumn+1).getValue();
    var gender = sheet.getRange(positionRow,positionColumn+3).getValue();
   // var classGroupName = "M1N";
   // var tutorialTeacherName = sheet.getRange('E1').getValue();
    var fullName = firstName + ' ' + lastName + ' ';
    
    
    // Open Document to putthe comments
    var doc = DocumentApp.openById('17FoCS8i12wCqcfV_FY-_SvdqGWWt6uMpt6uvoH7Qj30');  
    //var reportDocId = doc.getId();
    //var reportDocUrl = doc.getUrl();
    var body = doc.getBody();
    
    //body.clear();
    var text = body.editAsText();
    //body.setText(fullName);
  
   //Looping through comments
   for (var j = 1; j<=100; j++){  // This outer for loops through the ROWS of students
    
   if ((sheet.getRange(positionRow,positionColumn+2).getValue()) === '1S'){  //Checks only for KI students
    
    text.appendText(sheet.getRange(positionRow,positionColumn).getValue() + ' '); //Appends Name spaces
    for (var i = 4; i < 83; i++){  //This for loops through the COLUMNS for each student. 
      
      if (sheet.getRange(positionRow,positionColumn+i).getValue() != 'NA') {  //If it says NA skip it
          
        text.appendText(sheet.getRange(positionRow,positionColumn+i).getValue());  //Get the comment and append to report
        
        var random = Math.floor(Math.random() *100);   //The chances of appending the name rather than he/she is 89%
        
        if (i === 4) { //The first column will always display te name (after Unit Overview)
        
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
        
        }  else if (random < 89 && i >=5) {   // i >= 5 so it doesnt add he/she/name in Unit Overview Column 5
          
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn+3).getValue() + ' ');  // Appends He or She depending the gender
        
        } else if (random >= 89 && i >=5){   
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
          
        };
       }
     };
    text.appendText('\n');
    text.appendText('\n');
    }  //if KC
  
    positionRow++;
  } // ouuter for
    
  }     
   
  
  
  
  /////////////////////   YEAR 1 1L   //////////////////////////
  
  function generateReportGrade11L() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
   
    // Returns the active cell
    var sheet = ss.getSheets()[1];
    var cell = ss.getActiveCell();
   //cell.setBackground("#ffceee");
    var positionRow = cell.getRow();
    var positionColumn = cell.getColumn()
   
    var firstName = sheet.getRange(positionRow,positionColumn).getValue();
    var lastName = sheet.getRange(positionRow,positionColumn+1).getValue();
    var gender = sheet.getRange(positionRow,positionColumn+3).getValue();
    var fullName = firstName + ' ' + lastName + ' ';
    
    
    // Open Document to putthe comments
    var doc = DocumentApp.openById('1D24BXQgnCOJ5KBlAuRPszZPDPb-0LDZ0jDXPR-574mo');  
  
    var body = doc.getBody();
    
    var text = body.editAsText();
  
  //Looping through comments
   for (var j = 1; j<=100; j++){  // This outer for loops through the ROWS of students
    
   if ((sheet.getRange(positionRow,positionColumn+2).getValue()) === '1L'){  //Checks only for KI students
    
    text.appendText(sheet.getRange(positionRow,positionColumn).getValue() + ' '); //Appends Name spaces
    for (var i = 4; i < 83; i++){  //This for loops through the COLUMNS for each student. 
      
      if (sheet.getRange(positionRow,positionColumn+i).getValue() != 'NA') {  //If it says NA skip it
          
        text.appendText(sheet.getRange(positionRow,positionColumn+i).getValue());  //Get the comment and append to report
        
        var random = Math.floor(Math.random() *100);   //The chances of appending the name rather than he/she is 89%
        
        if (i === 4) { //The first column will always display te name (after Unit Overview)
        
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
        
        }  else if (random < 89 && i >=5) {   // i >= 5 so it doesnt add he/she/name in Unit Overview Column 5
          
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn+3).getValue() + ' ');  // Appends He or She depending the gender
        
        } else if (random >= 89 && i >=5){   
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
          
        };
       }
     };
    text.appendText('\n');
    text.appendText('\n');
    }  //if KC
  
    positionRow++;
  } // ouuter for
    
  }     
   
  
  
  
  
  /////////////////////   YEAR 2 Group 1  2F   //////////////////////////
  
  function generateReportGrade22F() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
   
    // Returns the active cell
    var sheet = ss.getSheets()[2];
    var cell = ss.getActiveCell();
   //cell.setBackground("#ffceee");
    var positionRow = cell.getRow();
    var positionColumn = cell.getColumn()
   
    var firstName = sheet.getRange(positionRow,positionColumn).getValue();
    var lastName = sheet.getRange(positionRow,positionColumn+1).getValue();
    var gender = sheet.getRange(positionRow,positionColumn+3).getValue();
    var fullName = firstName + ' ' + lastName + ' ';
    
    
    // Open Document to putthe comments
    var doc = DocumentApp.openById('1NraUW_T404-w1Sr9pt1OU6hkPtnGdOGypvI2n7lRQos');  
  
    var body = doc.getBody();
    
    var text = body.editAsText();
  
  //Looping through comments
   for (var j = 1; j<=100; j++){  // This outer for loops through the ROWS of students
    
   if ((sheet.getRange(positionRow,positionColumn+2).getValue()) === '2F'){  //Checks only for KI students
    
    text.appendText(sheet.getRange(positionRow,positionColumn).getValue() + ' '); //Appends Name spaces
    for (var i = 4; i < 113; i++){  //This for loops through the COLUMNS for each student. 
      
      if (sheet.getRange(positionRow,positionColumn+i).getValue() != 'NA') {  //If it says NA skip it
          
        text.appendText(sheet.getRange(positionRow,positionColumn+i).getValue());  //Get the comment and append to report
        
        var random = Math.floor(Math.random() *100);   //The chances of appending the name rather than he/she is 89%
        
        if (i === 4) { //The first column will always display te name (after Unit Overview)
        
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
        
        }  else if (random < 89 && i >=5) {   // i >= 5 so it doesnt add he/she/name in Unit Overview Column 5
          
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn+3).getValue() + ' ');  // Appends He or She depending the gender
        
        } else if (random >= 89 && i >=5){   
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
          
        };
       }
     };
    text.appendText('\n');
    text.appendText('\n');
    }  //if KC
  
    positionRow++;
  } // ouuter for
    
  }     
   
  
  
  /////////////////////   YEAR 2 Group 2  2M   //////////////////////////
  
  function generateReportGrade22M() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
   
    // Returns the active cell
    var sheet = ss.getSheets()[2];
    var cell = ss.getActiveCell();
   //cell.setBackground("#ffceee");
    var positionRow = cell.getRow();
    var positionColumn = cell.getColumn()
   
    var firstName = sheet.getRange(positionRow,positionColumn).getValue();
    var lastName = sheet.getRange(positionRow,positionColumn+1).getValue();
    var gender = sheet.getRange(positionRow,positionColumn+3).getValue();
    var fullName = firstName + ' ' + lastName + ' ';
    
    
    // Open Document to putthe comments
    var doc = DocumentApp.openById('11I_IpEpAHqvva6G5ixynNnIdNkHGX_VRbep_RoDNwh8');  
  
    var body = doc.getBody();
    
    var text = body.editAsText();
  
  //Looping through comments
   for (var j = 1; j<=100; j++){  // This outer for loops through the ROWS of students
    
   if ((sheet.getRange(positionRow,positionColumn+2).getValue()) === '2M'){  //Checks only for KI students
    
    text.appendText(sheet.getRange(positionRow,positionColumn).getValue() + ' '); //Appends Name spaces
    for (var i = 4; i < 113; i++){  //This for loops through the COLUMNS for each student. 
      
      if (sheet.getRange(positionRow,positionColumn+i).getValue() != 'NA') {  //If it says NA skip it
          
        text.appendText(sheet.getRange(positionRow,positionColumn+i).getValue());  //Get the comment and append to report
        
        var random = Math.floor(Math.random() *100);   //The chances of appending the name rather than he/she is 89%
        
        if (i === 4) { //The first column will always display te name (after Unit Overview)
        
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
        
        }  else if (random < 89 && i >=5) {   // i >= 5 so it doesnt add he/she/name in Unit Overview Column 5
          
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn+3).getValue() + ' ');  // Appends He or She depending the gender
        
        } else if (random >= 89 && i >=5){   
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
          
        };
       }
     };
    text.appendText('\n');
    text.appendText('\n');
    }  //if KC
  
    positionRow++;
  } // ouuter for
    
  }     
  
  
  /////////////////////   YEAR 3 Group 1 3M  //////////////////////////
  
  function generateReportGrade33M() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
   
    // Returns the active cell
    var sheet = ss.getSheets()[3];
    var cell = ss.getActiveCell();
   //cell.setBackground("#ffceee");
    var positionRow = cell.getRow();
    var positionColumn = cell.getColumn()
   
    var firstName = sheet.getRange(positionRow,positionColumn).getValue();
    var lastName = sheet.getRange(positionRow,positionColumn+1).getValue();
    var gender = sheet.getRange(positionRow,positionColumn+3).getValue();
    var fullName = firstName + ' ' + lastName + ' ';
    
    
    // Open Document to putthe comments
    var doc = DocumentApp.openById('1sEAFVng5QHLN_0ETrHclN-sy6T12K8xHUf-i98ABa8M');  
  
    var body = doc.getBody();
    
    var text = body.editAsText();
  
  //Looping through comments
   for (var j = 1; j<=100; j++){  // This outer for loops through the ROWS of students
    
   if ((sheet.getRange(positionRow,positionColumn+2).getValue()) === '3M'){  //Checks only for KI students
    
    text.appendText(sheet.getRange(positionRow,positionColumn).getValue() + ' '); //Appends Name spaces
    for (var i = 4; i < 91; i++){  //This for loops through the COLUMNS for each student. 
      
      if (sheet.getRange(positionRow,positionColumn+i).getValue() != 'NA') {  //If it says NA skip it
          
        text.appendText(sheet.getRange(positionRow,positionColumn+i).getValue());  //Get the comment and append to report
        
        var random = Math.floor(Math.random() *100);   //The chances of appending the name rather than he/she is 89%
        
        if (i === 4) { //The first column will always display te name (after Unit Overview)
        
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
        
        }  else if (random < 89 && i >=5) {   // i >= 5 so it doesnt add he/she/name in Unit Overview Column 5
          
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn+3).getValue() + ' ');  // Appends He or She depending the gender
        
        } else if (random >= 89 && i >=5){   
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
          
        };
       }
     };
    text.appendText('\n');
    text.appendText('\n');
    }  //if KC
  
    positionRow++;
  } // ouuter for
    
  }     
   
  
  
  /////////////////////   YEAR 3 Group 2 3C  //////////////////////////
  
  function generateReportGrade33C() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
   
    // Returns the active cell
    var sheet = ss.getSheets()[3];
    var cell = ss.getActiveCell();
   //cell.setBackground("#ffceee");
    var positionRow = cell.getRow();
    var positionColumn = cell.getColumn()
   
    var firstName = sheet.getRange(positionRow,positionColumn).getValue();
    var lastName = sheet.getRange(positionRow,positionColumn+1).getValue();
    var gender = sheet.getRange(positionRow,positionColumn+3).getValue();
    var fullName = firstName + ' ' + lastName + ' ';
    
    
    // Open Document to putthe comments
    var doc = DocumentApp.openById('1OfXJgWNbxJoS8btBYXaXNGClovgVJLt0Bss8Lq0a1Wk');  
  
    var body = doc.getBody();
    
    var text = body.editAsText();
  
  //Looping through comments
   for (var j = 1; j<=100; j++){  // This outer for loops through the ROWS of students
    
   if ((sheet.getRange(positionRow,positionColumn+2).getValue()) === '3C'){  //Checks only for KI students
    
    text.appendText(sheet.getRange(positionRow,positionColumn).getValue() + ' '); //Appends Name spaces
    for (var i = 4; i < 91; i++){  //This for loops through the COLUMNS for each student. 
      
      if (sheet.getRange(positionRow,positionColumn+i).getValue() != 'NA') {  //If it says NA skip it
          
        text.appendText(sheet.getRange(positionRow,positionColumn+i).getValue());  //Get the comment and append to report
        
        var random = Math.floor(Math.random() *100);   //The chances of appending the name rather than he/she is 89%
        
        if (i === 4) { //The first column will always display te name (after Unit Overview)
        
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
        
        }  else if (random < 89 && i >=5) {   // i >= 5 so it doesnt add he/she/name in Unit Overview Column 5
          
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn+3).getValue() + ' ');  // Appends He or She depending the gender
        
        } else if (random >= 89 && i >=5){   
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
          
        };
       }
     };
    text.appendText('\n');
    text.appendText('\n');
    }  //if KC
  
    positionRow++;
  } // ouuter for
    
  }      
  
  
  
  
  /////////////////////   YEAR 3 Group 3  3B   //////////////////////////
  
  function generateReportGrade33B() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
   
    // Returns the active cell
    var sheet = ss.getSheets()[3];
    var cell = ss.getActiveCell();
   //cell.setBackground("#ffceee");
    var positionRow = cell.getRow();
    var positionColumn = cell.getColumn()
   
    var firstName = sheet.getRange(positionRow,positionColumn).getValue();
    var lastName = sheet.getRange(positionRow,positionColumn+1).getValue();
    var gender = sheet.getRange(positionRow,positionColumn+3).getValue();
    var fullName = firstName + ' ' + lastName + ' ';
    
    
    // Open Document to putthe comments
    var doc = DocumentApp.openById('1ZQUD6-pX8EFIEeVrCUaKmWFEv8xYTDK3NrdCJgZx4rQ');  
  
    var body = doc.getBody();
    
    var text = body.editAsText();
  
  //Looping through comments
   for (var j = 1; j<=100; j++){  // This outer for loops through the ROWS of students
    
   if ((sheet.getRange(positionRow,positionColumn+2).getValue()) === '3B'){  //Checks only for KI students
    
    text.appendText(sheet.getRange(positionRow,positionColumn).getValue() + ' '); //Appends Name spaces
    for (var i = 4; i < 91; i++){  //This for loops through the COLUMNS for each student. 
      
      if (sheet.getRange(positionRow,positionColumn+i).getValue() != 'NA') {  //If it says NA skip it
          
        text.appendText(sheet.getRange(positionRow,positionColumn+i).getValue());  //Get the comment and append to report
        
        var random = Math.floor(Math.random() *100);   //The chances of appending the name rather than he/she is 89%
        
        if (i === 4) { //The first column will always display te name (after Unit Overview)
        
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
        
        }  else if (random < 89 && i >=5) {   // i >= 5 so it doesnt add he/she/name in Unit Overview Column 5
          
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn+3).getValue() + ' ');  // Appends He or She depending the gender
        
        } else if (random >= 89 && i >=5){   
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
          
        };
       }
     };
    text.appendText('\n');
    text.appendText('\n');
    }  //if KC
  
    positionRow++;
  } // ouuter for
    
  }     
  
  
  ////////////////////   YEAR 4 Group 1  4N    //////////////////////////
  
  function generateReportGrade4N() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
   
    // Returns the active cell
    var sheet = ss.getSheets()[4];
    var cell = ss.getActiveCell();
   //cell.setBackground("#ffceee");
    var positionRow = cell.getRow();
    var positionColumn = cell.getColumn()
   
    var firstName = sheet.getRange(positionRow,positionColumn).getValue();
    var lastName = sheet.getRange(positionRow,positionColumn+1).getValue();
    var gender = sheet.getRange(positionRow,positionColumn+3).getValue();
    var fullName = firstName + ' ' + lastName + ' ';
    
    
    // Open Document to putthe comments
    var doc = DocumentApp.openById('1XrgBYdX7swPXCR56PyxN8673WLxTA0OtxgENNykcwF4');  
  
    var body = doc.getBody();
    
    var text = body.editAsText();
  
  //Looping through comments
   for (var j = 1; j<=100; j++){  // This outer for loops through the ROWS of students
    
   if ((sheet.getRange(positionRow,positionColumn+2).getValue()) === '4N'){  //Checks only for KI students
    
    text.appendText(sheet.getRange(positionRow,positionColumn).getValue() + ' '); //Appends Name spaces
    for (var i = 4; i < 104; i++){  //This for loops through the COLUMNS for each student. 
      
      if (sheet.getRange(positionRow,positionColumn+i).getValue() != 'NA') {  //If it says NA skip it
          
        text.appendText(sheet.getRange(positionRow,positionColumn+i).getValue());  //Get the comment and append to report
        
        var random = Math.floor(Math.random() *100);   //The chances of appending the name rather than he/she is 89%
        
        if (i === 4) { //The first column will always display te name (after Unit Overview)
        
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
        
        }  else if (random < 89 && i >=5) {   // i >= 5 so it doesnt add he/she/name in Unit Overview Column 5
          
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn+3).getValue() + ' ');  // Appends He or She depending the gender
        
        } else if (random >= 89 && i >=5){   
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
          
        };
       }
     };
    text.appendText('\n');
    text.appendText('\n');
    }  //if KC
  
    positionRow++;
  } // ouuter for
    
  }     
  
  
  
  
  ////////////////////   YEAR 4 Group 2  4H //////////////////////////
  
  function generateReportGrade4H() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
   
    // Returns the active cell
    var sheet = ss.getSheets()[4];
    var cell = ss.getActiveCell();
   //cell.setBackground("#ffceee");
    var positionRow = cell.getRow();
    var positionColumn = cell.getColumn()
   
    var firstName = sheet.getRange(positionRow,positionColumn).getValue();
    var lastName = sheet.getRange(positionRow,positionColumn+1).getValue();
    var gender = sheet.getRange(positionRow,positionColumn+3).getValue();
    var fullName = firstName + ' ' + lastName + ' ';
    
    
    // Open Document to putthe comments
    var doc = DocumentApp.openById('1oGeeV4bXN_8zDGnGxuYh70OkcYIPAcviMYTqZvDZjrA');  
  
    var body = doc.getBody();
    
    var text = body.editAsText();
  
  //Looping through comments
   for (var j = 1; j<=100; j++){  // This outer for loops through the ROWS of students
    
   if ((sheet.getRange(positionRow,positionColumn+2).getValue()) === '4H'){  //Checks only for KI students
    
    text.appendText(sheet.getRange(positionRow,positionColumn).getValue() + ' '); //Appends Name spaces
    for (var i = 4; i < 104; i++){  //This for loops through the COLUMNS for each student. 
      
      if (sheet.getRange(positionRow,positionColumn+i).getValue() != 'NA') {  //If it says NA skip it
          
        text.appendText(sheet.getRange(positionRow,positionColumn+i).getValue());  //Get the comment and append to report
        
        var random = Math.floor(Math.random() *100);   //The chances of appending the name rather than he/she is 89%
        
        if (i === 4) { //The first column will always display te name (after Unit Overview)
        
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
        
        }  else if (random < 89 && i >=5) {   // i >= 5 so it doesnt add he/she/name in Unit Overview Column 5
          
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn+3).getValue() + ' ');  // Appends He or She depending the gender
        
        } else if (random >= 89 && i >=5){   
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
          
        };
       }
     };
    text.appendText('\n');
    text.appendText('\n');
    }  //if KC
  
    positionRow++;
  } // ouuter for
    
  }     
  
  
  
  
  ////////////////////   YEAR 4 Group 3  4W   //////////////////////////
  
  function generateReportGrade4W() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
   
    // Returns the active cell
    var sheet = ss.getSheets()[4];
    var cell = ss.getActiveCell();
   //cell.setBackground("#ffceee");
    var positionRow = cell.getRow();
    var positionColumn = cell.getColumn()
   
    var firstName = sheet.getRange(positionRow,positionColumn).getValue();
    var lastName = sheet.getRange(positionRow,positionColumn+1).getValue();
    var gender = sheet.getRange(positionRow,positionColumn+3).getValue();
    var fullName = firstName + ' ' + lastName + ' ';
    
    
    // Open Document to putthe comments
    var doc = DocumentApp.openById('1WAcRtIzRAaa6MvAgwhyqTK6nDrdST9H2IkXVDVvzD6E');  
  
    var body = doc.getBody();
    
    var text = body.editAsText();
  
  //Looping through comments
   for (var j = 1; j<=100; j++){  // This outer for loops through the ROWS of students
    
   if ((sheet.getRange(positionRow,positionColumn+2).getValue()) === '4W'){  //Checks only for KI students
    
    text.appendText(sheet.getRange(positionRow,positionColumn).getValue() + ' '); //Appends Name spaces
    for (var i = 4; i < 104; i++){  //This for loops through the COLUMNS for each student. 
      
      if (sheet.getRange(positionRow,positionColumn+i).getValue() != 'NA') {  //If it says NA skip it
          
        text.appendText(sheet.getRange(positionRow,positionColumn+i).getValue());  //Get the comment and append to report
        
        var random = Math.floor(Math.random() *100);   //The chances of appending the name rather than he/she is 89%
        
        if (i === 4) { //The first column will always display te name (after Unit Overview)
        
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
        
        }  else if (random < 89 && i >=5) {   // i >= 5 so it doesnt add he/she/name in Unit Overview Column 5
          
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn+3).getValue() + ' ');  // Appends He or She depending the gender
        
        } else if (random >= 89 && i >=5){   
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
          
        };
       }
     };
    text.appendText('\n');
    text.appendText('\n');
    }  //if KC
  
    positionRow++;
  } // ouuter for
    
  }     
  
  ////////////////////   YEAR 5 Group 1   5M  //////////////////////////
  
  function generateReportGrade55M() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
   
    // Returns the active cell
    var sheet = ss.getSheets()[5];
    var cell = ss.getActiveCell();
   //cell.setBackground("#ffceee");
    var positionRow = cell.getRow();
    var positionColumn = cell.getColumn()
   
    var firstName = sheet.getRange(positionRow,positionColumn).getValue();
    var lastName = sheet.getRange(positionRow,positionColumn+1).getValue();
    var gender = sheet.getRange(positionRow,positionColumn+3).getValue();
    var fullName = firstName + ' ' + lastName + ' ';
    
    
    // Open Document to putthe comments
    var doc = DocumentApp.openById('1acY_I2g0dIk6FQkE7RoPMRg1i2j-D12mHG4mUhZD6kA');  
  
    var body = doc.getBody();
    
    var text = body.editAsText();
  
  //Looping through comments
   for (var j = 1; j<=100; j++){  // This outer for loops through the ROWS of students
    
   if ((sheet.getRange(positionRow,positionColumn+2).getValue()) === '5M'){  //Checks only for KI students
    
    text.appendText(sheet.getRange(positionRow,positionColumn).getValue() + ' '); //Appends Name spaces
    for (var i = 4; i < 95; i++){  //This for loops through the COLUMNS for each student. 
      
      if (sheet.getRange(positionRow,positionColumn+i).getValue() != 'NA') {  //If it says NA skip it
          
        text.appendText(sheet.getRange(positionRow,positionColumn+i).getValue());  //Get the comment and append to report
        
        var random = Math.floor(Math.random() *100);   //The chances of appending the name rather than he/she is 89%
        
        if (i === 4) { //The first column will always display te name (after Unit Overview)
        
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
        
        }  else if (random < 89 && i >=5) {   // i >= 5 so it doesnt add he/she/name in Unit Overview Column 5
          
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn+3).getValue() + ' ');  // Appends He or She depending the gender
        
        } else if (random >= 89 && i >=5){   
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
          
        };
       }
     };
    text.appendText('\n');
    text.appendText('\n');
    }  //if KC
  
    positionRow++;
  } // ouuter for
    
  }     
  
  
  
  
  ////////////////////   YEAR 5 Group 2   5R   //////////////////////////
  
  function generateReportGrade55R() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
   
    // Returns the active cell
    var sheet = ss.getSheets()[5];
    var cell = ss.getActiveCell();
   //cell.setBackground("#ffceee");
    var positionRow = cell.getRow();
    var positionColumn = cell.getColumn()
   
    var firstName = sheet.getRange(positionRow,positionColumn).getValue();
    var lastName = sheet.getRange(positionRow,positionColumn+1).getValue();
    var gender = sheet.getRange(positionRow,positionColumn+3).getValue();
    var fullName = firstName + ' ' + lastName + ' ';
    
    
    // Open Document to putthe comments
    var doc = DocumentApp.openById('12NjA36qh_y4xmaJZ_cEjvzqwf9yyTfrtgeof1ZqPEAM');  
  
    var body = doc.getBody();
    
    var text = body.editAsText();
  
  //Looping through comments
   for (var j = 1; j<=100; j++){  // This outer for loops through the ROWS of students
    
   if ((sheet.getRange(positionRow,positionColumn+2).getValue()) === '5R'){  //Checks only for KI students
    
    text.appendText(sheet.getRange(positionRow,positionColumn).getValue() + ' '); //Appends Name spaces
    for (var i = 4; i < 95; i++){  //This for loops through the COLUMNS for each student. 
      
      if (sheet.getRange(positionRow,positionColumn+i).getValue() != 'NA') {  //If it says NA skip it
          
        text.appendText(sheet.getRange(positionRow,positionColumn+i).getValue());  //Get the comment and append to report
        
        var random = Math.floor(Math.random() *100);   //The chances of appending the name rather than he/she is 89%
        
        if (i === 4) { //The first column will always display te name (after Unit Overview)
        
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
        
        }  else if (random < 89 && i >=5) {   // i >= 5 so it doesnt add he/she/name in Unit Overview Column 5
          
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn+3).getValue() + ' ');  // Appends He or She depending the gender
        
        } else if (random >= 89 && i >=5){   
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
          
        };
       }
     };
    text.appendText('\n');
    text.appendText('\n');
    }  //if KC
  
    positionRow++;
  } // ouuter for
    
  }     
  
  
  ////////////////////   YEAR 5 Group 3   5G   //////////////////////////
  
  function generateReportGrade55G() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
   
    // Returns the active cell
    var sheet = ss.getSheets()[5];
    var cell = ss.getActiveCell();
   //cell.setBackground("#ffceee");
    var positionRow = cell.getRow();
    var positionColumn = cell.getColumn()
   
    var firstName = sheet.getRange(positionRow,positionColumn).getValue();
    var lastName = sheet.getRange(positionRow,positionColumn+1).getValue();
    var gender = sheet.getRange(positionRow,positionColumn+3).getValue();
    var fullName = firstName + ' ' + lastName + ' ';
    
    
    // Open Document to putthe comments
    var doc = DocumentApp.openById('10yAvr44G1EO0WfOuzQ6DH3yYN8-rwECBt-22DO7X7_Y');  
  
    var body = doc.getBody();
    
    var text = body.editAsText();
  
  //Looping through comments
   for (var j = 1; j<=100; j++){  // This outer for loops through the ROWS of students
    
   if ((sheet.getRange(positionRow,positionColumn+2).getValue()) === '5G'){  //Checks only for KI students
    
    text.appendText(sheet.getRange(positionRow,positionColumn).getValue() + ' '); //Appends Name spaces
    for (var i = 4; i < 95; i++){  //This for loops through the COLUMNS for each student. 
      
      if (sheet.getRange(positionRow,positionColumn+i).getValue() != 'NA') {  //If it says NA skip it
          
        text.appendText(sheet.getRange(positionRow,positionColumn+i).getValue());  //Get the comment and append to report
        
        var random = Math.floor(Math.random() *100);   //The chances of appending the name rather than he/she is 89%
        
        if (i === 4) { //The first column will always display te name (after Unit Overview)
        
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
        
        }  else if (random < 89 && i >=5) {   // i >= 5 so it doesnt add he/she/name in Unit Overview Column 5
          
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn+3).getValue() + ' ');  // Appends He or She depending the gender
        
        } else if (random >= 89 && i >=5){   
          text.appendText(' ' + sheet.getRange(positionRow,positionColumn).getValue() + ' ');  //student name between spaces
          
        };
       }
     };
    text.appendText('\n');
    text.appendText('\n');
    }  //if KC
  
    positionRow++;
  } // ouuter for
    
  }     