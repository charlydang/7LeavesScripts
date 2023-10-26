/////////////////////////////////////////////////////////////
// By: Charly - 09/21/2015
/////////////////////////////////////////////////////////////

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('- 7 Leaves -')
      .addItem('Clear Timecard', 'qsClearTimecardUpdated')
      //.addItem('Load Timecard', 'fillTimeCard') 
      .addItem('Load Timecard 2.0', 'fillTimeCardUpdated')
      .addItem('Load Exclusions', 'importRange')
      .addToUi();
}

////////////////////////////////////////////////////////////
// Menu call
////////////////////////////////////////////////////////////
function qsLoadTimecard() {    
  // Get range listed on Timecard. Range stored as double array.
  var sTimecard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Timecard");
  var rTimecard = sTimecard.getRange('B9:J999').getValues();  
  
  // Validation: Check for Total Tip Amount
  var totalTipAmt = sTimecard.getRange('F3').getValue();
  if (!qsIsNumeric(totalTipAmt)) {
    SpreadsheetApp.getUi().alert('Please enter Total Tips.');
    return;
  }
  
  // Get employee names and totol hours from data sheet
  var arrEmployees = qsDataGetEmployeeHours();
  var totalAllHours = qsDataGetAllTotalHours(arrEmployees);
  
  //fillTimeCard(arrEmployees, sTimecard);
  
      // Fill total hours on Timecard
      var fullnameTC;
      for(var i = 0; i <= arrEmployees.length-1; i++) {
        var indexTC = 0;
        while(indexTC <= rTimecard.length-1) {
          var fullnameTC = qsTimeCardGetFullName(indexTC, rTimecard);     
          var fullnameData = arrEmployees[i][0] + " " + arrEmployees[i][1];
          
          if (qsIsStringEqual(fullnameTC,fullnameData)) {    
          //if(!qsIsStringEmpty(fullnameData)){
            // Fill hours
            var totalHours = Number(arrEmployees[i][2]);
            Logger.log(totalHours);
            sTimecard.getRange("F"+(indexTC+9)).setValue(totalHours);
            rTimecard[indexTC][4] = totalHours;

            //Fill note
            var note = arrEmployees[i][3];
            sTimecard.getRange("J"+(indexTC+9)).setValue(note);
            note = note+"";
        
            // Fill tips amount
            var tipAmt = 0;
            //var note = rTimecard[indexTC][8] + "";
            if (!qsIsStringContains(note, "NT")) {
              tipAmt = totalTipAmt * (totalHours/totalAllHours);
            }        
            sTimecard.getRange("H"+(indexTC+9)).setValue(tipAmt);
        
            // Fill Total Gross:  (Rate * Hours) + (Tips Amt)
            var totalGross = (Number(rTimecard[indexTC][3]) * totalHours) + tipAmt;
            sTimecard.getRange("I"+(indexTC+9)).setValue(totalGross);
        
            indexTC = 9999;  // break from while loop
          }
          indexTC++;
        }
  }  
  
  // Alert names on POS data but not on Timecard
  var missingNames = qsTimecardGetMissingNames(rTimecard, arrEmployees);
  if (missingNames != "")    
    SpreadsheetApp.getUi().alert('WARNING! Missing from Timecard: \n' + missingNames);  
}


function qsTimecardGetMissingNames(rTimecard, arrEmployees) {
  var result = [];
  for(var i = 0; i <= arrEmployees.length-1; i++) {   
    var isExist = false;
    var indexTC = 0;
    while(indexTC <= rTimecard.length-1) {
      var fullnameTC = qsTimeCardGetFullName(indexTC, rTimecard);     
      var fullnameData = arrEmployees[i][0] + " " + arrEmployees[i][1];      
      if (qsIsStringEqual(fullnameTC, fullnameData)) {
        isExist = true;
        indexTC = 9999;
      }      
      indexTC++;
    }  
    
    if (isExist == false)
      result.push(fullnameData);
  }
  
  return result.toString();
}

// Return an array of arrays of employee (fname, lname, hrs) with hours worked from data dump. 
function qsDataGetEmployeeHours() {
  var result = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet().getSheets();  
  var rData = ss[ss.length-1].getRange('A1:H999').getValues();
  for(var i = 1; i < rData.length-1; i++) {  
    
    // Get employee name and total hours on same line.
    if (!qsIsStringEmpty(rData[i][1])) 
    {
      
      /***
      Control 1 -: Corporate 0070 - YMZ
      Control 2 -: Cypress 0070 - YMM
      Control 3 -: Fountain Valley 0070 - Y61
      Control 4 -: Irvine 0070 - Y60
      Control 5 -: Garden Grove (Euclid) 0070 - XZ0
      Control 6 -: Huntington Beach 0070 - XUD
      Control 7 -: Garden Grove (Westminster) 0070 - XUB
      Control 8 -: Chino Hills 0070 - XR0
      Control 9 -: Costa Mesa 0070 - XLD
      Control 10 -: Tustin 0070 - XLC
      Control 11 -: Alhambra 0070 - XD4
      Control 12 -: Artesia 0070 - X9C
      ***/            
      //if (qsIsStringEqual(rData[i][0], "XUB"))
      //{
        var fname = (rData[i][1]+"");
        var lname = (rData[i][3]+"");      
        var hrs = Number(rData[i][4]);
        var note = (rData[i][7])
        var number = Number(rData[i][5]);  
        //Logger.log(note);
        
        // Add to array
        var empArr = [fname, lname, hrs, note, number];
        result.push(empArr)
      //}
    }    
  } 
  
  return result;
}

////////////////////////////////////////////////////////////
// Menu call
////////////////////////////////////////////////////////////
function qsClearTimecard() {  
  var sTimecard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Timecard");
  sTimecard.getRange("B3").setValue("");  // Paycheck
  sTimecard.getRange("F3").setValue("");  // Total Tips
  sTimecard.getRange("H3").setValue("");  // Starting
  sTimecard.getRange("J3").setValue("");  // Ending
  
  // Clear previous data
  var rTimecard = sTimecard.getRange('A9:I999').getValues();  
  for(var i = 0; i < rTimecard.length-1; i++) {
    if (!qsIsStringEmpty(rTimecard[i][0])) {      
      
      sTimecard.getRange("A"+(i+9)).setValue("");  // Number
      sTimecard.getRange("B"+(i+9)).setValue("");  // First Name
      sTimecard.getRange("C"+(i+9)).setValue("");  // Last Name
      
      sTimecard.getRange("F"+(i+9)).setValue("");  // Rate
      sTimecard.getRange("H"+(i+9)).setValue("");  // Tips Amt
      sTimecard.getRange("I"+(i+9)).setValue("");  // Total Gross
      sTimecard.getRange("J"+(i+9)).setValue("");  // Note
    }
  }
}


function qsClearTimecardUpdated() {  
  var sTimecard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Timecard2.0");
  sTimecard.getRange("B3").setValue("");  // Paycheck
  sTimecard.getRange("F3").setValue("");  // Total Tips
  sTimecard.getRange("H3").setValue("");  // Starting
  sTimecard.getRange("L3").setValue("");  // Ending
  
  // Clear previous data
  var rTimecard = sTimecard.getRange('A9:I999').getValues();  
  for(var i = 0; i < rTimecard.length-1; i++) {
    if (!qsIsStringEmpty(rTimecard[i][1])) {      
      
      sTimecard.getRange("A"+(i+9)).setValue("");  // Number
      sTimecard.getRange("B"+(i+9)).setValue("");  // First Name
      sTimecard.getRange("C"+(i+9)).setValue("");  // Last Name
      sTimecard.getRange("D"+(i+9)).setValue("");  // 
      sTimecard.getRange("E"+(i+9)).setValue("");  // 
      sTimecard.getRange("F"+(i+9)).setValue("");  // Rate
      sTimecard.getRange("G"+(i+9)).setValue("");  // Last Name
      sTimecard.getRange("H"+(i+9)).setValue("");  // Tips Amt
      sTimecard.getRange("I"+(i+9)).setValue("");  // Total Gross
      sTimecard.getRange("J"+(i+9)).setValue("");  // Note
      sTimecard.getRange("K"+(i+9)).setValue("");  // 
      sTimecard.getRange("L"+(i+9)).setValue("");  // 
    }
  }
}

/////////////////////////////////////////////////////////////
// Related Helper functions
/////////////////////////////////////////////////////////////

function qsDataGetAllTotalHours(arrEmployees) {
  var result = 0;
  for(var i = 0; i <= arrEmployees.length-1; i++) {
    var note = arrEmployees[i][3]+"";
    if (qsIsNumeric(arrEmployees[i][2]) && !qsIsStringContains(note, "NT"))
        result = result + Number(arrEmployees[i][2]);
  }
  return result;
}

function qsTimeCardGetFullName(rowIndex, rTimecard) {
  var result;
  if (!qsIsStringEmpty(rTimecard[rowIndex][2]))
  {
    result =  (rTimecard[rowIndex][2] + "").trim();
  }
  else
  {
    //result = (rTimecard[rowIndex][0]+"").trim() + " " + (rTimecard[rowIndex][1]+"").trim();
    var fnameOnly = (rTimecard[rowIndex][0]+"").trim();
    fnameOnly = fnameOnly.split(' ')[0];
    result = fnameOnly + " " + (rTimecard[rowIndex][1]+"").trim();
  }
  
  return result;
}

/////////////////////////////////////////////////////////////
// Helper functions
/////////////////////////////////////////////////////////////

function qsIsStringEqual(obj1, obj2) {
   var s1 = (obj1+"").toLowerCase().trim();
   var s2 = (obj2+"").toLowerCase().trim();  
   return (s1 == s2);
}

function qsIsStringEmpty(obj) {
   var s1 = (obj+"").trim();
   return (s1 == "");
}

function qsIsStringContains(text, subtext) {
   if (text.toLowerCase().indexOf(subtext.toLowerCase()) > -1)
     return true;
  
   return false;
}

function qsIsNumeric(n) {
  return !isNaN(parseFloat(n)) && isFinite(n);
}

function qsArrayFirstUnique(array) {
  return array.filter(function (a, b, c) {
    // keeps first occurrence
    return c.indexOf(a) === b;
  });
}

//////////////////////////////////
// Charly Made Functions
//////////////////////////////////

function importRange() {
  const sourceSS = SpreadsheetApp.openById('1WcdoXYB5CoxAQA0Q7H3Lsh2LddHYpB0k3C3KS0vbzF8');
  const sourceRange = sourceSS.getRange('A1:A99');
  const sourceVals = sourceRange.getValues();

  var eRawData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Exclusion');
  var eTRawData = eRawData.getRange('A1:A99').getValues();

  eRawData.getRange('A1:A99').clear();

  eRawData.getRange('A1:A99').setValues(sourceVals);

}

function fillTimeCard(){
  // Get range listed on Timecard. Range stored as double array.
  var sTimecard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Timecard");
  var rTimecard = sTimecard.getRange('B9:J999').getValues();  
  
  // Get employee names and totol hours from data sheet
  var arrEmployees = qsDataGetEmployeeHours();
  var totalAllHours = qsDataGetAllTotalHours(arrEmployees);
  
  // Validation: Check for Total Tip Amount
  var totalTipAmt = sTimecard.getRange('F3').getValue();
  if (!qsIsNumeric(totalTipAmt)) {
    SpreadsheetApp.getUi().alert('Please enter Total Tips.');
    return;
  }
  
  for(var i = 0; i <=arrEmployees.length-1; i++ ){
        
    var fname = arrEmployees[i][0];
       sTimecard.getRange("B"+(i+9)).setValue(fname);
        
    var lname = arrEmployees[i][1];
        sTimecard.getRange("C"+(i+9)).setValue(lname);
        
    var number = arrEmployees[i][4]
        sTimecard.getRange("A"+(i+9)).setValue(number);
        
    var totalHours = Number(arrEmployees[i][2]);
        sTimecard.getRange("F"+(i+9)).setValue(totalHours);
        
    //Fill note
    var note = arrEmployees[i][3];
        sTimecard.getRange("J"+(i+9)).setValue(note);
        note = note+"";
        
    // Fill tips amount
    var tipAmt = 0;
    //var note = rTimecard[indexTC][8] + "";
    if (!qsIsStringContains(note, "NT")) {
    tipAmt = totalTipAmt * (totalHours/totalAllHours);
    }        
    sTimecard.getRange("H"+(i+9)).setValue(tipAmt);     
        
    // Fill Total Gross:  (Rate * Hours) + (Tips Amt)
    var totalGross = (Number(rTimecard[i][3]) * totalHours) + tipAmt;
    sTimecard.getRange("I"+(i+9)).setValue(totalGross);
  }
}

function fillTimeCardUpdated(){
  importRange();
  
  //Get range listed on Timecard. Range stored as double array.
  var sRawData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('RawData');
  var rRawData = sRawData.getRange('A8:H999').getValues();

  var eRawData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Exclusion');
  var eTRawData = sRawData.getRange('A1:A99').getValues();

  var sTimecard = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Timecard2.0");
  var rTimecard = sTimecard.getRange('A9:J999').getValues();  
  var totalTipHours = 0;  
  var calculateTips = [1];
  var counter = 9;
  var arrayCounter = 0;
  var totalTips = sTimecard.getRange('F3').getValue() ;

  var exlusionList = ["Start"]
  for(var i =1; i<=999; i++){
    var continueCondition = eRawData.getRange('A'+i).getValue();
    if(!qsIsStringEmpty(continueCondition)){
        exlusionList[i] = continueCondition
    }
    
  }

  console.log(exlusionList);

  var startingPoint = 5;

  if (!qsIsNumeric(totalTips)) {
    SpreadsheetApp.getUi().alert('Please enter Total Tips.');
    return;
  }

  for(var i=1; i <=rRawData.length; i++) {
    var breakCondition = sRawData.getRange('A'+(i)).getValue();
    if(breakCondition == 'Employee ID'){
      startingPoint = i;
      console.log('We broke successfully here!')
      console.log(startingPoint)
      break;
    }
  }

  
  for(var i = startingPoint+1; i <= rRawData.length; i++){
    var breakCondition = sRawData.getRange('A'+(i+1)).getValue();
    if(breakCondition == 'Report Total'){
      //console.log(i);
      break;
    }
    var employeeID = sRawData.getRange('A'+(i)).getValue(); 
    var fullName =  sRawData.getRange('B'+(i)).getValue();
    var role =  sRawData.getRange('D'+(i+1)).getValue();
    var regularHours =  sRawData.getRange('E'+(i)).getValue();
    var overTimeHours =  sRawData.getRange('F'+(i)).getValue();
    var holidayHours =  sRawData.getRange('G'+(i)).getValue();
    var totalHours =  sRawData.getRange('H'+(i)).getValue();
    
    
    totalHours = Number(regularHours)+(overTimeHours) + Number(holidayHours) + Number(totalHours);

    if(!qsIsStringEmpty(employeeID)) {
      sTimecard.getRange('A'+counter).setValue(employeeID)
      sTimecard.getRange('B'+counter).setValue(fullName)
      sTimecard.getRange('E'+counter).setValue(role)
      sTimecard.getRange('F'+counter).setValue(regularHours)
      sTimecard.getRange('G'+counter).setValue(overTimeHours)
      sTimecard.getRange('H'+counter).setValue(holidayHours)
      sTimecard.getRange('I'+counter).setValue(totalHours)
      totalTipHours += parseFloat(totalHours);  
        if(exlusionList.includes(fullName)){
          calculateTips[arrayCounter] = counter;
          arrayCounter +=1;
          sTimecard.getRange('L'+counter).setValue("NT")
          totalTipHours -= parseFloat(totalHours);
        }
      counter+=1;
    }
  } 

  var totalTipHoursAmt = parseFloat(totalTipHours.toFixed(2));
  console.log('Total Tip Hours:' + totalTipHoursAmt)

    for(var i = 9; i <= counter; i++){
      if(!calculateTips.includes(i)){
        var totalHours = sTimecard.getRange('I'+i).getValue();
        console.log("Total Hour for this person: " + totalHours);
        var denominator = totalHours / totalTipHoursAmt
        console.log("TotalTips: " + totalTips);
        sTimecard.getRange('J'+i).setValue(Number(totalTips) * Number(denominator) );
      }
    }

}


