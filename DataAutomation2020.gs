function onOpen() 
{    
  var ui = SpreadsheetApp.getUi();
  ui.createMenu(" -- 7 Leaves -- ")
      .addItem("Update Datasource", "updateRefData")
      .addItem("Create Sheets", "createSpreadSheets")
      .addToUi();
}

function createSpreadSheets() 
{
  var refSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("_RefData");
  var dataFolderID = refSheet.getRange("B1").getValue();
  var dateField = refSheet.getRange("H1").getValue();
  var newDate = refSheet.getRange("H2").getValue();
  // update all empty sheet ids (by stores) -----------------------
  var dataFolder = DriveApp.getFolderById(dataFolderID);
  var dataFiles = dataFolder.getFiles();  
 
  // fetch one time for performance
  var storesRange = refSheet.getRange("B1:B50").getValues();
  var sheetIdsRange = refSheet.getRange("A5:O5").getValues()[0];
  var sheetIds = refSheet.getRange("A1:O30").getValues();
  
  while (dataFiles.hasNext())              // loop thru all files and match files with name of month
  {
    var dataFile = dataFiles.next();
    var fileName = dataFile.getName();
    var fileStoreName = fileName.split("_")[0];
    var fileDate = fileName.split("_")[2];    
    
    // only process specific month
    if (fileDate == dateField)
    {

      var file = DriveApp.getFileById("1GbWq_oJFTMmFwJYJ6KQaXIEDfdYc_1zWnjV6PyrLRPw").getName()
      var folder = DriveApp.getFolderById(dataFolderID)
      var newFile = file.makeCopy(file, folder)
      newFile.setName(fileStoreName.concat('_',newDate))

    }
  }
  return;
}


function updateRefData() 
{
  var refSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("_RefData");
  var dataFolderID = refSheet.getRange("B1").getValue();
  var dateField = refSheet.getRange("H1").getValue();

  // update all empty sheet ids (by stores) -----------------------
  var dataFolder = DriveApp.getFolderById(dataFolderID);
  var dataFiles = dataFolder.getFiles();  
 
  // fetch one time for performance
  var storesRange = refSheet.getRange("B1:B50").getValues();
  var sheetIdsRange = refSheet.getRange("A5:O5").getValues()[0];
  var sheetIds = refSheet.getRange("A1:O50").getValues();
  
  while (dataFiles.hasNext())              // loop thru all files and match files with name of month
  {
    var dataFile = dataFiles.next();
    var fileName = dataFile.getName();
    var fileStoreName = fileName.split("_")[0];
    var fileDate = fileName.split("_")[2];    
    
    // only process specific month
    if (fileDate == dateField)
    {
      var rowIndex = findMatchRow(fileStoreName, storesRange);
      var colIndex = findMatchCol(fileDate, sheetIdsRange);
      
      // update data sheet id
      if (rowIndex >= 0 && colIndex >= 0) 
      {
        var cellValue = sheetIds[rowIndex][colIndex];
        if (isEmptyNull_(cellValue)) {
          refSheet.getRange(rowIndex+1, colIndex+1).setValue(dataFile.getId());
        }
      }
    }
  }
  return;
}

function findMatchRow(searchValue, rangeValues) 
{ 
  for (var i = 0; i < rangeValues.length; i++)
  {
    var s1 = rangeValues[i];
    if (rangeValues[i].indexOf(searchValue) >= 0){
      return i; }
  }
  return -1;
}

function findMatchCol(searchValue, colValues) 
{ 
  for (var i = 0; i < colValues.length; i++)
  {
    var s1 = colValues[i];
    if (colValues[i].indexOf(searchValue) >= 0){
      return i; }
  }
  return -1;
}

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// generate sheets 
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

// by stores
function createSheetsStores() 
{
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var refSheet = ss.getSheetByName("_RefData"); 

  // create new sheet for each store, skipping existing 
  var list = refSheet.getRange("range_storeabrvnames").getValues();
  for (var i = 0; i < list.length; i++)
  {
    var storeName = list[i][1];
    var storeAbrv = list[i][0];    
    if (!ss.getSheetByName(storeAbrv))
    {
      if (!isEmptyNull_(storeAbrv))
      {
        var newSheet = ss.getSheetByName("_Template").copyTo(ss);
        newSheet.setName(storeAbrv);
        newSheet.showSheet();
        
        // set values in cells needed to activate sheet data
        newSheet.getRange("B4").setValue(storeName);
      }
    }
  }
  return;
}


/**
// by days
function createSheetsDays() 
{
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var refSheet = ss.getSheetByName("_RefData"); 
  var dataDate = new Date(refSheet.getRange("range_datadate").getValues());

  // create new sheet for each day, skipping existing sheets (matching day)
  var daysCount = new Date(dataDate.getYear(), dataDate.getMonth(), 0).getDate();  
  for (var i = 1; i <= daysCount; i++)
  {
    var sheetDate = new Date(dataDate.getYear(), dataDate.getMonth(), i);
    var sheetName = (sheetDate.getMonth()+1) + "/" + sheetDate.getDate();
    if (!ss.getSheetByName(sheetName))
    {
      var newSheet = ss.getSheetByName("_Template").copyTo(ss);
      newSheet.setName(sheetName);
      newSheet.showSheet();
            
      // set values in cells needed to activate sheet data
      newSheet.getRange("B4").setValue(sheetName);      // for code behind
    }
  }
  return;
}
**/

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// clear out all store sheets
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function clearAll() 
{
  var ss = SpreadsheetApp.getActiveSpreadsheet();  
  var sheets = ss.getSheets();
  sheets.forEach(delSheet_);
}

function delSheet_(sheet)
{
  var nodeleting = ["_RefData", "_Template", "By Date", "All"];
  if (nodeleting.indexOf(sheet.getName()) < 0)
  {
    SpreadsheetApp.getActiveSpreadsheet().deleteSheet(sheet);
  }
}   
  

///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// utilities / helpers
///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function isEmptyNull_(obj) 
{
   var s1 = (obj+"").trim();
   return (s1 == "");
}




