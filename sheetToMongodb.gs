//
// exportToMongoDB
//
// By Heinz Baumann - May 2024
//
// This function will export either a specified language or all the translation to the
// translation collection.
// 
// initial implementation: all translations, specific language
// updates existing entries based on the langId or creates a new entry if no langId is not found

//
// Things that need to be donw:
// - clean up code and remove all the unncessary Debug stuff
// - add production access point and triggers

var debug = true;     // Debug logs enabled
var jsonString = "";  // global jsonString

function exportToMongoDB(strLanguage = null) {
  var langStrArray = [];
  var tempLang = null;

  // if we have no given language we read all the languages from the sheet and 
  // stuff it into the langStrArray.
  if (strLanguage !== null){
    langStrArray[0] = strLanguage;
  } else {
    langStrArray = getLanguages();
  }
  
  // process and write language after language
  tempLang = null;
  for (var k = 0; k < langStrArray.length; k++) {
    // write the language and the date stamp into the jsonString
    if (tempLang === null) {
      tempLang = langStrArray[k];
      // write the language section header
      jsonString = '{ "langId": "' + tempLang + '",';
      // add the lastUpdated field
      const d = new Date();
      jsonString += '"lastUpdated": "' + d.toISOString().substring(0,19) + "Z" + '",';
    }

    // process the purposes
    processPurposes(tempLang);

    // process the stacks
    processStacks(tempLang);
    
    // finally process the taxonomy
    processTaxonomies(tempLang);
    
    // terminate the language section and reset the tempLang string
    jsonString += "}";

/*
    // Debug output to check for specific JSON string errors
    Logger.log(jsonString.substring(5000, 6000));
    Logger.log(jsonString.substring(5865, 5885) + " chatAt: " + jsonString.charAt(5878) + " charCodeAt: " + jsonString.charCodeAt(5879) + " " + jsonString.charCodeAt(5879));
*/

    // Transmit to MongoDB
    // For the transaction we need to stringify the data and stuff in into a name: value
    // pair, where name is 'jstring' and the value is the JSON string
    // We will add another name value pair to handle updates where the name is id and the 
    // value is the actual mongo id returned from the original transaction.

    /*
    // debug data  
    var jStr = {
      langId: "en",
      purposes: {
        1: {
          id: 1,
          name: "Store and/or access information on a device",
          description: "Cookies, device or similiar...",
          illustration: [
            "Most purposes explained..."
          ]
        }
      }
    }
    */

    var jStr = JSON.parse(jsonString);
    var params = {
      'method': 'post',
      'muteHttpExceptions': true,
      'payload': { recordId: tempLang, jstring: JSON.stringify(jStr) }
      // 'payload': JSON.stringify(formData)
    };
  
    // write to mongoDB
    var result = UrlFetchApp.fetch('https://eu-central-1.aws.data.mongodb-api.com/app/gsheettomongo-jqpnesh/endpoint/SheetToMongo', params);

    Logger.log("DEBUG: LangId=" +tempLang + " result: " + result); 

    // Write the result value back into column H in the first row of the respective langauge section in the purpose sheet
    // Debug:
    // Logger.log("Posted to MongoDB: " + result.getContentText("UTF-8").replaceAll(/"/g,"").replaceAll(/\\/g, ""));
    var sheetName = "masterPurposesV2.2";
    var sheetRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getDataRange();
    var firstRaw = findFirstLanguageRow(sheetRange.getValues(), tempLang)
    sheetRange.getCell(firstRaw, 8).setValue(result.getContentText("UTF-8").replaceAll(/"/g,"").replaceAll(/\\/g, ""));

    jsonString = "";
    tempLang = null;
  }
}
//
// processPurposes(strLanguage)
//

function processPurposes(strLanguage) {
  var sheetName = "masterPurposesV2.2";
  var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var sheetRange = sheet1.getDataRange();
  var sheetValues = sheetRange.getValues();
  var workingRange = null;
  var workingValues = null;
  var bPurposeSection = false;
  var bSpecialPurposeSection = false;
  var bFeatureSection = false;
  var bSpecialFeatureSection = false;

  if (strLanguage !== null){
    var rangeArray = findRangeByValue(sheetValues, strLanguage);
    workingRange = sheet1.getRange(rangeArray[0], rangeArray[1], rangeArray[2], rangeArray[3]);
  } else {
    Logger.log("Error: processPurposes - No language specified");
    return;
  }

  workingValues = workingRange.getValues();

  if (false /* debug */) {  // debug
    if (workingRange != null) {
      Logger.log("rows found: " + workingRange.getNumRows());
      Logger.log("columns found: " + workingRange.getNumColumns());
      Logger.log("Range found: " + workingRange.getA1Notation());
      if (workingValues != null) {
        for (var i = 0; i < workingValues.length; i++)
          workingValues[i].forEach((element) => Logger.log(element));
      }
    } else {
      Logger.log("Value not found.");
    }
  } // debug

  if (workingValues != null) {
    for (var i = 0; i < workingValues.length; i++) {
      // if we have a purpose 
      if (workingValues[i][1] === "purpose") {
        if (!bPurposeSection) {
          jsonString += '"purposes": {';                   // purpose section
          bPurposeSection = true;
        }
        jsonString += '"' + workingValues[i][2] + '": {';  // get the purpose id
        jsonString += '"id": ' + workingValues[i][2] + ',';
        jsonString += '"name": "' + workingValues[i][3].replace(/\n/g,"") + '",';
        jsonString += '"description": "' + workingValues[i][4].replace(/\n/g,"") + '",';
        jsonString += '"illustrations": ["' + workingValues[i][5].replace(/\n/g,"").replace(/"/g, "\\\"") + '"';
        if (workingValues[i][6] !== "") 
          jsonString += ', "' + workingValues[i][6].replace(/\n/g,"").replace(/"/g, "\\\"") + '"]';
        else 
          jsonString += ']';
        jsonString += '}'
        if (workingValues[i+1][1] === "purpose") { // there are more purposes
          jsonString += ',';
        } else {                                   // we reached the end of the purpose section
          jsonString += '}';
          bPurposeSection = false;
        }
      } else if (workingValues[i][1] === "specialPurpose") {  // special purposes
        if (!bSpecialPurposeSection) {
          jsonString += ', "specialPurposes": {';             // special purpose section
          bSpecialPurposeSection = true;
        }
        jsonString += '"' + workingValues[i][2] + '": {';     // get the specialPurpose id
        jsonString += '"id": ' + workingValues[i][2] + ',';
        jsonString += '"name": "' + workingValues[i][3].replace(/\n/g,"") + '",';
        jsonString += '"description": "' + workingValues[i][4].replace(/\n/g,"") + '",';
        jsonString += '"illustrations": ["' + workingValues[i][5].replace(/\n/g,"").replace(/"/g, "\\\"") + '"';
        if (workingValues[i][6] !== "") 
          jsonString += ', "' + workingValues[i][6].replace(/\n/g,"").replace(/"/g, "\\\"") + '"]';
        else 
          jsonString += ']';
        jsonString += '}'
        if (workingValues[i+1][1] === "specialPurpose") {   // there are more special purposes
          jsonString += ',';
        } else {                          // we reached the end of the special purpose section
          jsonString += '}';
          bSpecialPurposeSection = false;
        } 
      } else if (workingValues[i][1] === "feature") {       // features
        if (!bFeatureSection) {
          jsonString += ', "features": {'                   // feature section
          bFeatureSection = true;
        }
        jsonString += '"' + workingValues[i][2] + '": {';   // get the feature id
        jsonString += '"id": ' + workingValues[i][2] + ',';
        jsonString += '"name": "' + workingValues[i][3].replace(/\n/g,"") + '",';
        jsonString += '"description": "' + workingValues[i][4].replace(/\n/g,"") + '",';
        jsonString += workingValues[i][5] !== "" ? '"illustrations": ["' + workingValues[i][5].replace(/\n/g,"").replace(/"/g, "\\\"") + '"' : '"illustrations": [ ';
        if (workingValues[i][6] !== "") 
          jsonString += ', "' + workingValues[i][6].replace(/\n/g,"").replace(/"/g, "\\\"") + '"]';
        else 
          jsonString += ']';
        jsonString += '}'
        if (workingValues[i+1][1] === "feature") {   // there are more features
          jsonString += ',';
        } else {                                     // we reached the end of the feature section
          jsonString += '}';
          bFeatureSection = false;
        }
      } else if (workingValues[i][1] === "specialFeature") {       // special features
        if (!bSpecialFeatureSection) {
          jsonString += ', "specialFeatures": {'                   // special feature section
          bSpecialFeatureSection = true;
        }
        jsonString += '"' + workingValues[i][2] + '": {';   // get the feature id
        jsonString += '"id": ' + workingValues[i][2] + ',';
        jsonString += '"name": "' + workingValues[i][3].replace(/\n/g,"") + '",';
        jsonString += '"description": "' + workingValues[i][4].replace(/\n/g,"") + '",';
        jsonString += workingValues[i][5] !== "" ? '"illustrations": ["' + workingValues[i][5].replace(/\n/g,"").replace(/"/g, "\\\"") + '"' : '"illustrations": [ ';
        if (workingValues[i][6] !== "") 
          jsonString += ', "' + workingValues[i][6].replace(/\n/g,"").replace(/"/g, "\\\"") + '"]';
        else 
          jsonString += ']';
        jsonString += '}'
        if (i+1 < workingValues.length && workingValues[i+1][1] === "specialFeature") {   // there are more special features
          jsonString += ',';
        } else {                                     // we reached the end of the feature section
          jsonString += '}';
          bSpecialFeatureSection = false;
        }
      }
    }
  }
}

//
// processStacks
// Switch the Stacks sheet
// Create a range for the language in hand
// Write the stack information into our json file
//
function processStacks(strLanguage) {
  var sheetName = "masterStacksV2.2";
  var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var sheetRange = sheet1.getDataRange();
  var sheetValues = sheetRange.getValues();
  var workingRange = null;
  var workingValues = null;

  if (strLanguage !== null){
    var rangeArray = findRangeByValue(sheetValues, strLanguage);
    workingRange = sheet1.getRange(rangeArray[0], rangeArray[1], rangeArray[2], rangeArray[3]);
  } else {
    Logger.log("Error: processPurposes - No language specified");
    return;
  }
  workingValues = workingRange.getValues();

  if (workingValues != null) {
    // write the section header
    jsonString += ',"stacks": {';
    var bfirstValueWritten = false;
    for (var i = 0; i < workingValues.length; i++) {
      jsonString += '"' + workingValues[i][1] + '": {';  // get the stack id
      jsonString += '"id": ' + workingValues[i][1] + ',';
      jsonString += '"purposes": [';
      bfirstValueWritten = false;
      for (var l = 4; l < 15; l++) {
        if (workingValues[i][l] === "X") {
          if (bfirstValueWritten) {
            jsonString += ', ';
          }
          jsonString += l - 3;
          bfirstValueWritten = true;
        }
      }
      jsonString += '],';
      jsonString += '"specialFeatures": [';
      bfirstValueWritten = false;
      for (var l = 15; l < 17; l++) {
        if (workingValues[i][l] === "X") {
          if (bfirstValueWritten) {
            jsonString += ', ';
          }
          jsonString += l - 14;
          bfirstValueWritten = true;
        }
      }
      jsonString += '],';
      jsonString += '"name": "' + workingValues[i][2].replace(/\n/g,"").replace(/\r/g,"") + '",';
      jsonString += '"description": "' + workingValues[i][3].replace(/\n/g,"").replace(/\r/g,"") + '"';
      jsonString += '}'
      if (i+1 !== workingValues.length)
        jsonString += ',';
    }
    // close the Stack section
    jsonString += '}';
  }
}
// 
// processTaxonomies
// - create range for language in hand
// - witch the sheet and process the taxonomy
// - write the taxonomies into the jsonString
//
function processTaxonomies(strLanguage) {
  var sheetName = "masterDataTaxonomyV2.2";
  var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var sheetRange = sheet1.getDataRange();
  var sheetValues = sheetRange.getValues();
  var workingRange = null;
  var workingValues = null;

  if (strLanguage === null){
    Logger.log("Error: processPurposes - No language specified");
    return;
  }

  var rangeArray = findRangeByValue(sheetValues, strLanguage);
  if (rangeArray[0] === undefined) {
    console.log("langId: " + strLanguage + " not found in Taxonomy sheet. Aborting.")
    return;
  }
  workingRange = sheet1.getRange(rangeArray[0], rangeArray[1], rangeArray[2], rangeArray[3]);

  workingValues = workingRange.getValues();

  if (workingValues != null) {
    // write the section header
    jsonString += ',"dataCategories": {';
    for (var i = 0; i < workingValues.length; i++) {
      jsonString += '"' + workingValues[i][1] + '": {';  // get the dataCategories id
      jsonString += '"id": ' + workingValues[i][1] + ',';
      jsonString += '"name": "' + workingValues[i][2].replace(/\n/g,"") + '",';
      jsonString += '"description": "' + workingValues[i][3].replace(/\n/g,"").replace(/\"/g,"\\\"") + '",';
      jsonString += '"vendorGuidance": "' + workingValues[i][4].replace(/\n/g,"").replace(/\"/g,"\\\"") + '"';
      jsonString += '}';
      if (i+1 !== workingValues.length)
        jsonString += ',';
    }
    // close the Stack section
    jsonString += '}';
  }
}

//
// Update translation in MongoDB by a specified language or all lanagauge in sheet
//
function updateRecordInMongoDB(strLanguage = null) {
  var langStrArray = [];
  var tempLang = null;

  // if we have no given language we read all the languages from the sheet and 
  // stuff it into the langStrArray.
  if (strLanguage !== null){
    langStrArray[0] = strLanguage;
  } else {
    langStrArray = getLanguages();
  }
  
  // process and write language after language
  tempLang = null;
  for (var k = 0; k < langStrArray.length; k++) {
    // write the language and the date stamp into the jsonString
    if (tempLang === null) {
      tempLang = langStrArray[k];
      // write the language section header
      jsonString = '{ "langId": "' + tempLang + '",';
      // add the lastUpdated field
      const d = new Date();
      jsonString += '"lastUpdated": "' + d.toISOString().substring(0,19) + "Z" + '",';
    }
    // process the purposes
    processPurposes(tempLang);

    // process the stacks
    processStacks(tempLang);
    
    // finally process the taxonomy
    processTaxonomies(tempLang);
    
    // terminate the language section and reset the tempLang string
    jsonString += "}";

    // Transmit to MongoDB
    var formData = JSON.parse(jsonString);

    // find the mongo DB record id in the sheet
    var sheetName = "masterPurposesV2.2";
    var sheetRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getDataRange();
    /* 
      We don';'t use the record id because that is not working through the MongoDB Script instead we look up the by language string

      var cell1 = sheetRange.getCell(findFirstLanguageRow(sheetRange.getValues(), tempLang), 8);
      if (cell1.isBlank()) {
        Logger.log("Error no recordId found in sheet, proceed with a new export.");
        break;
      } else {
    */
    if (findFirstLanguageRow(sheetRange.getValues(), tempLang) === -1) {
      Logger.log("Error " + tempLang + " not found in sheet. Abort.");
      break;
    } else {
    /*
      var recordId = cell1.getValue();
      Logger.log(recordId.toString());
    */
    // we really should use the returned id but so far the _id is not working to query against in MongoDB
    // so for now we use the language id
      recordId = tempLang;
    }
    
    // setup the props to update
    var params = {
      'method': 'post',
      'payload': { recordId: recordId.toString(), jstring: JSON.stringify(formData) }
    };
  
    // call the DB and update the record
    var resultId = UrlFetchApp.fetch('https://eu-central-1.aws.data.mongodb-api.com/app/gsheettomongo-jqpnesh/endpoint/SheetToMongo', params);

    // Write the result value back into column H in the first row of the respective langauge section in the purpose sheet
    // Debug:
    Logger.log("Posted to MongoDB: " + resultId.getContentText("UTF-8").replaceAll(/"/g,"").replaceAll(/\\/g, ""));
    var sheetName = "masterPurposesV2.2";
    var sheetRange = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName).getDataRange();
    var firstRaw = findFirstLanguageRow(sheetRange.getValues(), strLanguage)
    sheetRange.getCell(firstRaw, 8).setValue(resultId.getContentText("UTF-8").replaceAll(/"/g,"").replaceAll(/\\/g, ""));
    
    jsonString = "";
    tempLang = null;
  }
}

//
// import new translation file by a specified language
//
function importNewTranslationToSheet(strLanguage) {
  
}

//
// Import a single or multiple new entries into the table
// Example: translation for a new Special Prupose 3
//
function importSingleEntryToSheet(strLanguage = null, sectionString, entryNumber) {

}

//
// GUI function
//
// Add a custom menu to execute the load to MongoDB
//
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('IAB DB Connectivity')
      .addItem('Select a language to update...', 'menuItem1')
      .addItem('Update all', 'menuItem2')
      .addItem('Import Data from MongoDB', 'importFromMongoDb')
      .addToUi();
}

// export the selected language
function menuItem1() {
  const html = HtmlService.createHtmlOutputFromFile('page')
      .setWidth(350)
      .setHeight(150);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select a language to export')
}

function onSelectedItem(value) {
  updateRecordInMongoDB(value);
  SpreadsheetApp.getUi() // Or DocumentApp, SlidesApp or FormApp.   // need to style the error msg
     .alert("Language translation for " + value + " exported.");
}

// export all
function menuItem2() {
  updateRecordInMongoDB();
  SpreadsheetApp.getUi() // Or DocumentApp, SlidesApp or FormApp.   // need to style the error msg
     .alert('Exporting all languages!');
}

//
// Support  and test functions
//
function myTest() {
  exportToMongoDB("ar"); // "en"
}

function myTestUpdate() {
  updateRecordInMongoDB("en");
}

function getLanguages() {
  var sheetName = "masterPurposesV2.2";
  var sheet1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
  var values = sheet1.getRange(2,1, sheet1.getLastRow()).getValues();
  var langStrArray = [];
  var tempLang = null;
  var j = 0;
  for (var i = 0; i < values.length; i++) {
    if (tempLang === null && values[i][0] !== "") {
      langStrArray[j++] = tempLang = values[i][0];
    } else if (i+1 < values.length && values[i+1][0] !== tempLang) {
      tempLang = null;
    }
  }
  return langStrArray;
}

function findRangeByValue(values, strLanguage) {
  var rowArray = [];
  var i = j = 0;
  // var j = 0;

  // First find the given language rows
  for (i = 0; i < values.length; i++) {
    if (values[i][0] === strLanguage) {  // language is the first column
      rowArray[j++] = i+1;
    }
  }
  
  // Set the number of columns
  var numCol = values[0].length+1;

  // Find min and max row
  var minRow = rowArray[0];
  var maxRow = 0;
  for (i = 0; i < rowArray.length; i++) {
    if (rowArray[i] <= minRow)
        minRow = rowArray[i];
  }
  for (i = 0; i < rowArray.length; i++) {
    if (rowArray[i] >= maxRow)
        maxRow = rowArray[i];
  }

  // return range array minRow, minCol, maxRow, minCol
  return [minRow,
    1,
    maxRow + 1 - minRow,
    numCol];
}

function findFirstLanguageRow(values, strLanguage) {
  
  for (var i = 0; i < values.length; i++) {
    if (values[i][0] === strLanguage) {  // language is the first column
      return i+1;
    }
  }
  return -1;
}

// EOF