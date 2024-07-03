//
// Read from MongoDB and write out into the 3 different sheets
//
function importFromMongoDb() {
  var getData = UrlFetchApp.fetch('https://eu-central-1.aws.data.mongodb-api.com/app/gsheettomongo-jqpnesh/endpoint/MongoToSheet').getContentText();
  var jsonObj = JSON.parse(getData);

  Logger.log(jsonObj.length + " records found and starting to import...");

  // scan the json and write into the 3 different sheets: purposes, stacks and data taxonomies
  // scan through the array, process lanuage by langauge
  var langIdArray = [];
  var proceed = true;
  var rowNum = 1;
  var cellCnt = 1;
  var shIndex = SpreadsheetApp.getActiveSpreadsheet().getSheets().length;
  var k = 1;
  var m = 0;
  var recordsImported = 0

  for (i = 0; i < jsonObj.length; i++) {
    proceed = true;
    // check for the language and whether it is a duplicate
    if (langIdArray.length > 0) {
      for (j = 0; j < langIdArray.length; j++) {
        // found a duplicate (we print a warning and ignore that entry)
        if (langIdArray[j] === jsonObj[i].langId) { 
          Logger.log("Warning: Duplicate langauge " + jsonObj[i].langId + " found. Record will be ignored.");
          proceed = false;
        }
      }
    } 
    
    if (proceed) {
      recordsImported++;
      langIdArray[i] = jsonObj[i].langId;

      // Purpose, special purpose, features & special features
      //
      // language	
      // type	
      // Id	
      // title	
      // description	
      // illustration[0]	
      // illustration[1]
      var purposeSheetName = "masterPurposesV2.2_NEW";
      var sh1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(purposeSheetName);
      if (sh1 === null) { // create sheet
        sh1 = SpreadsheetApp.getActive().insertSheet(purposeSheetName, shIndex++);
      }

      // write the header into the sheet
      rowNum = 1;
      if (sh1.getRange(rowNum,1).isBlank()) {
        sh1.getRange(rowNum,1, 1, 8).setFontWeight("bold");
        sh1.getRange(rowNum,1).setValue("language");
        sh1.getRange(rowNum,2).setValue("type");
        sh1.getRange(rowNum,3).setValue("id");
        sh1.getRange(rowNum,4).setValue("title");
        sh1.getRange(rowNum,5).setValue("description");
        sh1.getRange(rowNum,6).setValue("illustration[0]");
        sh1.getRange(rowNum,7).setValue("illustration[1]");
      }

      // get last rowNum and set our start row to populate the content
      rowNum = sh1.getLastRow();
      rowNum++;

      // purposes
      for (k = 1; k < 12; k++) {
        sh1.getRange(rowNum, 1).setValue(jsonObj[i].langId);
        sh1.getRange(rowNum, 2).setValue("purpose");
        sh1.getRange(rowNum, 3).setValue(jsonObj[i].purposes[k].id);
        sh1.getRange(rowNum, 4).setValue(jsonObj[i].purposes[k].name);
        sh1.getRange(rowNum, 5).setValue(jsonObj[i].purposes[k].description);
        sh1.getRange(rowNum, 6).setValue(jsonObj[i].purposes[k].illustrations[0]);
        if (jsonObj[i].purposes[k].illustrations.length > 1)
          sh1.getRange(rowNum, 7).setValue(jsonObj[i].purposes[k].illustrations[1]);
        rowNum++;
      }

      // special purposes
      for (k = 1; k < 3; k++) {
        sh1.getRange(rowNum, 1).setValue(jsonObj[i].langId);
        sh1.getRange(rowNum, 2).setValue("specialPurpose");
        sh1.getRange(rowNum, 3).setValue(jsonObj[i].specialPurposes[k].id);
        sh1.getRange(rowNum, 4).setValue(jsonObj[i].specialPurposes[k].name);
        sh1.getRange(rowNum, 5).setValue(jsonObj[i].specialPurposes[k].description);
        sh1.getRange(rowNum, 6).setValue(jsonObj[i].specialPurposes[k].illustrations[0]);
        if (jsonObj[i].specialPurposes[k].illustrations.length > 1)
          sh1.getRange(rowNum, 7).setValue(jsonObj[i].specialPurposes[k].illustrations[1]);
        rowNum++;
      }

      // features
      for (k = 1; k < 4; k++) {
        sh1.getRange(rowNum, 1).setValue(jsonObj[i].langId);
        sh1.getRange(rowNum, 2).setValue("feature");
        sh1.getRange(rowNum, 3).setValue(jsonObj[i].features[k].id);
        sh1.getRange(rowNum, 4).setValue(jsonObj[i].features[k].name);
        sh1.getRange(rowNum, 5).setValue(jsonObj[i].features[k].description);
        sh1.getRange(rowNum, 6).setValue(jsonObj[i].features[k].illustrations[0]);
        if (jsonObj[i].features[k].illustrations.length > 1)
          sh1.getRange(rowNum, 7).setValue(jsonObj[i].features[k].illustrations[1]);
        rowNum++;
      }

      // special features
      for (k = 1; k < 3; k++) {
        sh1.getRange(rowNum, 1).setValue(jsonObj[i].langId);
        sh1.getRange(rowNum, 2).setValue("specialFeature");
        sh1.getRange(rowNum, 3).setValue(jsonObj[i].specialFeatures[k].id);
        sh1.getRange(rowNum, 4).setValue(jsonObj[i].specialFeatures[k].name);
        sh1.getRange(rowNum, 5).setValue(jsonObj[i].specialFeatures[k].description);
        sh1.getRange(rowNum, 6).setValue(jsonObj[i].specialFeatures[k].illustrations[0]);
        if (jsonObj[i].specialFeatures[k].illustrations.length > 1)
          sh1.getRange(rowNum, 7).setValue(jsonObj[i].specialFeatures[k].illustrations[1]);
        rowNum++;
      }

      // stacks
      // language	id
      // name	
      // description	
      // v2.consent.1	v2.consent.2	v2.consent.3	v2.consent.4	v2.consent.5	v2.consent.6	v2.consent.7	
      // v2.consent.8	v2.consent.9	v2.consent.10	v2.consent.11	
      // v2.specialFeature.1	v2.specialFeature.2

      var stackSheetName = "masterStacksV2.2_NEW";
      var sh2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(stackSheetName);
      if (sh2 === null) { // create sheet
        sh2 = SpreadsheetApp.getActive().insertSheet(stackSheetName, shIndex++);
      }
      // write the header into the sheet
      rowNum = 1;
      if (sh2.getRange(rowNum,1).isBlank()) {
        sh2.getRange(rowNum,1, 1, 17).setFontWeight("bold");
        sh2.getRange(rowNum,1).setValue("language");
        sh2.getRange(rowNum,2).setValue("id");
        sh2.getRange(rowNum,3).setValue("name");
        sh2.getRange(rowNum,4).setValue("description");
        sh2.getRange(rowNum,5).setValue("v2.consent.1");
        sh2.getRange(rowNum,6).setValue("v2.consent.2");
        sh2.getRange(rowNum,7).setValue("v2.consent.3");
        sh2.getRange(rowNum,8).setValue("v2.consent.4");
        sh2.getRange(rowNum,9).setValue("v2.consent.5");
        sh2.getRange(rowNum,10).setValue("v2.consent.6");
        sh2.getRange(rowNum,11).setValue("v2.consent.7");
        sh2.getRange(rowNum,12).setValue("v2.consent.8");
        sh2.getRange(rowNum,13).setValue("v2.consent.9");
        sh2.getRange(rowNum,14).setValue("v2.consent.10");
        sh2.getRange(rowNum,15).setValue("v2.consent.11");
        sh2.getRange(rowNum,16).setValue("v2.specialFeature.1");
        sh2.getRange(rowNum,17).setValue("v2.specialFeature.2");
      }

      // get last rowNum and set our start row to populate the content
      rowNum = sh2.getLastRow();
      rowNum++;
      
      for (k = 1; k < 46; k++) {
        sh2.getRange(rowNum, 1).setValue(jsonObj[i].langId);
        sh2.getRange(rowNum, 2).setValue(jsonObj[i].stacks[k].id);
        sh2.getRange(rowNum, 3).setValue(jsonObj[i].stacks[k].name);
        sh2.getRange(rowNum, 4).setValue(jsonObj[i].stacks[k].description);

        // convert the purpose and special feature object and write into the individual cells
        if (jsonObj[i].stacks[k].purposes.length != 0) {
          for (m = 0; m < jsonObj[i].stacks[k].purposes.length; m++) {
            switch (jsonObj[i].stacks[k].purposes[m]) {
              case 1:
                sh2.getRange(rowNum, 5).setValue("X");
                break;
              case 2: 
                sh2.getRange(rowNum, 6).setValue("X");
                break;
              case 3:
                sh2.getRange(rowNum, 7).setValue("X");
                break;
              case 4:
                sh2.getRange(rowNum, 8).setValue("X");
                break;
              case 5:
                sh2.getRange(rowNum, 9).setValue("X");
                break;
              case 6:
                sh2.getRange(rowNum, 10).setValue("X");
                break;
              case 7:
                sh2.getRange(rowNum, 11).setValue("X");
                break;
              case 8:
                sh2.getRange(rowNum, 12).setValue("X");
                break;
              case 9:
                sh2.getRange(rowNum, 13).setValue("X");
                break;
              case 10:
                sh2.getRange(rowNum, 14).setValue("X");
                break;
              case 11:
                sh2.getRange(rowNum, 15).setValue("X");
                break;
              default:
                Logger.log("Error: value out of range");
                break;
            }
          }
        }
        if (jsonObj[i].stacks[k].specialFeatures.length != 0) {
          for (m = 0; m < jsonObj[i].stacks[k].specialFeatures.length; m++) {
            switch (jsonObj[i].stacks[k].specialFeatures[m]) {
              case 1:
                sh2.getRange(rowNum, 16).setValue("X");
                break;
              case 2: 
                sh2.getRange(rowNum, 17).setValue("X");
                break;
              default:
                Logger.log("Error: value out of range");
                break;
            }
          }
        }
        rowNum++;
      }

      // end of stack

      // DataTaxonomies
      // column 1 - langId
      // column 2 - id
      // column 3 - categories
      // column 4 - description	
      // column 5 - vendorGuidance

      var taxonomySheetName = "masterDataTaxonomyV2.2_NEW";
      var sh3 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(taxonomySheetName);
      if (sh3 === null) { // create sheet
        sh3 = SpreadsheetApp.getActive().insertSheet(taxonomySheetName, shIndex);
      }
      // write header into the sheet
      rowNum = 1;
      if (sh3.getRange(rowNum,1).isBlank()) {
        sh3.getRange(rowNum,1, 1, 5).setFontWeight("bold");
        sh3.getRange(rowNum,1).setValue("LangId");
        sh3.getRange(rowNum,2).setValue("Id");
        sh3.getRange(rowNum,3).setValue("Categories");
        sh3.getRange(rowNum,4).setValue("Description");
        sh3.getRange(rowNum,5).setValue("VendorGuidance");
      } 

      // get last rowNum and set our start row to populate the content
      rowNum = sh3.getLastRow();
      rowNum++;
      for (k = 1; k < 12; k++) {
        sh3.getRange(rowNum, 1).setValue(jsonObj[i].langId);
        sh3.getRange(rowNum, 2).setValue(jsonObj[i].dataCategories[k].id);
        sh3.getRange(rowNum, 3).setValue(jsonObj[i].dataCategories[k].name);
        sh3.getRange(rowNum++, 4).setValue(jsonObj[i].dataCategories[k].description);
        // sh3.getRange(rowNum, 5).setValue(jsonObj[i].dataCategories[k].vendorGuidance);  // enable when we output the vendorGuidance
      }
      // end of masterDataTaxonomy
    }

  }  // end of jsonObj.length for loop
  Logger.log("All Done. Imported " + recordsImported + " of " + jsonObj.length + ". Please check the sheets with suffix _NEW.")
}