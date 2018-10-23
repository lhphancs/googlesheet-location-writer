function displayMsg(msg){
  Browser.msgBox(msg);
}

var RowColCoordinate = function(row, col){
  this.rowIndex = row;
  this.colIndex = col;
};

var SheetInfo = function(sheet){
  this.title = sheet.getName();
  
  var rangeData = sheet.getDataRange();
  this.amtRow = rangeData.getLastRow();
  this.amtCol = rangeData.getLastColumn();
  this.sheetValues = sheet.getRange(1, 1, this.amtRow, this.amtCol).getValues(); //Retrives values as 2d array
}

var ErrorMsgsContainer = function(sheet){
  this.errorMgs = [];
  
  this.addError = function(msg){
    this.errorMgs.push(msg);
  }
  
  this.getErrorMsgs = function(){
    var msg = "";
    for(var i = 0; i < this.errorMgs.length; ++i)
      msg += this.errorMgs[i] + '\\n';
    return msg;
  }
  
  this.displayErrorMsgs = function(optionalMsgAtStart){
    var finalErrorMsg = optionalMsgAtStart == undefined ? "": optionalMsgAtStart;
    finalErrorMsg += this.getErrorMsgs();
    
    displayMsg(finalErrorMsg);
  }
}

function getSheetErrorString(sheetTitle, msg){
  return sheetTitle + ": " + msg;
}

function addUndefinedHeaderErrors(sheetInfo, errorMsgsContainer, arrayOfHeadersNotFound){
  var headersNotFoundStr = "";
  for each(var header in arrayOfHeadersNotFound)
    headersNotFoundStr += header + ', ';
  headersNotFoundStr = headersNotFoundStr.substring(0, headersNotFoundStr.length - 2); // Remove ending space and comma
  
  var errorMsg = getSheetErrorString(sheetInfo.title, "These headers inside bracket were not found\\n[  " + headersNotFoundStr + "  ]"); 
  errorMsgsContainer.addError(errorMsg);
}

function getRowColCoordinateOfStr(sheetInfo, str){
  for(i = 0; i < sheetInfo.amtRow; ++i){
    for(j = 0; j < sheetInfo.amtCol; ++j){
      var cellVal = sheetInfo.sheetValues[i][j];
      if( typeof(cellVal) != 'string' )
        continue;
      
      if(cellVal.toUpperCase() == str.toUpperCase() )
        return new RowColCoordinate(i, j);
    }
  }
  return undefined;
}

function getDataFromRowIndex(sheetValues, rowIndex, dictOfValidHeadersRowColCoordinate){
  var data = {};
  for(var header in dictOfValidHeadersRowColCoordinate){
    var headerCoordinate = dictOfValidHeadersRowColCoordinate[header];
    data[header] = sheetValues[rowIndex][headerCoordinate.colIndex];
  }
  return data;
}

function readSheetValuesToCompleteDataDict(sheetInfo, dataDict, headerKeyCoordinate, dictOfValidHeadersRowColCoordinate){
  var sheetValues = sheetInfo.sheetValues;
  
  for(i = headerKeyCoordinate.rowIndex+1; i < sheetInfo.amtRow; ++i){
    var keyCellVal = sheetValues[i][headerKeyCoordinate.colIndex];
    if(keyCellVal != "")
      dataDict[keyCellVal] = getDataFromRowIndex(sheetValues, i, dictOfValidHeadersRowColCoordinate);
  }
}

function getDictWithValidValuesOnly(dict, errorMsgsContainer, sheetTitle){
  var retDict = {};
  for(var key in dict)
    if(dict[key] == undefined)
      errorMsgsContainer.addError( getSheetErrorString(sheetTitle, "'" + key + "' was not found") );
    else
      retDict[key] = dict[key];
  return retDict;
}

function getDataDict(wholesaleSpreadSheet, headerKey, writeHeaders, errorMsgsContainer){
  var dataDict = {};
  var sheets = wholesaleSpreadSheet.getSheets();
  
  for(var i = 0; i<sheets.length; ++i){
    var sheetInfo = new SheetInfo( sheets[i] );
    var headerKeyCoordinate = getRowColCoordinateOfStr(sheetInfo, headerKey);
    
    if(headerKeyCoordinate == undefined)
      errorMsgsContainer.addError( getSheetErrorString(sheetInfo.title, headerKey + "' was not found in sheet.") );
    else{
      var dictOfHeadersRowColCoordinate = getDictOfCoordinates(sheetInfo, writeHeaders);
      var dictOfValidHeadersRowColCoordinate = getDictWithValidValuesOnly(dictOfHeadersRowColCoordinate, errorMsgsContainer, sheetInfo.title);
      readSheetValuesToCompleteDataDict(sheetInfo, dataDict, headerKeyCoordinate, dictOfValidHeadersRowColCoordinate);
    }
  }
  return dataDict;
}

function getValueFromDictWithKeyAndHeader(wholesaleDataDict, keyCellVal, header){
  if(header in wholesaleDataDict[keyCellVal]){
    return wholesaleDataDict[keyCellVal][header];
  }
  return undefined;
}

function writeLocation(writeSheet, writeSheetInfo, wholesaleDataDict, writeKeyRowColCoordinate, dictOfWriteHeadersRowColCoordinate){
  var writeKeyColIndex = writeKeyRowColCoordinate.colIndex;
  var sheetValues = writeSheetInfo.sheetValues;
  
  // rowIndex has + 1 because we want to skip the header
  for(rowIndex = writeKeyRowColCoordinate.rowIndex + 1; rowIndex < writeSheetInfo.amtRow; ++rowIndex){
    var keyCellVal = sheetValues[rowIndex][writeKeyColIndex];
    if(keyCellVal in wholesaleDataDict){
      for(var header in dictOfWriteHeadersRowColCoordinate){
        var writeLocationCol = dictOfWriteHeadersRowColCoordinate[header].colIndex + 1;
        var writeVal = getValueFromDictWithKeyAndHeader(wholesaleDataDict, keyCellVal, header);
        writeSheet.getRange(rowIndex+1, writeLocationCol).setValue(writeVal);
      }
    }
  }
}

function getDictOfCoordinates(sheetInfo, strs){
  var dictOfCoordinates = {};
  for each(var str in strs)
    dictOfCoordinates[str] = getRowColCoordinateOfStr(sheetInfo, str);
  
  return dictOfCoordinates;
}

function getArrayOfUndefinedHeaders(dictOfHeadersRowColCoordinate){
  var arrayOfUndefinedHeaders = [];
  for(var header in dictOfHeadersRowColCoordinate){
    if(dictOfHeadersRowColCoordinate[header] == undefined)
      arrayOfUndefinedHeaders.push(header);
  }
  return arrayOfUndefinedHeaders;
}

function allHeadersAreFound(sheetInfo, dictOfWriteHeadersRowColCoordinate, errorMsgsContainer){
  var allHeadersAreFound = true;
  var arrayOfUndefinedHeaders = getArrayOfUndefinedHeaders(dictOfWriteHeadersRowColCoordinate);
      
  if(arrayOfUndefinedHeaders.length != 0){
    addUndefinedHeaderErrors(sheetInfo, errorMsgsContainer, arrayOfUndefinedHeaders);
    allHeadersAreFound = false;
  }
    
  return allHeadersAreFound;
}

function main() {
  const wholesaleGoogleSheetId = 'PUT_READ_ID_HERE';
  var headerKey = 'asin'
  var writeHeaders = ['location']; //These headers will have column written. Can add more to this array like 'product name'.

  var writeSheet = SpreadsheetApp.getActiveSheet();
  var writeSheetInfo = new SheetInfo(writeSheet);
  var writeKeyRowColCoordinate = getRowColCoordinateOfStr(writeSheetInfo, headerKey);
  var dictOfWriteHeadersRowColCoordinate = getDictOfCoordinates(writeSheetInfo, writeHeaders);
  
  var errorMsgsContainer = new ErrorMsgsContainer();
  if( allHeadersAreFound(writeSheetInfo, dictOfWriteHeadersRowColCoordinate, errorMsgsContainer) ){
    
    var wholesaleSpreadSheet = SpreadsheetApp.openById(wholesaleGoogleSheetId);
    var wholesaleDataDict = getDataDict(wholesaleSpreadSheet, headerKey, writeHeaders, errorMsgsContainer);

    writeLocation(writeSheet, writeSheetInfo, wholesaleDataDict, writeKeyRowColCoordinate, dictOfWriteHeadersRowColCoordinate);
    var successMsg = "Write successful.";
    if(errorMsgsContainer.errorMgs.length > 0){
      successMsg += "\\n\\nWarnings:\\n"
      successMsg += errorMsgsContainer.getErrorMsgs();
    }
    displayMsg(successMsg); 
  }
  else
    errorMsgsContainer.displayErrorMsgs("Error: Undetected headers in sheet where write is to be made. No edits were made.\\n");
}
