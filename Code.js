function onOpen() {
  SpreadsheetApp.getActiveSpreadsheet()
  .addMenu('Prima nota', [
  {
    name: "Export to PDF",
    functionName: "processExports"
  }])
}

function processExports() {
  var controlPanelSpreadsheet = SpreadsheetApp.getActive();
  var folder = parentFolder(controlPanelSpreadsheet);   
  
  cleanUp(folder);  
  forEachXls(folder, processExport(folder));
  Browser.msgBox('Export processed!');
}

function processExport(folder) {
  return function(file) {
    processFile(folder, file)
  }
}

function processFile(folder, xlsFile) {  
  var file = convertXls(folder, xlsFile)
  modifyFile(file)
  SpreadsheetApp.flush();
  exportRefund(file, folder.getFoldersByName("processati").next())
  moveFile(folder, folder.getFoldersByName("originali").next(), xlsFile)
}

function moveFile(folder, destFolder, file) {
  folder.removeFile(file);
  destFolder.addFile(file);
}

function modifyFile(file) {
  var spreadsheet = SpreadsheetApp.open(file);
  var sheet = spreadsheet.getSheets()[0];

  removeFirstTwoLines(sheet);
  removeBalanceColumn(sheet);
  
  var size = getDataSize(sheet);  
  resetColors(size, sheet);
  cleanupText(sheet);
  setTextWrap(sheet);
  setGrid(size, sheet);
  resizeAll(size, sheet);
}

function convertXls(folder, file) {
  // auto-convert xls in gsheet require Advanced Drive Service
  // https://developers.google.com/apps-script/reference/drive  
  var xlsFile = DriveApp.getFileById(file.getId()); // File instance of Excel file    
  var xlsBlob = xlsFile.getBlob(); // Blob source of Excel file for conversion
  var options = {
    title: file.getName().replace(".xls", ""),
    parents: [ { id: folder.getId() } ]
  };
  var result = Drive.Files.insert(options, xlsBlob, { convert: true });
  return DriveApp.getFileById(result.id);
}

function exportRefund(file, exportFolder) {
  var pdfFileName = file.getName() + '.pdf';
  
  if(fileAlreadyExists(exportFolder, pdfFileName)) {
    throw "File already exists: " + pdfFileName;
  }
  
  var url = 'https://docs.google.com/spreadsheets/d/' + file.getId() + '/export?exportFormat=pdf&format=pdf'
  + '&size=A4&portrait=false'
  + '&top_margin=0.50&bottom_margin=0.50&left_margin=0.50&right_margin=0.50' // All four margins must be set!
  + '&sheetnames=false&printtitle=false'
  + '&pagenumbers=false&gridlines=false';
  
  var response = UrlFetchApp.fetch(url, {
    headers: {
      'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken()
    }
  });

  var blob = response.getBlob().setName(pdfFileName);
  exportFolder.createFile(blob).setName(blob.getName());
}

// -- file modifiers

function removeFirstTwoLines(sheet) {
  sheet.deleteRows(1, 2);
}

function removeBalanceColumn(sheet) {
  sheet.deleteColumn(7)
}

function resetColors(size, sheet) {
  var range = sheet.getRange(size.first, 1, size.count, 6);
  range.setBackground(null)
  range.setFontColor(null)
}

function cleanupText(sheet) {
  replaceAll('Banca Unicredit', 'Banca', sheet);
  replaceAll('EUR 0.00', '', sheet);
}

function setTextWrap(sheet) {
  var cells = sheet.getRange("C:D");
  cells.setWrap(true);
}

function setGrid(size, sheet) {
  var range = sheet.getRange(size.first, 1, size.count, 6);
  range.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID)
}

function resizeAll(size, sheet) {
  sheet.autoResizeColumn(1);
  sheet.setColumnWidth(2, 65);
  sheet.setColumnWidth(3, 510);
  sheet.setColumnWidth(4, 180);
  sheet.autoResizeColumn(5);
  sheet.autoResizeColumn(6);
  sheet.setRowHeights(size.first, size.count, 90);  
}

// -- utility

function cleanUp(folder) {
  var files = folder.getFiles()
  while (files.hasNext()){
    var file = files.next();
    if (file.getName().startsWith("prima nota") 
        && !file.getName().endsWith(".xls")) {
      file.setTrashed(true);
    }
  }
}

function fileAlreadyExists(folder, fileName){
  return folder.getFilesByName(fileName).hasNext();
}

function replaceAll(from, to, sheet) {
  var textFinder = sheet.createTextFinder(from);
  textFinder.matchCase(true);
  textFinder.matchEntireCell(true)
  textFinder.replaceAllWith(to);
}

function getDataSize(sheet) {
  var column = sheet.getRange('A:A');
  var values = column.getValues();
  
  var ctdata = 0;
  while (values[ctdata][0] != "Data" ) {
    ctdata++;
  }
  
  var offset = ctdata + 1;
  var ct = offset;
  while (values[ct][0] != "" ) {
    ct++;
  }
  return { first: offset + 1, count: ct - offset};
}

function parentFolder(spreadsheet) {
  var spreadsheetId = spreadsheet.getId();
  var file = DriveApp.getFileById(spreadsheetId);  
  var folders = file.getParents();  
  while (folders.hasNext()){
    return folders.next();
  }
  throw "Unable to find parent folder"
}

function forEachXls(folder, func) {
  var files = folder.getFiles()
  while (files.hasNext()){
    var file = files.next();
    if (file.getName().startsWith("prima nota") 
        && file.getName().endsWith(".xls")) {
      func(file)
    }
  }  
}