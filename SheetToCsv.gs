
function onOpen() {
    SpreadsheetApp.getUi()
                  .createMenu("Export CSV")
                  .addItem("This Sheet to CSV","showDialog")
                  .addToUi();

};

function showDialog()
{
  var html = HtmlService.createHtmlOutputFromFile("Download");
  SpreadsheetApp.getUi().showModalDialog(html,"CSV Download");
} 

function SaveAndGetFileUrl() 
{
  var ss = SpreadsheetApp.getActiveSpreadsheet(); 
  var sheet = ss.getActiveSheet();
  // create a folder from the name of the spreadsheet
  var folder = DriveApp.createFolder(ss.getName().toLowerCase().replace(/ /g,'_') + '_csv_' + new Date().getTime());
  // append ".csv" extension to the sheet name
  fileName = sheet.getName() + ".csv";
  // convert all available sheet data to csv format
  var csvData = convertRangeToCsv(sheet);
  // create a file in the Docs List with the given name and the csv data
  var file = folder.createFile(fileName, csvData ,MimeType.CSV);
  //File downlaod
  var downloadURL = file.getDownloadUrl().replace("?e=download&gd=true","");

  return downloadURL;
} 

function convertRangeToCsv(sheet) 
{
  // get available data range in the spreadsheet
  var activeRange = sheet.getDataRange();
  try {
    var data = activeRange.getValues();
    var csvFile = undefined;

    // loop through the data in the range and build a string with the csv data
    if (data.length > 1) {
      var csv = "";
      for (var row = 0; row < data.length; row++) {
        for (var col = 0; col < data[row].length; col++) {
          if (data[row][col].toString().indexOf(",") != -1) {
            data[row][col] = "\"" + data[row][col] + "\"";
          }
        }

        // join each row's columns
        // add a carriage return to end of each row, except for the last one
        if (row < data.length-1) {
          csv += data[row].join(",") + "\r\n";
        }
        else {
          csv += data[row];
        }
      }
      csvFile = csv;
    }
    return csvFile;
  }
  catch(err) {
    Logger.log(err);
    Browser.msgBox(err);
  }
}
