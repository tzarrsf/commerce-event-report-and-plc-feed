const debugMode = true;

let sheetName = "";
let sheetWithFocus = undefined;
const initialXY = {x: -1, y: -1};

// Column headers we care about for zoom data are Email, Name, Attended
const emailHeader = "Email";
let emailColumnIndex = initialXY
const nameHeader = "User Name (Original Name)";
let nameColumnIndex =  initialXY;
const attendedHeader = "Attended";
let attendedColumnIndex = initialXY;

function onOpen()
{  
	// Load the menu options
  addCustomMenu();
}

/*
* Add some menu options invoking functions to our tools
*/
function addCustomMenu()
{
  const menuName = "Commerce Team Tools";
  const ui = SpreadsheetApp.getUi();   
  let topMenu = ui.createMenu(menuName);
  topMenu.addSeparator();
  topMenu.addItem('Generate \'Registered\' report for PLC', 'menuGenerateRegisteredCsv');
  topMenu.addItem('Generate \'Attended\' report for PLC', 'menuGenerateAttendedCsv');
  topMenu.addSeparator();
  topMenu.addToUi();
  Browser.msgBox("Custom menu \"" + menuName + "\" ready for action! Click OK to continue.");
}

/*
* Convert the array of arrays (rows with columns) structure to CSV and push to download
*/
function saveAsCsvFile(data, description)
{
  const csvExtension = ".csv";
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  // Create a folder from the name of the spreadsheet including the description
  let folder = DriveApp.createFolder(ss.getName().toLowerCase().replace(/ /g,'_') + '_csv_' + new Date().getTime());
  // Append ".csv" extension to the sheet name
  let outputFileName = sheetWithFocus.getName() + " - " + description + csvExtension;
  // Assemble payload in CSV format
  let csvPayload = convertDataToCsvFile(data);
  // Create file in the Docs List with the given name containing the csv payload
  let file = folder.createFile(outputFileName, csvPayload);
  // Present the user with a download link
  let downloadURL = file.getDownloadUrl().slice(0, -8);
  showDownloadLink(downloadURL, description);
}

/*
 * Show a download link to the user to retrieve the CSV file
 */
function showDownloadLink(downloadURL, description)
{
    let link = HtmlService.createHtmlOutput('<a href="' + downloadURL + '">Click here to download</a>');
    SpreadsheetApp.getUi().showModalDialog(link, "Your " + description + " csv file is ready!");
}

/*
* Converts an array (rows) of arrays (columns) into csv data
*/
function convertDataToCsvFile(data)
{
  const crlf = "\r\n";
  const comma = ",";
  const quote = "\"";
  let csv = "";

  for (let row = 0; row < data.length; row++) {
    for (let col = 0; col < data[row].length; col++) {
      // Put quotes around data that already has a comma in the string
      if (data[row][col].toString().indexOf(comma) != -1) {
        data[row][col] = quote + data[row][col] + quote;
      }
    }

    // Join each row's columns adding a CRLF to the end of each row except the last one
    if(row < data.length-1)
    {
      csv += data[row].join(comma) + crlf;
    }
    else
    {
      csv += data[row];
    }
  }
  
  return csv;
}

/*
* Generate the "Registered" CSV for an event which is everyone in the sheet
*/
function menuGenerateRegisteredCsv()
{
  let registered = extractRegisteredFromSheet();
  saveAsCsvFile(registered, "Registered");
}

/*
* Generate the "Attended" CSV for an event - everyone in the sheet with Attended="Yes"
*/
function menuGenerateAttendedCsv()
{
  let attended = extractAttendedFromSheet();
  saveAsCsvFile(attended, "Attended");
}

/*
* Locate the x, y coordinates of our headers of interest and load the data from those points into
* an array of arrays (rows with columns).
*/
function getRegistrationDataAsTable()
{
  // Get the name of the active sheet and assign it for use throughout
  sheetName = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getName();
  
  // Get an x,y coordinate for each column we care about so we know where to scan
  emailColumnIndex = getHeaderXandYIndex(sheetName, emailHeader);
  nameColumnIndex = getHeaderXandYIndex(sheetName, nameHeader);
  attendedColumnIndex = getHeaderXandYIndex(sheetName, attendedHeader);

  /* Normalize the y coordinates across all the columns to be on the safe side favoring the highest value
  *  (to avoid issues with summary reports at the top of the sheet)
  */
  if(emailColumnIndex.y > nameColumnIndex.y)
  {
    nameColumnIndex.y = emailColumnIndex.y; 
  }
  if(nameColumnIndex.y > attendedColumnIndex.y)
  {
    attendedColumnIndex.y =nameColumnIndex.y; 
  }
  if(attendedColumnIndex.y > emailColumnIndex.y)
  {
    emailColumnIndex.y = attendedColumnIndex.y;
  }
  
  // Now build up the data for those columns we care about
  let dataRange = sheetWithFocus.getRange(1, 1, sheetWithFocus.getLastRow(), sheetWithFocus.getLastColumn());
  let data = dataRange.getValues();
  return data;
}

/*
 * Get all data that contains an email address and push it on to a clean array
 */
function extractRegisteredFromSheet()
{
  const emptyString = "";
  let data = getRegistrationDataAsTable();
  let result = [];
  
  for(let y = emailColumnIndex.y; y < data.length; y++)
  {
    if(data[y][emailColumnIndex.x] != null && data[y][emailColumnIndex.x].toString() !== emptyString){
      result.push([data[y][emailColumnIndex.x], data[y][nameColumnIndex.x], data[y][attendedColumnIndex.x]]);
    }
  }

  Logger.log(result.length + " registered extracted from sheet with focus.");
  return result;
}

/*
 * Get all data that contains an email address and Attended="Yes" and push it on to a clean array
 */
function extractAttendedFromSheet()
{
  const emptyString = "";
  const yes = "Yes";
  const attended = "Attended";
  let data = getRegistrationDataAsTable();
  let result = [];
  
  for(let y = emailColumnIndex.y; y < data.length; y++)
  {
    let attendedHeaderOrYes = (data[y][attendedColumnIndex.x] === yes || data[y][attendedColumnIndex.x] === attended);
    if(data[y][emailColumnIndex.x] != null && data[y][emailColumnIndex.x].toString() != emptyString && attendedHeaderOrYes){
      result.push([data[y][emailColumnIndex.x], data[y][nameColumnIndex.x], data[y][attendedColumnIndex.x]]);
    }
  }

  Logger.log(result.length + " attended extracted from sheet with focus.");
  return result;
}

/*
 * Find the x, y coordinates for headers using a nested reverse seek loop to avoid conflicts with text which is not a "real" header
 */
function getHeaderXandYIndex(sheetName, headerText)
{
  // Set an initial x, y coordinate result
  let result = {x: -1, y: -1};
  
  // Get sheet we have focus on
  sheetWithFocus = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  
  if(sheetWithFocus == null)
  {
    Logger.log('Sheet not located in getHeaderXandYIndex using: ' + sheetName);
    return result;
  }

  let dataRange = sheetWithFocus.getRange(1, 1, sheetWithFocus.getLastRow(), sheetWithFocus.getLastColumn());
  let data = dataRange.getValues();

  // Brute force seek the columns in reverse to keep from running into other occurences of "Attended" or "Email" in the summary reports top of sheet
  let found = false;
  for (result.y = data.length - 1; result.y > -1; result.y--) {
    for (result.x = data[result.y].length - 1; result.x > -1; result.x--) {
      if(data[result.y][result.x].toString() == headerText){
        found = true;
        break;
      }
    }
    if(found)
    {
      break;
    }
  }
  
  return result;
}
