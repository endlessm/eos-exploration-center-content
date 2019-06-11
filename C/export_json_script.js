//Use this script with the Google Spreadsheet to export the Combined sheet into json

var base = "DC_Tiles_";
var colors = ["Blue", "Chartreuse", "Fuschia", "Green", "HotPink", "Orange", "Purple", "Slate"];
var max_num = 8;

function randomTile() {
  return base+colors[getRandom(0, max_num)]+"_"+getRandom(0, max_num)+".png";
}

function getRandom(min,max){
  return Math.floor(Math.random() * (max - min)) +min;
}

///////////////////////////
/**
 * Adds a custom menu to the active spreadsheet, containing a single menu item
 * for invoking the exportJSON() function specified above.
 * The onOpen() function, when defined, is automatically invoked whenever the
 * spreadsheet is opened.
 * For more information on using the Spreadsheet API, see
 * https://developers.google.com/apps-script/service_spreadsheet
 */
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var entries = [{
    name : "Do It!",
    functionName : "exportJSON"
  }];
  sheet.addMenu("Export JSON", entries);
}

// triggers parsing and displays results in a text area inside a custom modal window
function exportJSON() {
  //var app = UiApp.createApplication().setTitle('JSON export results - select all and copy!');
  //var textArea = app.createTextArea();
  //textArea.setValue(makeJson(SpreadsheetApp.getActiveSheet().getDataRange()));
  //app.add(textArea);
  //textArea.setSize("100%", "100%");
  //SpreadsheetApp.getActiveSpreadsheet().show(app);
  createGoogleDriveTextFile();
}


function createGoogleDriveTextFile() {
  var content,fileName,newFile;//Declare variable names

  fileName = "Hack_Chrome_data_ " + new Date().toString().slice(0,15)+".json";//Create a new file name with date on end
  content = makeHackJson(SpreadsheetApp.getActiveSheet().getDataRange());
  var fileContent =JSON.stringify(content, null, "\t"); // Indented with tab

  newFile = DriveApp.createFile(fileName,fileContent);//Create a new text file in the root folder
};

function makeHackJson(dataRange){
  var frozenRows = SpreadsheetApp.getActiveSheet().getFrozenRows();
  var dataRangeArray = dataRange.getValues();
  var dataWidth = dataRange.getWidth();
  var dataHeight = dataRange.getHeight() - frozenRows;

  // range of names - we assume that the last frozen row is the list of properties
  var nameRangeArray = dataRangeArray[frozenRows - 1];

  var Result = [];

  for (var h = 0; h < dataHeight ; ++h) {

    var entry = {};

    for (var i = 0; i < dataWidth; ++i) {

      var param = nameRangeArray[i];
      var data = dataRangeArray[h + frozenRows][i];

      if (param.length >0) {
        if (param == "image"){
          //add image base url
          var base = "https://exploration-center.coding.endlessdeployment.com/img/";
          entry["image"] =  jsonEscape(base+data);
        } else {
          // add data
          entry[param] =  jsonEscape(data);
        }

      }
    }
    Result[h] = entry;
  }
  return Result;
}

function makeJson(dataRange) {
  var charSep = '"';

  var result = "", thisName = "", thisData = "";

  var frozenRows = SpreadsheetApp.getActiveSheet().getFrozenRows();
  var dataRangeArray = dataRange.getValues();
  var dataWidth = dataRange.getWidth();
  var dataHeight = dataRange.getHeight() - frozenRows;

  // range of names - we assume that the last frozen row is the list of properties
  var nameRangeArray = dataRangeArray[frozenRows - 1];

  // open JSON object - if there's a extra frozen row on the top wrap results into that as property (only supports one for now)
  result += frozenRows > 1 ? '{"' + dataRangeArray[frozenRows - 2][0] + '": [' : '[';

  for (var h = 0; h < dataHeight ; ++h) {

    result += '{';

    for (var i = 0; i < dataWidth; ++i) {

      thisName = nameRangeArray[i];
      thisData = dataRangeArray[h + frozenRows][i];

      // add name
      result += charSep + thisName + charSep + ':'

      if (thisName == "image"){
        //add image base url
        var base = "https://exploration-center.coding.endlessdeployment.com/img/";
        result += charSep + jsonEscape(base+thisData) + charSep + ', ';

      } else {
        // add data
        result += charSep + jsonEscape(thisData) + charSep + ', ';
      }

    }

    //remove last comma and space
    result = result.slice(0,-2);

    result += '},\n';

  }

  //remove last comma and line break
  result = result.slice(0,-2);

  // close object
  result += frozenRows > 1 ? ']}' : ']';

  return result;

}

function jsonEscape(str)  {
  if (typeof str === "string" && str !== "") {
    return str.replace(/\n/g, "<br/>").replace(/\r/g, "<br/>").replace(/\t/g, "\\t");
  } else {
    return str;
  }
}
