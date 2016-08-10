//Global variables
var ss = SpreadsheetApp.getActiveSpreadsheet(),
    sheet = ss.getActiveSheet(),
    cells = ss.getActiveRange();

function areaToCoordinates() {
    
  //Initialize data column and row
  var areaColumn = 1;
  var areaRow;
  
  //Insert new sheets
  var longSheet = sheet.insertColumnsAfter(areaColumn, 1);
  var latSheet = sheet.insertColumnsAfter(areaColumn, 2);
  
  //Set long and lat columns
  var longColumn = areaColumn + 1;
  var latColumn = areaColumn + 2;
  
  //Set column headers for new columns
  var longHeader = longSheet.getRange(1, areaColumn + 1).setValue('longitude');
  var latHeader = latSheet.getRange(1, areaColumn + 2).setValue('latitude');

  //Initialize geocoder
  var geocoder = Maps.newGeocoder();
  
  for (areaRow = 2; areaRow <= cells.getNumRows(); ++areaRow) { 
      //Retrieve value of area in each row excluding header and pass it to the geocoder object
      var area = cells.getCell(areaRow, areaColumn).getValue();
      var location = geocoder.geocode(area);
      
      //Set the long and lat values to its cell if location is populated
      if (location.status == 'OK') {
        
        //Parse hashtable of list of hashtables to retrieve approximate location
        var long = location['results'][0]['geometry']['location']['lng'];
        var lat = location['results'][0]['geometry']['location']['lat'];
      
        sheet.getRange(areaRow, longColumn).setValue(long);
        sheet.getRange(areaRow, latColumn).setValue(lat);
      } else if (location.status == 'ZERO_RESULTS') { 
        continue;
      }
  }
};