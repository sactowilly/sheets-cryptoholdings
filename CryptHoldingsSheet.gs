/*====================================================================================================================================*
  CryptoHoldings by Will Zrnchik  
  ====================================================================================================================================
  Version:      0.0.1
  Project Page: https://github.com/sactowilly/sheets-cryptoholdings
  Copyright:    (c) 2019 by Will Zrnchik
  License:      GNU General Public License, version 3 (GPL-3.0) 
                http://www.opensource.org/licenses/gpl-3.0.html
  ------------------------------------------------------------------------------------------------------------------------------------
  A script to help calculate personal crypto holdings using Google Sheets and Scripts. Functions include:
     onOpen .              Invoked with the spreadsheet opens.
     getPrefs              Initial loading of user references.    
     updatePrefs           Updates changes to user preferences. 
     callWorldCoinIndex    Gets the current prices of WCI's top 2000 coins    
     GETHISTORICALPRICE    Simply put, it gets historical data from a cell    
  
  For future enhancements see https://github.com/sactowilly/sheets-cryptholdings/issues?q=is%3Aissue+is%3Aopen+label%3Aenhancement
  
  For bug reports you could go to https://github.com/sactowilly/sheets-cryptholdings/issues, but you might be better off Googling shit.
  
  Shit...at least the comments make it look like I am a code boss or something. Wait... no?
  ------------------------------------------------------------------------------------------------------------------------------------
  Changelog:
  
  1.0.0  Initial release (2019-03-10)
  *====================================================================================================================================*/
 

/*====================================================================================================================================*
          Global Variables
 =====================================================================================================================================*
 *
 * Set some globals variables so we don't have to keep doing local ones
 *
 * sheetPref         the Preferences sheet
 * sheetData .       the sheet where coin data will be displayed
 * WCI_URL .         the URL to World Coin Index's API for v2getmarkets
 * key               YOUR API key from World Coin Index. Get yours at https://www.worldcoinindex.com/apiservice.
 * fiat              the three-character code for the fiat currency you want the results returned against.
 *
 **/
 
  var app = SpreadsheetApp;
  var ss = app.getActiveSpreadsheet();
  
  // Get preference values
  var sheetPref = ss.getSheetByName("zPreferences");
  var key = sheetPref.getRange(16, 9).getValue();
  var fiat = sheetPref.getRange(17, 9).getValue();
  var WCI_URL = sheetPref.getRange(18, 9).getValue();
  
  var sheetData = ss.getSheetByName("zData");


/*====================================================================================================================================*
          onOpen
 =====================================================================================================================================*
 *
 * Creates a menu item when opening the CryptoHoldings Worksheet
 * 
 **/

function onOpen() {
  // Add custom menu items under one main menu item  
  var ui = app.getUi();
  ui.createMenu('Crypto')
      .addItem('Update WCI Values', 'callWorldCoinIndex')
      .addItem('Update Preferences', 'updatePrefs')
      .addToUi();
}


/*====================================================================================================================================*
          getPrefs
 =====================================================================================================================================*
 *
 * Gets user preferences from the Prferences sheet
 *
 * sheetPref         the Preferences sheet
 * WCI_URL .         the URL to World Coin Index's API for v2getmarkets
 * key               YOUR API key from World Coin Index. Get yours at https://www.worldcoinindex.com/apiservice.
 * fiat              the three-character code for the fiat currency you want the results returned against.
 *
 **/
 
 function getPrefs() {
  var key = sheetPref.getRange(16, 9).getValue();
  var fiat = sheetPref.getRange(17, 9).getValue();
  var WCI_URL = sheetPref.getRange(18, 9).getValue();
  // Logger.log(key);
  // Logger.log(fiat);
}


/*====================================================================================================================================*
          updatePrefs
 =====================================================================================================================================*
 *
 * Updates user preferences from the Prferences sheet and executes a Toast script
 *
 * sheetPref         the Preferences sheet
 * WCI_URL .         the URL to World Coin Index's API for v2getmarkets
 * key               YOUR API key from World Coin Index. Get yours at https://www.worldcoinindex.com/apiservice.
 * fiat              the three-character code for the fiat currency you want the results returned against.
 *
 **/

function updatePrefs() {
  var key = sheetPref.getRange(16, 9).getValue();
  var fiat = sheetPref.getRange(17, 9).getValue();
  var WCI_URL = sheetPref.getRange(18, 9).getValue();  
  ss.toast(
    "The following preferences have been updated: API Key " + key + 
    ", fiat currency to" + fiat + 
    ", and the WCI URL is " + WCI_URL, 
    "Preferences Updated", 
    5)
}


/*====================================================================================================================================*
          callWorldCoinIndex
 =====================================================================================================================================*
 *
 * Using your own World Coin Index API Imports a JSON feed of crytop currency values and returns the results to be inserted into a Google Spreadsheet. data.youtube.com/feeds/api/standardfeeds/most_popular?v=2&alt=json", "/feed/entry/title,/feed/entry/content", "noInherit,noTruncate,rawHeaders")
 * 
 **/
 
 /**
 * @OnlyCurrentDoc
 */


function callWorldCoinIndex() {

  // Build the URL to call the World Coin Index 
  var k = '?key='+ key;
  var f = '&fiat=' + fiat;
  var dataURL = WCI_URL + k + f + "&limit=2000";

  // Call the World Coin Index API for current rates
  var response = UrlFetchApp.fetch(dataURL);
  var jsondata = JSON.parse(response.getContentText());  
  var dataAll = jsondata
  var dataSet = dataAll.data;
  var dataLength = dataSet.length;
  Logger.log(dataLength);
  
  var rows = [],
    data;

  if (dataSet.length > 0) {
    ss.toast("Inserting " + dataLength + " rows.", "Getting Crypto Data");
    for (i = 0; i < dataSet.length; i++) {
      data = dataSet[i];
      rows.push([
        data.rank,
        data.symbol,
        data.name, 
        data.id, 
        data.priceUsd, 
        data.volumeUsd24Hr, 
        data.marketCapUsd, 
        data.vwap24Hr, 
        data.changePercent24Hr,  
        data.maxSupply, 
        data.supply
        ]);
        }
     } else {
    ss.toast("All done", "Getting Crypto Data", 5);
  }  

  if (data.length>0){
    ss.toast("Inserting "+data.length+" rows");
    sheet.insertRowsAfter(1, data.length);
    setRowsData(sheet, data);
  } else {
    ss.toast("All done");
  } 

  // Logger.log(rows);
  
  // [row to start on], [column to start on], [number of rows], [number of entities]
  dataRange = sheetData.getRange(2, 1, rows.length, 11);
  dataRange.setValues(rows);

}

/*====================================================================================================================================*
          GETHISTORICALPRICE
 =====================================================================================================================================*
 *
 * Using the Multiexplorer API to get historical value of a currency at the time it was used for purchasing another crypto 
 *
 * @idSymbol          pulls the character code for a currency
 * @timeUNIX          calculates the price for said currency at the time of the transacion
 * 
 **/
 
 /**
 * @OnlyCurrentDoc
 */
 
 function GETHISTORICALPRICE(idSymbol,timeUNIX) {
  var dataURL = "https://multiexplorer.com/api/historical_price?fiat=usd&currency=";
  dataURL += idSymbol; 
  dataURL += "&time=";
  dataURL += timeUNIX;

  // Call Multiexplorer API for historical rates and parse the response
  var response = UrlFetchApp.fetch(dataURL);
  var json = JSON.parse(response);
  Logger.log(json.price);
  Logger.log(dataURL);
  Logger.log(response);
  
  // Now pull the only thing that really matters and print it: the price at a certain time
  var coinPrice = json.price;
  return coinPrice;
  
}
