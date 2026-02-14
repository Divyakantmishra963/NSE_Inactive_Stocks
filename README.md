Google Sheets Apps Script to easillly get  a few NSE Listed Inactive Stocks is ready for Coppy Paste in the Extension than Deleting the Exsisting Code and replacing by the following Script: function runMasterAutomation() {
  if (!isTradingDay()) {
    Logger.log("आज मार्केट बंद है। स्कैनिंग नहीं होगी।");
    return;
  }
  cleanupOldData();
  runOptimizedBatchScanner();
}

function isTradingDay() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheetByName("Settings");
  const today = new Date();
  if (today.getDay() === 0 || today.getDay() === 6) return false;
  
  try {
    const holidayRange = settingsSheet.getRange("A2:A100").getValues();
    const todayStr = Utilities.formatDate(today, "IST", "yyyy-MM-dd");
    for (let i = 0; i < holidayRange.length; i++) {
      let dateVal = holidayRange[i][0];
      if (dateVal instanceof Date && Utilities.formatDate(dateVal, "IST", "yyyy-MM-dd") === todayStr) return false;
    }
  } catch(e) { return true; }
  return true;
}

function runOptimizedBatchScanner() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Portfolio"); 
  const targetSheet = ss.getSheetByName("Inactive_Stocks") || ss.insertSheet("Inactive_Stocks");
  
  const lastRow = sourceSheet.getLastRow();
  if (lastRow < 2) return;
  
  const data = sourceSheet.getRange(2, 1, lastRow - 1, 1).getValues();
  const results = [];
  const scanDate = new Date();

  for (let i = 0; i < data.length; i++) {
    let symbol = data[i][0].toString().trim();
    if (!symbol) continue;

    let ticker = "NSE:" + symbol;
    let tempCell = targetSheet.getRange("Z1");

    try {
      tempCell.setFormula(`=IFERROR(GOOGLEFINANCE("${ticker}", "marketcap")/10000000, 0)`);
      SpreadsheetApp.flush();
      Utilities.sleep(1000);
      let mCap = tempCell.getValue();

      tempCell.setFormula(`=IFERROR(AVERAGE(QUERY(GOOGLEFINANCE("${ticker}", "volume", TODAY()-30, TODAY()), "select Col2 offset 1", 0)), 0)`);
      SpreadsheetApp.flush();
      Utilities.sleep(2000);
      let avgVol = tempCell.getValue();

      Logger.log("Ticker: " + ticker + " | MCap: " + mCap + " | Vol: " + avgVol);

      if (mCap >= 0 && mCap < 500 && avgVol < 500000) {
        results.push([symbol, mCap.toFixed(2), avgVol.toFixed(2), scanDate]);
      }
    } catch (err) {
      Logger.log("Error processing " + ticker + ": " + err.message);
    }
    
    if (i % 5 === 0) Utilities.sleep(2000);
  }

  if (results.length > 0) {
    targetSheet.getRange(targetSheet.getLastRow() + 1, 1, results.length, 4).setValues(results);
  }
  targetSheet.getRange("Z1").clear();
  ss.getSheetByName("Settings").getRange("C2").setValue("Last Scan: " + Utilities.formatDate(new Date(), "IST", "dd-MMM-yyyy HH:mm 'IST'"));
}

function cleanupOldData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName("Inactive_Stocks");
  if (!targetSheet || targetSheet.getLastRow() < 2) return;
  const lastRow = targetSheet.getLastRow();
  const data = targetSheet.getRange(2, 4, lastRow - 1, 1).getValues();
  const today = new Date();
  for (let i = data.length - 1; i >= 0; i--) {
    let rowDate = new Date(data[i][0]);
    if (!isNaN(rowDate.getTime()) && (today - rowDate > 604800000)) targetSheet.deleteRow(i + 2);
  }
}

function runLot(lotNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName("Portfolio");
  const targetSheet = ss.getSheetByName("Inactive_Stocks") || ss.insertSheet("Inactive_Stocks");

  const batchSize = 100;
  const startRow = (lotNumber - 1) * batchSize + 2;
  const endRow = Math.min(startRow + batchSize - 1, sourceSheet.getLastRow());

  const data = sourceSheet.getRange(startRow, 1, endRow - startRow + 1, 1).getValues();
  const results = [];
  const scanDate = new Date();

  for (let i = 0; i < data.length; i++) {
    let symbol = data[i][0].toString().trim();
    if (!symbol) continue;
    let ticker = "NSE:" + symbol;
    let tempCell = targetSheet.getRange("Z1");
    try {
      tempCell.setFormula(`=IFERROR(GOOGLEFINANCE("${ticker}", "marketcap")/10000000, 0)`);
      SpreadsheetApp.flush();
      Utilities.sleep(500);
      let mCap = tempCell.getValue();

      tempCell.setFormula(`=IFERROR(AVERAGE(QUERY(GOOGLEFINANCE("${ticker}", "volume", TODAY()-30, TODAY()), "select Col2 offset 1", 0)), 0)`);
      SpreadsheetApp.flush();
      Utilities.sleep(1000);
      let avgVol = tempCell.getValue();

      if (mCap >= 0 && mCap < 500 && avgVol < 500000) {
        results.push([symbol, mCap.toFixed(2), avgVol.toFixed(2), scanDate]);
      }
    } catch (err) {
      Logger.log("Error processing " + ticker + ": " + err.message);
    }
  }

  if (results.length > 0) {
    targetSheet.getRange(targetSheet.getLastRow() + 1, 1, results.length, 4).setValues(results);
  }
  targetSheet.getRange("Z1").clear();

  Logger.log("Lot " + lotNumber + " processed from row " + startRow + " to " + endRow);
}

function runLot1() { runLot(1); }
function runLot2() { runLot(2); }
function runLot3() { runLot(3); }
function runLot4() { runLot(4); }
function runLot5() { runLot(5); }

function resetInactiveStocks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const targetSheet = ss.getSheetByName("Inactive_Stocks");
  if (targetSheet) {
    if (targetSheet.getLastRow() > 1) {
      targetSheet.deleteRows(2, targetSheet.getLastRow() - 1);
    }
    Logger.log("Inactive_Stocks reset completed at " + new Date());
  }
  const settingsSheet = ss.getSheetByName("Settings");
  if (settingsSheet) {
    settingsSheet.getRange("C2").setValue("Last reset: " + Utilities.formatDate(new Date(), "IST", "dd-MMM-yyyy HH:mm 'IST'"));
  }
} 2- Now Rename your Project as You like and Save. 3- Set the Triggeers as You wish better after the Market. Allow the Permission to Google. 3- This is designed for the 500 Tickers which can be ffchnaged to smaller or the larger size. 4- Make sub Sheets name them Portfolio,Inactive_Stocks and Settings.Just Paste the NSE Stock's Ticker or Symbols in Cell A2 and Downwards with a Header in Cell A1 SYMBOL.In the next Sheet Inactive_Stocks make Headers manually once as shown : Cell A1 TICKER	Cell B1 MARKET CAP	Cell C1 AVG VOLUME (1M)	Cell D1 Scan Date (IST)	.in Cell E1 use the Formula =IF(COUNTA(D2:D)=0, "No Scan Data", "✅ Status:") and in Cell F1 Apply the Formula =IF(MAX(D2:D)=0, "Waiting for Scan...", MAX(D2:D)).The Cell A2 will  show the Inactive Stock if any and fill up the Cells A2,B2,C2,D2,E2.Well the E and F provided the Actual Status of the Script functioning.5- Settings sheets needs once in a year manual copy paste the NSE Holidays list in December or In January in Cells A3 and B3 and downwards.Headers in Cell A2 and B2 are respectively Date & Holiday Name.Thats all .It is a Time Saver which protects Effective Time from many angles. Thank You. You can further add improvements but for me it is enough to have a glimpse.Best Regards.The Link is as given Below: https://docs.google.com/spreadsheets/d/1P110bH4j7-1Oqv4ND1agqevpcNrwz8noQF5QVJX0tbA/edit?usp=sharing.
