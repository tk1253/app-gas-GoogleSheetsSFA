function executeActivityReports() {
  var mainSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var mainSheet = mainSpreadsheet.getSheetByName('活動一覧'); 

  // B～H列の4行目以下をクリア
  mainSheet.getRange('B4:H').clearContent();

  var fromDate = new Date(mainSheet.getRange("C1").getValue());
  var toDate = new Date(mainSheet.getRange("C2").getValue());

  Logger.log("From Date: " + fromDate);
  Logger.log("To Date: " + toDate);

  var referenceSheet = mainSpreadsheet.getSheetByName('参照範囲');
  if (!referenceSheet) {
    Logger.log("Error: '参照範囲'シートが見つかりません。");
    return;
  }

  var urlRange = referenceSheet.getRange("D2:D").getValues().flat().filter(String);
  var rangeStart = referenceSheet.getRange("E2:E").getValues().flat().filter(String);
  var rangeEnd = referenceSheet.getRange("F2:F").getValues().flat().filter(String);

  Logger.log("URLs: " + urlRange);
  Logger.log("Range Starts: " + rangeStart);
  Logger.log("Range Ends: " + rangeEnd);

  var spreadsheetIds = urlRange.map(function(url) {
    return url.split('/')[5]; // スプレッドシートIDをURLから抽出
  });

  Logger.log("Spreadsheet IDs: " + spreadsheetIds);

  var targetRow = 4; // メインシートのB4から転記を開始

  spreadsheetIds.forEach(function(spreadsheetId, index) {
    try {
      var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      var sheets = spreadsheet.getSheets();
      var spreadsheetUrl = urlRange[index];
      var startCell = rangeStart[index];
      var endCell = rangeEnd[index];

      Logger.log("Processing Spreadsheet ID: " + spreadsheetId);
      Logger.log("Using range: " + startCell + ":" + endCell);

      sheets.forEach(function(sheet) {
        if (sheet.getName() != '一覧' && sheet.getName() != 'コピー元') {
          Logger.log("Processing Sheet: " + sheet.getName());
          
          // 16行目からヘッダーを取得
          var headers = sheet.getRange("C16:Z16").getValues()[0]; 

          // 各列のインデックスを探す
          var dateColIndex = headers.indexOf("日付");
          var visitColIndex = headers.indexOf("訪問先");
          var visitorColIndex = headers.indexOf("訪問者");
          var usageColIndex = headers.indexOf("用途");
          var contactMemoColIndex = headers.indexOf("コンタクトメモ");

          if (dateColIndex === -1) dateColIndex = 0; // 日付が見つからなければC列とみなす

          Logger.log("Column Indices - Date: " + dateColIndex + ", Visit: " + visitColIndex + ", Visitor: " + visitorColIndex + ", Usage: " + usageColIndex + ", Contact Memo: " + contactMemoColIndex);

          var range = sheet.getRange(startCell + ":" + endCell);
          var values = range.getValues();
          var sheetUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit#gid=${sheet.getSheetId()}&range=A1`;

          values.forEach(function(row, rowIndex) {
            var cellDate = new Date(row[dateColIndex]);
            Logger.log("Row " + (rowIndex + 1) + ": Cell Date - " + cellDate);

            if (isNaN(cellDate.getTime())) {
              Logger.log("Row " + (rowIndex + 1) + ": Invalid date format");
              return; // Skip invalid dates
            }

            if (cellDate >= fromDate && cellDate <= toDate) {
              Logger.log("Date Match Found: " + cellDate);
              mainSheet.getRange("B" + targetRow).setValue(cellDate); // 日付をB列に
              mainSheet.getRange("C" + targetRow).setValue(sheet.getName()); // シート名をC列に
              if (visitColIndex >= 0) mainSheet.getRange("D" + targetRow).setValue(row[visitColIndex]); // 訪問先をD列に
              if (visitorColIndex >= 0) mainSheet.getRange("E" + targetRow).setValue(row[visitorColIndex]); // 訪問者をE列に
              if (usageColIndex >= 0) mainSheet.getRange("F" + targetRow).setValue(row[usageColIndex]); // 用途をF列に
              if (contactMemoColIndex >= 0) mainSheet.getRange("G" + targetRow).setValue(row[contactMemoColIndex]); // コンタクトメモをG列に
              mainSheet.getRange("H" + targetRow).setValue(sheetUrl); // シートのA1セルへのリンクをH列に
              Logger.log("Copied Data to row " + targetRow);
              targetRow++;
            }
          });
        }
      });
    } catch (e) {
      Logger.log("Error processing spreadsheet ID " + spreadsheetId + ": " + e.message);
    }
  });
}
