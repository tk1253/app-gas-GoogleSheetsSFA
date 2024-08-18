function createNewSheetFromCheckedRows() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName('一覧'); // 元のデータがあるシート名を設定
  var templateSheet = ss.getSheetByName('コピー元'); // テンプレートシートの名前

  if (!sourceSheet) {
    Logger.log('元のシート「一覧」が見つかりません。');
    return;
  }
  if (!templateSheet) {
    Logger.log('テンプレートシート「コピー元」が見つかりません。');
    return;
  }

  var data = sourceSheet.getDataRange().getValues();
  
  var columnIndexI = 8; // I列
  var columnIndexC = 2; // C列
  var columnIndexA = 0; // A列
  var columnIndexF = 5; // F列
  var columnIndexG = 6; // G列
  var columnIndexH = 7; // H列

  // データを走査して、I列にチェックが入っている行を新しいシートに追加
  for (var i = 1; i < data.length; i++) {
    if (data[i][columnIndexI] === true) { // I列にチェックが入っている場合
      var sheetName = data[i][columnIndexC]; // C列の値をシート名にする
      var valueToCopyA = data[i][columnIndexA]; // A列の値を取得
      var valueToCopyF = data[i][columnIndexF]; // F列の値を取得
      var valueToCopyG = data[i][columnIndexG]; // G列の値を取得
      var valueToCopyH = data[i][columnIndexH]; // H列の値を取得
      
      Logger.log('シートを作成中: ' + sheetName);
      createSheetFromTemplate(ss, templateSheet, sheetName, valueToCopyA, valueToCopyF, valueToCopyG, valueToCopyH);
      
      // ハイパーリンクを設定
      var sheetLink = getSheetLink(ss, sheetName); // 新しいシートへのリンクを取得
      sourceSheet.getRange(i + 1, columnIndexC + 1).setFormula('=HYPERLINK("' + sheetLink + '", "' + sheetName + '")');
      
      // チェックボックスを外す
      sourceSheet.getRange(i + 1, columnIndexI + 1).setValue(false); // getRangeは1ベースのインデックス
    }
  }
}

function createSheetFromTemplate(ss, templateSheet, sheetName, valueToCopyA, valueToCopyF, valueToCopyG, valueToCopyH) {
  var newSheet = ss.getSheetByName(sheetName);
  if (newSheet) {
    ss.deleteSheet(newSheet); // 既存のシートがある場合は削除
  }
  newSheet = templateSheet.copyTo(ss).setName(sheetName); // テンプレートシートを複製して新しいシートを作成

  Logger.log('新しいシート「' + sheetName + '」を作成しました。');
  newSheet.getRange('D6').setValue(valueToCopyA); // A列の値をD6に設定
  newSheet.getRange('D10').setValue(valueToCopyF); // F列の値をD10に設定
  newSheet.getRange('C17').setValue(valueToCopyG); // G列の日付をC17に設定
  newSheet.getRange('D17').setValue(valueToCopyH); // H列の値をD17に設定
}

function getSheetLink(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (sheet) {
    var sheetId = sheet.getSheetId();
    var spreadsheetUrl = ss.getUrl();
    return spreadsheetUrl + '#gid=' + sheetId; // シートへのURLを生成
  } else {
    Logger.log('シート「' + sheetName + '」が見つかりません。');
    return '';
  }
}
