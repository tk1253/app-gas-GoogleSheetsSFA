function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('[GAS]')
    .addSubMenu(
      ui.createMenu("GAS実行")
        .addItem('チェックした取引先のシートを作成', 'createNewSheetFromCheckedRows')
    )
    .addToUi();
}
