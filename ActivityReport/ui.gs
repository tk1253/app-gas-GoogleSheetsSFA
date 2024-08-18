function onOpen() {
  let ui = SpreadsheetApp.getUi();
  ui.createMenu('[GAS]')
    .addSubMenu(
      ui.createMenu("GAS実行")
        .addItem('活動内容の更新(取得)', 'executeActivityReports')
    )
    .addToUi();
}
