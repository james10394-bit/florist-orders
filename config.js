/**
 * 隅花相識網站設定
 * 完成 Google Apps Script 部署後，只要修改下方兩個網址。
 */
window.AMOR_CONFIG = Object.freeze({
  // Apps Script 網頁應用程式網址，必須以 /exec 結尾。
  orderApi: 'https://script.google.com/macros/s/AKfycbwfAUnpAgc54KDfvpmbQsomUpfBZ0wve0v6ad3fLFgiN7NZ0eh2UdPziDFrzJL-LYka/exec',

  // setupFlowerBackend 執行記錄顯示的 Google 試算表網址。
  adminSheetUrl: 'https://docs.google.com/spreadsheets/d/148vwIZXG8HnB5yQStNhMz47EVplQd6tdHXX345OUIow/edit?gid=0#gid=0',

  instagramUrl: 'https://www.instagram.com/amorflores2018/',
  productRefreshMs: 120000,
});
