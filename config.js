/**
 * 隅花相識網站設定
 * 完成 Google Apps Script 部署後，只要修改下方兩個網址。
 */
window.AMOR_CONFIG = Object.freeze({
  // Apps Script 網頁應用程式網址，必須以 /exec 結尾。
  orderApi: '',

  // setupFlowerBackend 執行記錄顯示的 Google 試算表網址。
  adminSheetUrl: '',

  instagramUrl: 'https://www.instagram.com/amorflores2018/',
  productRefreshMs: 120000,
});
