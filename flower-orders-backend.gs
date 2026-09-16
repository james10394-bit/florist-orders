/**
 * 隅花相識 Amor Flores Design 顧客預訂＋商品後端 v2.2.0
 * 執行 setupFlowerBackend() 一次，再「部署 → 新增部署 → 網頁應用程式」。
 * 執行身分：我；存取權：任何人。
 */
const SHOP = Object.freeze({
  calendarName: '隅花相識｜顧客預訂',
  sheetName: '隅花相識｜顧客預訂資料',
  timezone: 'Asia/Taipei',
});

function setupFlowerBackend() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('CALENDAR_ID')) {
    const calendar = CalendarApp.createCalendar(SHOP.calendarName, {timeZone: SHOP.timezone});
    props.setProperty('CALENDAR_ID', calendar.getId());
  }
  let sheetFile;
  if (!props.getProperty('SHEET_ID')) {
    sheetFile = SpreadsheetApp.create(SHOP.sheetName);
    props.setProperty('SHEET_ID', sheetFile.getId());
  } else {
    sheetFile = SpreadsheetApp.openById(props.getProperty('SHEET_ID'));
  }
  ensureOrdersSheet_(sheetFile);
  ensureProductsSheet_(sheetFile);
  console.log('後端已準備完成。接著請部署為網頁應用程式。');
  console.log('試算表：https://docs.google.com/spreadsheets/d/' + props.getProperty('SHEET_ID'));
}

function doGet(e) {
  if (e && e.parameter && e.parameter.action === 'catalog') {
    return catalogResponse_(e.parameter.callback);
  }
  return HtmlService.createHtmlOutput('隅花相識 Amor Flores Design 訂單後端運作中');
}

function ensureOrdersSheet_(sheetFile) {
  let sheet = sheetFile.getSheetByName('訂單');
  if (!sheet) {
    sheet = sheetFile.getSheets()[0];
    if (sheet.getLastRow() === 0) sheet.setName('訂單');
    else sheet = sheetFile.insertSheet('訂單');
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['建立時間','訂單編號','狀態','商品','單價','數量','預算','用途','色系','避免花材','交付方式','交付日期','時段','訂購人','訂購人電話','Email','聯絡方式','收件人','收件人電話','地址','卡片','備註','日曆事件ID']);
    sheet.setFrozenRows(1);
  }
}

function ensureProductsSheet_(sheetFile) {
  let sheet = sheetFile.getSheetByName('商品');
  if (sheet) return;
  sheet = sheetFile.insertSheet('商品');
  sheet.appendRow(['上架','排序','商品名稱','副標','價格','分類','圖片網址','按鈕文字']);
  sheet.getRange(2, 1, 5, 8).setValues([
    [true,1,'晨光粉玫花束','柔粉・奶油白',1680,'生日、告白／紀念、日常','','選這束'],
    [true,2,'炙熱紅玫花束','酒紅・裸粉',2280,'告白／紀念','','選這束'],
    [true,3,'小太陽鮮花束','橙黃・嫩綠',1580,'生日、日常','','選這束'],
    [true,4,'森系開幕盆花','白綠・大器層次',3200,'開幕','','選這款'],
    [true,5,'暮色紫霧花束','霧紫・藕粉',1980,'生日、告白／紀念、日常','','選這束'],
  ]);
  sheet.getRange('A2:A1000').insertCheckboxes();
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 8);
}

function catalogResponse_(callback) {
  const callbackName = String(callback || 'renderFlowerCatalog');
  if (!/^[A-Za-z_$][0-9A-Za-z_$\.]{0,80}$/.test(callbackName)) {
    return ContentService.createTextOutput('/* invalid callback */')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  const cache = CacheService.getScriptCache();
  let json = cache.get('PUBLIC_PRODUCT_CATALOG');
  if (!json) {
    const props = PropertiesService.getScriptProperties();
    const sheet = SpreadsheetApp.openById(props.getProperty('SHEET_ID')).getSheetByName('商品');
    if (!sheet) throw new Error('找不到商品分頁，請再執行一次 setupFlowerBackend。');
    const rows = sheet.getDataRange().getDisplayValues();
    const headers = rows.shift();
    const column = {};
    headers.forEach(function(header, index) { column[header] = index; });
    const products = rows.map(function(row) {
      const enabled = String(row[column['上架']] || '').toUpperCase();
      if (!['TRUE','是','Y','YES','1','上架'].includes(enabled)) return null;
      return {
        order: Number(row[column['排序']]) || 9999,
        name: String(row[column['商品名稱']] || '').trim(),
        subtitle: String(row[column['副標']] || '').trim(),
        price: Number(String(row[column['價格']] || '').replace(/[^0-9.]/g, '')) || 0,
        category: String(row[column['分類']] || '').trim(),
        imageUrl: normalizeImageUrl_(row[column['圖片網址']]),
        buttonText: String(row[column['按鈕文字']] || '選這束').trim(),
      };
    }).filter(function(product) { return product && product.name; })
      .sort(function(a, b) { return a.order - b.order; });
    json = JSON.stringify({ok:true,updatedAt:new Date().toISOString(),refreshSeconds:120,products:products});
    cache.put('PUBLIC_PRODUCT_CATALOG', json, 60);
  }
  return ContentService.createTextOutput(callbackName + '(' + json.replace(/</g, '\\u003c') + ');')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function normalizeImageUrl_(value) {
  const url = String(value || '').trim();
  if (!url) return '';
  const driveId = (url.match(/\/file\/d\/([a-zA-Z0-9_-]+)/) || url.match(/[?&]id=([a-zA-Z0-9_-]+)/) || [])[1];
  if (driveId) return 'https://drive.google.com/thumbnail?id=' + driveId + '&sz=w1200';
  return /^https:\/\//i.test(url) ? url : '';
}

function doPost(e) {
  try {
    if (!e || !e.parameter) throw new Error('沒有收到訂單資料。');
    const p = e.parameter;
    if (p.website) return response_(false, '', '無法送出。');
    validate_(p);

    const props = PropertiesService.getScriptProperties();
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const dedupeKey = 'ORDER_' + p.orderId;
      if (props.getProperty(dedupeKey)) return response_(true, p.orderId, '此訂單已建立。');

      const calendar = CalendarApp.getCalendarById(props.getProperty('CALENDAR_ID'));
      const interval = interval_(p.deliveryDate, p.slot);
      const title = '[待確認] ' + p.method + '｜' + p.product + ' × ' + p.quantity + '｜' + p.customerName;
      const description = [
        '訂單編號：' + p.orderId,
        '商品：' + p.product + ' × ' + p.quantity,
        '單價：' + (p.unitPrice || '客製報價'),
        '預算：' + (p.budget || ''),
        '用途：' + p.occasion,
        '色系：' + p.colors,
        '避免：' + (p.avoid || ''),
        '',
        '訂購人：' + p.customerName + '｜' + p.customerPhone,
        'Email：' + (p.email || ''),
        '聯絡：' + p.contact,
        '',
        '收件人：' + (p.recipientName || '') + '｜' + (p.recipientPhone || ''),
        '地址：' + (p.address || ''),
        '',
        '卡片：' + (p.card || ''),
        '備註：' + (p.notes || ''),
      ].join('\n');
      const event = calendar.createEvent(title, interval.start, interval.end, {
        description: description,
        location: p.method === '店家配送' ? p.address : '門市取貨',
      });
      event.addPopupReminder(24 * 60);
      event.addPopupReminder(120);

      const sheet = SpreadsheetApp.openById(props.getProperty('SHEET_ID')).getSheetByName('訂單');
      sheet.appendRow([new Date(),p.orderId,'待確認',p.product,p.unitPrice || '',p.quantity,p.budget || '',p.occasion,p.colors,p.avoid || '',p.method,p.deliveryDate,p.slot,p.customerName,p.customerPhone,p.email || '',p.contact,p.recipientName || '',p.recipientPhone || '',p.address || '',p.card || '',p.notes || '',event.getId()]);
      props.setProperty(dedupeKey, event.getId());
      return response_(true, p.orderId, '預訂需求已送出。');
    } finally {
      lock.releaseLock();
    }
  } catch (error) {
    console.error(error);
    return response_(false, '', error.message || '送出失敗。');
  }
}

function validate_(p) {
  ['orderId','product','quantity','occasion','colors','method','deliveryDate','slot','customerName','customerPhone','contact'].forEach(function(key) {
    if (!String(p[key] || '').trim()) throw new Error('必填資料不完整：' + key);
  });
  if (!/^FL-[A-Z0-9-]{8,40}$/.test(p.orderId)) throw new Error('訂單編號格式錯誤。');
  if (!/^[0-9+()\-\s]{8,20}$/.test(p.customerPhone)) throw new Error('訂購人電話格式錯誤。');
  if (Number(p.quantity) < 1 || Number(p.quantity) > 20) throw new Error('數量必須是 1～20。');
  if (p.method === '店家配送' && (!p.address || !p.recipientName || !p.recipientPhone)) throw new Error('配送訂單請填完整收件資料。');
  const requested = new Date(p.deliveryDate + 'T00:00:00+08:00');
  const today = new Date(Utilities.formatDate(new Date(), SHOP.timezone, 'yyyy-MM-dd') + 'T00:00:00+08:00');
  if (requested <= today) throw new Error('交付日期至少需選擇明天。');
}

function interval_(date, slot) {
  const slots = {'上午 09:00–12:00':['09:00','12:00'],'下午 12:00–15:00':['12:00','15:00'],'下午 15:00–18:00':['15:00','18:00'],'晚上 18:00–21:00':['18:00','21:00']};
  if (!slots[slot]) throw new Error('無法辨識交付時段。');
  return {start:new Date(date + 'T' + slots[slot][0] + ':00+08:00'),end:new Date(date + 'T' + slots[slot][1] + ':00+08:00')};
}

function response_(ok, orderId, message) {
  const payload = JSON.stringify({app:'flower-order',ok:ok,orderId:orderId,message:message}).replace(/</g, '\\u003c');
  return HtmlService.createHtmlOutput('<!doctype html><meta charset="utf-8"><script>parent.postMessage(' + payload + ',"*")<\/script>');
}
