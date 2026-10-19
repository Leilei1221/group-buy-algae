const PRODUCT_LIST = [
  { id: 'p01', name: '黑松露噴霧油', price: 339 },
  { id: 'p02', name: '玄米油噴霧油', price: 339 },
  { id: 'p03', name: '亞麻籽油噴霧油', price: 339 },
  { id: 'p04', name: '葡萄籽油噴霧油', price: 339 },
  { id: 'p05', name: '酪梨油噴霧油', price: 339 },
  { id: 'p06', name: '黃金蒜酥', price: 99 },
  { id: 'p07', name: '韓式三切岩燒海苔（原味）', price: 130 },
  { id: 'p08', name: '韓式三切岩燒海苔（梅子）', price: 130 },
  { id: 'p09', name: '六色海藻', price: 329 },
  { id: 'p10', name: '海帶芽湯（蔬菜）', price: 230 },
  { id: 'p11', name: '海帶芽湯（青蔥）', price: 230 },
  { id: 'p12', name: '昆布鹽', price: 220 },
  { id: 'p13', name: '椎茸昆布鹽', price: 220 },
  { id: 'p14', name: '焙煎胡麻醬', price: 160 },
  { id: 'p15', name: '胡麻和風醬', price: 160 },
  { id: 'p16', name: '柚子和風醬', price: 160 },
  { id: 'p17', name: '昆布佃煮海苔醬', price: 160 },
  { id: 'p18', name: '昆布椎茸高湯包', price: 360 },
  { id: 'p19', name: '番茄昆布高湯包', price: 360 },
  { id: 'p20', name: '藻安！海藻隨身包', price: 360 },
  { id: 'p21', name: '焙煎胡麻隨身包', price: 130 },
  { id: 'p22', name: '胡麻和風隨身包', price: 130 },
  { id: 'p23', name: '柚子和風隨身包', price: 130 },
];

function getOrCreateSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  let ssId = props.getProperty('SPREADSHEET_ID');
  if (ssId) {
    try {
      return SpreadsheetApp.openById(ssId);
    } catch (e) {
      // Spreadsheet deleted, recreate
    }
  }
  const ss = SpreadsheetApp.create('藻作坊｜蒜蒜屋｜VIVO噴霧油 補貨團 - 訂單');
  props.setProperty('SPREADSHEET_ID', ss.getId());
  setupHeaders(ss.getActiveSheet());
  return ss;
}

function setupHeaders(sheet) {
  const headers = ['時間戳記', '姓名'];
  PRODUCT_LIST.forEach(p => headers.push(p.name + '\n$' + p.price));
  headers.push('訂購摘要', '總金額（元）', '備註及回覆');

  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1a5c2a')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
  sheet.setRowHeight(1, 48);

  // Column widths
  sheet.setColumnWidth(1, 155); // 時間
  sheet.setColumnWidth(2, 75);  // 姓名
  for (let i = 0; i < PRODUCT_LIST.length; i++) {
    sheet.setColumnWidth(3 + i, 78);
  }
  const summaryCol = 3 + PRODUCT_LIST.length;
  sheet.setColumnWidth(summaryCol, 220);     // 訂購摘要
  sheet.setColumnWidth(summaryCol + 1, 90);  // 總金額
  sheet.setColumnWidth(summaryCol + 2, 160); // 備註及回覆

  // Freeze
  sheet.setFrozenRows(1);
  sheet.setFrozenColumns(2);
}

function doPost(e) {
  const params = e.parameter;
  const name = params.name || '';
  const summary = params.summary || '';
  const total = parseInt(params.total || '0', 10);
  const note = params.note || '';

  const ss = getOrCreateSpreadsheet();
  const sheet = ss.getActiveSheet();

  // Check if headers exist
  if (!sheet.getRange(1, 3).getValue()) {
    setupHeaders(sheet);
  }

  const now = Utilities.formatDate(new Date(), 'Asia/Taipei', 'yyyy/MM/dd HH:mm:ss');
  const row = [now, name];

  PRODUCT_LIST.forEach(p => {
    const q = parseInt(params[p.id] || '0', 10);
    row.push(q > 0 ? q : '');
  });

  // Build reply message
  const itemLines = [];
  PRODUCT_LIST.forEach(p => {
    const q = parseInt(params[p.id] || '0', 10);
    if (q > 0) itemLines.push(p.name + ' ' + q + '份 單價' + p.price + '元');
  });
  const replyMsg = name + '你好，你所訂購的產品如下：' + itemLines.join('、') + '、總金額：' + total + '元' + (note ? '\n備註：' + note : '');

  row.push(summary, total, replyMsg);

  const newRow = sheet.getLastRow() + 1;
  sheet.getRange(newRow, 1, 1, row.length).setValues([row]);

  // Format: center qty columns
  const qtyRange = sheet.getRange(newRow, 3, 1, PRODUCT_LIST.length);
  qtyRange.setHorizontalAlignment('center');

  // Wrap summary
  const summaryCol = 3 + PRODUCT_LIST.length;
  sheet.getRange(newRow, summaryCol).setWrap(true);
  sheet.getRange(newRow, summaryCol + 2).setWrap(true);

  // Alternating row color
  const bgColor = (newRow % 2 === 0) ? '#e8f5ea' : '#ffffff';
  sheet.getRange(newRow, 1, 1, row.length).setBackground(bgColor);

  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSpreadsheetUrl() {
  const props = PropertiesService.getScriptProperties();
  const ssId = props.getProperty('SPREADSHEET_ID');
  if (ssId) {
    Logger.log('試算表網址: https://docs.google.com/spreadsheets/d/' + ssId);
  } else {
    Logger.log('尚未建立試算表，第一筆訂單送出後會自動建立');
  }
}
