function encodeInvoiceRow(fields, sheetRow) {
  return fields.map(value => {
    const text = String(value == null ? '' : value);
    if (/[\r\n]/.test(text)) {
      throw new Error('第 ' + sheetRow + ' 列的發票欄位不能換行。');
    }
    return /[",]/.test(text) ? '"' + text.replace(/"/g, '""') + '"' : text;
  }).join(',');
}

function generateInvoiceCsv() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert('當前沒有選定工作表。');
    return;
  }
  const sheetName = sheet.getName();

  const dataRange = sheet.getDataRange();
  const data = dataRange.getValues();
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  
  let serial_number_increment = 0;
  
  // 表頭記錄 這裡要改 ezPay電子發票會員編號、商店代號 這兩個
  const invoiceRows = [];
  
  try {
    const header = ['H', 'INVO', 'C12231771956', '326376941', sheetName, ...Array(11).fill('')];
    invoiceRows.push(encodeInvoiceRow(header, 1));
    data.slice(1).forEach((row, index) => {
      const sheetRow = index + 2;
      const serial_number = sheetName + ('00' + (serial_number_increment + 1)).slice(-3);
      serial_number_increment++;
      
      const sale_type = row[0];
      let sale_row;
      if (sale_type === 'B2B') {
        sale_row = ['S', serial_number, sale_type, row[1], row[2], row[4], '', '', '', '', 'Y', '1', '5'];
      } else if (sale_type === 'B2C') {
        sale_row = ['S', serial_number, sale_type, '', row[3], row[4], '', '', '', '', 'Y', '1', '5'];
      } else {
        throw new Error('第 ' + sheetRow + ' 列的發票類型必須是 B2B 或 B2C。');
      }
      
      const product_price = row[5];
      const sale_amount = Math.round(product_price / 1.05);
      const tax_amount = product_price - sale_amount;
      sale_row = sale_row.concat([sale_amount, tax_amount, product_price]);
      invoiceRows.push(encodeInvoiceRow(sale_row, sheetRow));
      
      const product_name = row[6];
      const quantity = row[7];
      const unit = row[8];
      const item_total = quantity * product_price;
      const item_row = ['I', serial_number, product_name, quantity, unit, product_price, item_total];
      invoiceRows.push(encodeInvoiceRow(item_row, sheetRow));
    });
  } catch (error) {
    SpreadsheetApp.getUi().alert('未產生發票檔案：' + error.message);
    return;
  }

  // 同一區的最下面我有設定生成的檔案檔名，這裡也要改成你的商店代號
  const fileName = '326376941_' + sheetName + '.csv';
  const blob = Utilities.newBlob(invoiceRows.join('\n') + '\n', 'text/csv', fileName);
  
  // 指定文件夾 ID 後面改成你雲端後面的網址 https://drive.google.com/drive/u/1/folders/ 這串後面的網址
  const folderId = '1hwnEhYitCMIdaWUkzr6uFrj89eDCm';
  const folder = DriveApp.getFolderById(folderId);
  const file = folder.createFile(blob);
  const fileUrl = file.getUrl();
  
  const lastColumn = sheet.getLastColumn();
  sheet.getRange(1, lastColumn + 1).setValue('執行時間');
  sheet.getRange(2, lastColumn + 1, data.length - 1, 1).setValue(timestamp);
  sheet.getRange(1, lastColumn + 2).setValue('檔案位置');
  sheet.getRange(2, lastColumn + 2).setValue(fileUrl);
  
  SpreadsheetApp.getUi().alert('CSV file created: ' + fileUrl);
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('Generate Invoice CSV', 'generateInvoiceCsv')
    .addToUi();
}
