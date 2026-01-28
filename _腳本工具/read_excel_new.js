const XLSX = require('xlsx');

// 檔案 1: 威斯特門用五金裝箱單
const file1 = String.raw`C:\Users\User\Desktop\總代理產品資料\00_確認落地產品\威斯特_門用五金\03_訂單管理\20251201下單資料\台灣 摩瑪建材 装箱单 12-07.xlsx`;

// 檔案 2: 信德成輕奢石裝箱單
const file2 = String.raw`C:\Users\User\Desktop\總代理產品資料\00_確認落地產品\信德成_輕奢石\03_訂單管理\已出貨\台湾单子装箱单(2).xls`;

function readExcelFile(filePath, fileName) {
    console.log('='.repeat(80));
    console.log('檔案:', fileName);
    console.log('='.repeat(80));

    try {
        const workbook = XLSX.readFile(filePath);
        console.log('工作表數量:', workbook.SheetNames.length);
        console.log('工作表名稱:', workbook.SheetNames);
        console.log();

        workbook.SheetNames.forEach(sheetName => {
            console.log('--- 工作表:', sheetName, '---');
            const worksheet = workbook.Sheets[sheetName];
            console.log('資料範圍:', worksheet['!ref']);

            // 取得合併儲存格資訊
            if (worksheet['!merges']) {
                console.log('合併儲存格數量:', worksheet['!merges'].length);
            }

            const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
            console.log('資料列數:', data.length);
            console.log();

            console.log('前 35 列內容:');
            data.slice(0, 35).forEach((row, idx) => {
                console.log('Row ' + (idx + 1) + ':', JSON.stringify(row));
            });
            console.log();
        });
    } catch (error) {
        console.error('讀取錯誤:', error.message);
    }
}

readExcelFile(file1, '台灣 摩瑪建材 装箱单 12-07.xlsx');
console.log('\n\n');
readExcelFile(file2, '台湾单子装箱单(2).xls');
