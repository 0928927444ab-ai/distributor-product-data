const XLSX = require('xlsx');

const filePath = String.raw`c:\Users\User\Desktop\總代理產品資料\04_五金\威斯特-VSTE\門用五金\03_訂單管理\20251114第二批庫存\摩玛(威思特产品）报价单2025-11-14(調整後)(2).xlsx`;

try {
    const workbook = XLSX.readFile(filePath);
    console.log('工作表名稱:', workbook.SheetNames);

    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];

    console.log('資料範圍:', worksheet['!ref']);

    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

    console.log('\n=== 完整資料 ===\n');
    data.forEach((row, idx) => {
        console.log('Row ' + (idx + 1) + ':', JSON.stringify(row));
    });

} catch (error) {
    console.error('錯誤:', error.message);
}
