const XLSX = require('xlsx');

const filePath = 'c:\\Users\\User\\Desktop\\總代理產品資料\\04_五金\\威斯特-VSTE\\門用五金\\05_產品資料\\型錄\\威士特\\VSTE威思特分销单价表2025(1).xlsx';

// 讀取 Excel 檔案
const workbook = XLSX.readFile(filePath);

console.log('='.repeat(100));
console.log('VSTE威思特分销单价表2025(1).xlsx');
console.log('='.repeat(100));
console.log(`工作表數量: ${workbook.SheetNames.length}`);
console.log(`工作表列表: ${workbook.SheetNames.join(', ')}`);
console.log();

// 遍歷每個工作表
workbook.SheetNames.forEach(sheetName => {
    console.log('='.repeat(100));
    console.log(`工作表: ${sheetName}`);
    console.log('='.repeat(100));

    const worksheet = workbook.Sheets[sheetName];
    const range = XLSX.utils.decode_range(worksheet['!ref'] || 'A1');

    const rows = range.e.r - range.s.r + 1;
    const cols = range.e.c - range.s.c + 1;
    console.log(`範圍: ${rows} 行 x ${cols} 列`);
    console.log('-'.repeat(100));

    // 將工作表轉換為 JSON 格式 (以陣列形式，包含標題列)
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

    // 輸出每一行
    data.forEach((row, index) => {
        // 檢查是否有內容
        const hasContent = row.some(cell => cell !== '' && cell !== null && cell !== undefined);
        if (hasContent) {
            const rowStr = row.map(cell => {
                if (cell === null || cell === undefined || cell === '') return '';
                return String(cell);
            }).join(' | ');
            console.log(`第${String(index + 1).padStart(3, ' ')}行: ${rowStr}`);
        }
    });

    console.log();
});

console.log('讀取完成！');
