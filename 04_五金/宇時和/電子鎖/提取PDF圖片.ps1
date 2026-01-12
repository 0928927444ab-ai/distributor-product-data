# PDF 圖片提取輔助腳本
# 此腳本會打開相關網頁幫助你提取 PDF 圖片

Write-Host "=== PDF 圖片提取助手 ===" -ForegroundColor Cyan
Write-Host ""

$pdfPath = "C:\Users\User\Desktop\總代理產品資料\04_五金\宇時和\電子鎖\(電子鎖)凱恩斯\KA酒店锁图册修订版.pdf"
$imagesPath = "C:\Users\User\Desktop\總代理產品資料\04_五金\宇時和\電子鎖\images"

# 確保 images 資料夾存在
if (-not (Test-Path $imagesPath)) {
    New-Item -ItemType Directory -Path $imagesPath -Force | Out-Null
    Write-Host "已創建 images 資料夾" -ForegroundColor Green
}

Write-Host "PDF 檔案位置：" -ForegroundColor Yellow
Write-Host $pdfPath
Write-Host ""
Write-Host "圖片儲存位置：" -ForegroundColor Yellow
Write-Host $imagesPath
Write-Host ""

# 提供選項
Write-Host "請選擇提取方式：" -ForegroundColor Cyan
Write-Host "1. 打開線上工具 (iLovePDF) - 推薦"
Write-Host "2. 打開線上工具 (PDF.io)"
Write-Host "3. 打開 PDF 檔案（手動截圖）"
Write-Host "4. 打開 images 資料夾"
Write-Host "5. 全部打開"
Write-Host ""

$choice = Read-Host "請輸入選項 (1-5)"

switch ($choice) {
    "1" {
        Start-Process "https://www.ilovepdf.com/pdf_to_jpg"
        Write-Host "已打開 iLovePDF，請上傳 PDF 檔案並下載圖片" -ForegroundColor Green
    }
    "2" {
        Start-Process "https://pdf.io/tw/pdf2jpg/"
        Write-Host "已打開 PDF.io，請上傳 PDF 檔案並下載圖片" -ForegroundColor Green
    }
    "3" {
        if (Test-Path $pdfPath) {
            Start-Process $pdfPath
            Write-Host "已打開 PDF 檔案，請使用 Win+Shift+S 截圖" -ForegroundColor Green
        } else {
            Write-Host "找不到 PDF 檔案" -ForegroundColor Red
        }
    }
    "4" {
        Start-Process explorer.exe $imagesPath
        Write-Host "已打開 images 資料夾" -ForegroundColor Green
    }
    "5" {
        Start-Process "https://www.ilovepdf.com/pdf_to_jpg"
        Start-Process explorer.exe $imagesPath
        if (Test-Path $pdfPath) {
            Start-Process $pdfPath
        }
        Write-Host "已打開所有相關視窗" -ForegroundColor Green
    }
    default {
        Write-Host "無效選項" -ForegroundColor Red
    }
}

Write-Host ""
Write-Host "=== 圖片命名建議 ===" -ForegroundColor Cyan
Write-Host "按照產品型號命名，例如："
Write-Host "  - YSHLF832.png (第5頁)"
Write-Host "  - YSHLF855-23.png (第6頁)"
Write-Host "  - YSHLT837.png (第22頁)"
Write-Host "  - lock_5B.png (第50頁)"
Write-Host ""
Write-Host "完成後，圖片會自動顯示在 Markdown 報價單中"
Write-Host ""

Read-Host "按 Enter 結束"
