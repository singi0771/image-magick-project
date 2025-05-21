# PowerShell script for batch processing images using ImageMagick
# 批次影像處理腳本

param (
    [Parameter(Mandatory=$true)]
    [string]$folderPath,
    [ValidateSet('jpg', 'png', 'both')]
    [string]$outputFormat = 'jpg',
    [int]$quality = 90
)

# 支援的影像格式
$supportedFormats = @(".png", ".tiff", ".bmp", ".gif", ".webp",".tif")

# 驗證資料夾路徑
if (-Not (Test-Path $folderPath)) {
    Write-Host "錯誤: 指定的資料夾路徑不存在" -ForegroundColor Red
    exit 1
}

# 搜尋所有支援的影像檔
$imageFiles = Get-ChildItem -Path $folderPath -Recurse | 
    Where-Object { $supportedFormats -contains $_.Extension.ToLower() }

if ($imageFiles.Count -eq 0) {
    Write-Host "未找到支援的影像檔" -ForegroundColor Yellow
    exit 0
}

# 顯示找到的檔案清單
Write-Host "`n找到以下影像檔:" -ForegroundColor Cyan
$imageFiles | ForEach-Object { Write-Host $_.FullName }

# 確認是否繼續
$confirmation = Read-Host "`n找到 $($imageFiles.Count) 個檔案。是否要繼續處理? [Y/N] (預設: Y)"
if ($confirmation -and $confirmation.ToUpper() -eq 'N') {
    Write-Host "操作已取消" -ForegroundColor Yellow
    exit 0
}

# 處理進度計數
$total = $imageFiles.Count
$current = 0

# 批次處理影像
foreach ($file in $imageFiles) {
    $current++
    Write-Progress -Activity "處理影像中" -Status "處理: $($file.Name)" `
        -PercentComplete (($current / $total) * 100)
    
    Write-Host "`n處理中 ($current/$total): $($file.FullName)" -ForegroundColor Cyan

    # 定義要處理的輸出格式
    $formatsToProcess = if ($outputFormat -eq 'both') { @('jpg', 'png') } else { @($outputFormat) }
    
    # 檢查並處理定位檔
    $worldFiles = @(
        [System.IO.Path]::ChangeExtension($file.FullName, "tfw"),
        [System.IO.Path]::ChangeExtension($file.FullName, "tifw"),
        [System.IO.Path]::ChangeExtension($file.FullName, "wld")
    )

    $foundWorldFile = $null
    # 找到第一個存在的定位檔
    foreach ($worldFile in $worldFiles) {
        if (Test-Path $worldFile) {
            $foundWorldFile = $worldFile
            break
        }
    }

    # 處理每種輸出格式
    foreach ($format in $formatsToProcess) {
        $outputFile = [System.IO.Path]::Combine(
            $file.DirectoryName, 
            [System.IO.Path]::GetFileNameWithoutExtension($file.Name) + ".$format"
        )

        $needConversion = $true
        # 檢查目標檔案是否已存在並進行大小比較
        if (Test-Path $outputFile) {
            $originalSize = (Get-Item $file.FullName).Length
            $convertedSize = (Get-Item $outputFile).Length
            
            # 如果轉換後的檔案小於原始檔案的 1/10，判定為可能轉換失敗
            if ($format -eq 'jpg' -and $convertedSize -lt ($originalSize / 10)) {
                Write-Host "發現異常小的轉換檔案，準備重新轉換: $outputFile" -ForegroundColor Yellow
                $needConversion = $true
            } else {
                Write-Host "目標檔案已存在且大小正常，跳過轉換: $outputFile" -ForegroundColor Yellow
                $needConversion = $false
            }
        }

        if (-not $needConversion) {
            continue
        }

        # 設定對應的世界檔案副檔名
        $worldFileExt = if ($format -eq 'jpg') { 'jgw' } else { 'pgw' }

        # 如果找到定位檔，則複製到對應的輸出格式
        if ($foundWorldFile) {
            $newWorldFile = [System.IO.Path]::ChangeExtension($outputFile, $worldFileExt)
            try {
                Copy-Item -Path $foundWorldFile -Destination $newWorldFile -Force
                Write-Host "已複製定位檔: $newWorldFile" -ForegroundColor Green
            }
            catch {
                Write-Host "複製定位檔失敗: $foundWorldFile" -ForegroundColor Red
                Write-Host "錯誤訊息: $_" -ForegroundColor Red
            }
        }
          try {
            & magick $file.FullName -resize "100%" -quality $quality $outputFile
            Write-Host "已成功轉換為 ${format}，輸出: $outputFile" -ForegroundColor Green
        }
        catch {
            Write-Host "處理失敗 (${format})，檔案: $($file.FullName)" -ForegroundColor Red
            Write-Host "錯誤訊息: $_" -ForegroundColor Red
        }
    }
}

Write-Host "`n批次處理完成!" -ForegroundColor Green
