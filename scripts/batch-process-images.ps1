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

# 獲取所有子資料夾
$subFolders = Get-ChildItem -Path $folderPath -Directory

# 建立轉換報告
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
$reportFileName = "ConversionReport_$timestamp.md"
$reportFilePath = Join-Path -Path $folderPath -ChildPath $reportFileName

# 報告標頭
$report = @"
# 影像批次處理報告

- **掃描時間**: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
- **掃描資料夾**: $folderPath
- **輸出格式**: $outputFormat
- **JPG 品質**: $quality

## 子資料夾分析

共掃描 $($subFolders.Count) 個子資料夾:

| 序號 | 子資料夾名稱 | 可轉換影像 | 影像格式 |
|------|------------|-----------|---------|
"@

# 填充子資料夾分析
$index = 1
foreach ($folder in $subFolders) {
    $folderFiles = $imageFiles | Where-Object { $_.DirectoryName -eq $folder.FullName }
    $hasImages = $folderFiles.Count -gt 0
    $imageType = ""
    
    if ($hasImages) {
        $imageType = ($folderFiles | Select-Object -First 1).Extension
    }
    
    $hasImagesText = if ($hasImages) { "✓ ($($folderFiles.Count)個檔案)" } else { "✗" }
    $report += "`n| $index | $($folder.Name) | $hasImagesText | $imageType |"
    $index++
}

$report += @"

## 影像檔分析

已找到 $($imageFiles.Count) 個可處理的影像檔。

| 序號 | 檔案路徑 | 原始格式 | JPG 存在? | PNG 存在? | 檔案大小比對 | 狀態 |
|------|----------|----------|-----------|-----------|------------|------|
"@

# 填充影像檔分析
$index = 1
foreach ($file in $imageFiles) {
    $jpgExists = Test-Path ([System.IO.Path]::ChangeExtension($file.FullName, "jpg"))
    $pngExists = Test-Path ([System.IO.Path]::ChangeExtension($file.FullName, "png"))
    $jpgStatus = if ($jpgExists) { "✓" } else { "✗" }
    $pngStatus = if ($pngExists) { "✓" } else { "✗" }
    
    # 計算檔案大小比例
    $sizeRatio = "N/A"
    $status = "✅ 正常"
    
    if ($jpgExists) {
        $originalSize = (Get-Item $file.FullName).Length
        $jpgPath = [System.IO.Path]::ChangeExtension($file.FullName, "jpg")
        $jpgSize = (Get-Item $jpgPath).Length
        
        # 計算比例，並保留兩位小數
        if ($originalSize -gt 0) {
            $ratio = [math]::Round(($jpgSize / $originalSize) * 100, 2)
            $sizeRatio = "$ratio% ($([math]::Round($jpgSize / 1MB, 2)) MB / $([math]::Round($originalSize / 1MB, 2)) MB)"
            
            # 判斷是否異常
            if ($ratio -lt 10) {
                $status = "⚠️ JPG過小，可能有問題"
            }
        }
    }
    
    $report += "`n| $index | $($file.FullName) | $($file.Extension) | $jpgStatus | $pngStatus | $sizeRatio | $status |"
    $index++
}

# 寫入報告檔案
try {
    Set-Content -Path $reportFilePath -Value $report -Encoding UTF8
    Write-Host "`n已產生轉換前報告: $reportFilePath" -ForegroundColor Green
} catch {
    Write-Host "無法寫入報告檔案: $_" -ForegroundColor Red
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

# 更新報告檔
if (Test-Path $reportFilePath) {
    try {        # 搜尋更新後的狀態
        $processedJpgCount = 0
        $processedPngCount = 0
        $smallJpgCount = 0
        
        foreach ($file in $imageFiles) {
            $jpgPath = [System.IO.Path]::ChangeExtension($file.FullName, "jpg")
            $pngPath = [System.IO.Path]::ChangeExtension($file.FullName, "png")
            
            if (Test-Path $jpgPath) {
                $processedJpgCount++
                
                # 檢查檔案大小比例
                $originalSize = (Get-Item $file.FullName).Length
                $jpgSize = (Get-Item $jpgPath).Length
                
                if ($originalSize -gt 0 -and ($jpgSize / $originalSize) * 100 -lt 10) {
                    $smallJpgCount++
                }
            }
            
            if (Test-Path $pngPath) {
                $processedPngCount++
            }
        }
        
        $completionInfo = @"

## 處理結果

- **完成時間**: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
- **處理檔案總數**: $total
- **輸出格式**: $outputFormat

### 轉換統計
- **輸入檔案總數**: $($imageFiles.Count)
- **子資料夾總數**: $($subFolders.Count)
- **JPG 檔案總數**: $processedJpgCount
- **PNG 檔案總數**: $processedPngCount
- **檢測到異常小的 JPG 檔案**: $smallJpgCount

### 備註
- 如果檔案已存在且大小正常，則跳過處理
- 若找到異常小的 JPG 檔案（小於原始檔案的 1/10），則重新轉換
- 定位檔案會一併轉換（.tfw -> .jgw / .pgw）

"@

        # 增加異常檔案清單
        if ($smallJpgCount -gt 0) {
            $abnormalFilesInfo = @"
### 異常檔案清單

以下檔案的 JPG 大小異常（小於原始檔案的 10%），可能需要重新轉換：

| 序號 | 檔案路徑 | 原始大小 | JPG 大小 | 比例 |
|------|----------|----------|----------|------|
"@
            
            $abnormalIndex = 1
            foreach ($file in $imageFiles) {
                $jpgPath = [System.IO.Path]::ChangeExtension($file.FullName, "jpg")
                if (Test-Path $jpgPath) {
                    $originalSize = (Get-Item $file.FullName).Length
                    $jpgSize = (Get-Item $jpgPath).Length
                    
                    if ($originalSize -gt 0 -and ($jpgSize / $originalSize) * 100 -lt 10) {
                        $ratio = [math]::Round(($jpgSize / $originalSize) * 100, 2)
                        $originalSizeMB = [math]::Round($originalSize / 1MB, 2)
                        $jpgSizeMB = [math]::Round($jpgSize / 1MB, 2)
                        
                        $abnormalFilesInfo += "`n| $abnormalIndex | $($file.FullName) | $originalSizeMB MB | $jpgSizeMB MB | $ratio% |"
                        $abnormalIndex++
                    }
                }
            }
            
            Add-Content -Path $reportFilePath -Value $abnormalFilesInfo -Encoding UTF8
        }
        else {
            Add-Content -Path $reportFilePath -Value "`n### 無異常檔案檢測" -Encoding UTF8
        }
"@
        Add-Content -Path $reportFilePath -Value $completionInfo -Encoding UTF8
        Write-Host "已更新處理報告: $reportFilePath" -ForegroundColor Green
        
        # 嘗試在檔案管理器中打開報告檔
        Start-Process "explorer.exe" -ArgumentList "/select,`"$reportFilePath`""
    }
    catch {
        Write-Host "無法更新報告檔案: $_" -ForegroundColor Red
    }
}
