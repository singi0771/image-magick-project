# 影像批次處理工具

這個工具可以批次將各種格式的影像檔案轉換為 JPG 或 PNG 格式，並且支援地理資訊檔案（World File）的轉換。

## 環境需求

1. PowerShell 7.0 或更新版本
2. ImageMagick

### 安裝 ImageMagick

1. 使用 winget（推薦）：
```powershell
winget install ImageMagick.ImageMagick
```

2. 或使用 Chocolatey：
```powershell
choco install imagemagick.app
```

3. 或從官方網站下載安裝檔：
   - 前往 [ImageMagick 官方網站](https://imagemagick.org/script/download.php)
   - 下載並執行安裝檔
   - 安裝時請確保選擇「Add application directory to your system path」

### 驗證安裝

在 PowerShell 中執行以下命令，確認 ImageMagick 已正確安裝：
```powershell
magick -version
```

## 支援的檔案格式

### 輸入格式
- PNG (.png)
- TIFF (.tiff, .tif)
- BMP (.bmp)
- GIF (.gif)
- WebP (.webp)

### 輸出格式
- JPEG (.jpg)
- PNG (.png)
- 可同時輸出兩種格式

### 地理資訊檔案支援
- 支援的輸入格式：.tfw, .tifw, .wld
- 輸出格式：
  - JPG 對應 .jgw
  - PNG 對應 .pgw

## 使用方法

### 基本用法

1. 轉換為 JPG（預設）：
```powershell
# 方法1：使用完整路徑執行（推薦）
D:\3_CodingProject\image-magick-project\scripts\batch-process-images.ps1 -folderPath "圖片資料夾路徑"

# 方法2：先切換到腳本目錄
cd D:\3_CodingProject\image-magick-project\scripts
.\batch-process-images.ps1 -folderPath "圖片資料夾路徑"
```

2. 轉換為 PNG：
```powershell
.\scripts\batch-process-images.ps1 -folderPath "圖片資料夾路徑" -outputFormat png
```

3. 同時轉換為 JPG 和 PNG：
```powershell
.\scripts\batch-process-images.ps1 -folderPath "圖片資料夾路徑" -outputFormat both
```

### 進階選項

- 設定 JPG 壓縮品質（1-100）：
```powershell
.\scripts\batch-process-images.ps1 -folderPath "圖片資料夾路徑" -quality 95
```

### 參數說明

- `-folderPath`：必要參數，指定要處理的資料夾路徑
- `-outputFormat`：可選參數，指定輸出格式
  - `jpg`：輸出 JPG 格式（預設）
  - `png`：輸出 PNG 格式
  - `both`：同時輸出 JPG 和 PNG 格式
- `-quality`：可選參數，設定 JPG 壓縮品質（1-100，預設 90）

## 功能特點

1. 遞迴處理：自動處理指定資料夾及其所有子資料夾中的圖片
2. 保留地理資訊：自動轉換和複製相關的定位檔案
3. 批次處理：一次處理多個檔案
4. 進度顯示：即時顯示處理進度和狀態
5. 錯誤處理：完整的錯誤處理和提示
6. 使用者確認：在開始處理前顯示檔案清單並請求確認

## 注意事項

1. 轉換後的檔案會存放在與原始檔案相同的資料夾中
2. 如果輸出檔案已存在，將會被覆寫
3. 處理大量檔案時，建議先進行小規模測試
4. 確保有足夠的硬碟空間存放輸出檔案

### 中斷執行

如需中斷正在執行的處理程序：
1. 按下 `Ctrl + C` 可以中斷當前執行的 PowerShell 指令
2. 如果 ImageMagick 工具仍在背景執行：
   ```powershell
   # 列出所有 ImageMagick 相關程序
   Get-Process magick* | Select-Object Id, ProcessName
   
   # 結束特定程序（將 <ID> 替換為實際的處理程序 ID）
   Stop-Process -Id <ID>
   
   # 或結束所有 ImageMagick 相關程序
   Get-Process magick* | Stop-Process
   ```

## 錯誤處理

如果遇到「無法執行指令碼」的錯誤，請在 PowerShell 中執行：
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

## 授權資訊

本工具採用 MIT 授權。

## 開發者資訊

如果您想要修改或擴展此工具的功能，可以查看 `scripts` 資料夾中的 PowerShell 腳本。主要檔案說明：

```
image-magick-project/
├── scripts/
│   ├── batch-process-images.ps1    # 批次處理主腳本
│   └── batch-process-images.ps1.bak # 腳本備份
└── README.md                       # 本說明檔案
```