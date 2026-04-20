# TGOS Geocoder

`TGOS Geocoder` 是一個以 Electron 製作的桌面工具，專門用來整理地址資料、建立候選地址，並透過 TGOS 查詢座標與匹配結果。它適合處理來自 Excel 或 CSV 的批次地址資料，將清洗、候選地址組合、地理編碼與匯出流程集中在同一個工作介面中。

## 功能特色

- 匯入 `Excel` 與 `CSV` 檔案，且每個工作表會自動展開成可操作分頁
- 支援欄位清理
  - 去除前後空白
  - 連續空白壓縮
  - 全形空白轉半形
  - 全形字元轉半形
  - 移除換行
  - 自訂正則替換規則
- 支援將清理設定儲存成腳本，並在腳本庫中重複套用
- 可由多個欄位組合出 `candidate_address`
- 支援多組備選候選地址格式，當主要格式查無結果時會依序重試
- 透過 TGOS 進行單筆與批次 geocoding
- 支援只處理勾選列，或重新處理已成功／未成功匹配的資料
- 內建進度條、暫停與停止控制
- 提供地圖編輯視窗，以 OpenStreetMap 底圖檢視已帶座標的資料
- 可將處理後資料匯出為 `xlsx` 或 `csv`
- 會保存工作區狀態與 TGOS 查詢所需快取資料，方便下次續作

## 使用情境

這個工具特別適合以下工作：

- 整理需要批次標準化的地址欄位
- 從多欄位資料組合出完整地址
- 將地址送往 TGOS 做批次地理編碼
- 對查詢結果進行人工檢視與後續匯出

## 技術架構

- 桌面框架：`Electron`
- 資料請求：`axios`
- HTML 解析：`cheerio`
- 檔案讀寫：`xlsx`
- 本地快取：Node 內建 `sqlite` (`DatabaseSync`)

## 系統需求

- Node.js 20 以上
- npm 10 以上
- macOS 或 Windows

註：專案目前已配置 `electron-builder` 的 macOS 與 Windows 打包設定。

## 安裝與啟動

1. 安裝依賴

```bash
npm install
```

2. 啟動桌面程式

```bash
npm run dev
```

或

```bash
npm start
```

## 打包

打包全部預設平台設定：

```bash
npm run dist
```

只打包 macOS：

```bash
npm run dist:mac
```

只打包 Windows：

```bash
npm run dist:win
```

產物會輸出到 `dist/`。

## 使用流程

1. 在首頁匯入 `Excel / CSV`
2. 每個工作表會變成一個分頁，可先在檔案中心整理資料
3. 進入資料分頁後，先用左側「欄位清理」處理原始欄位
4. 在「候選地址」頁籤選擇欄位並組合出 `candidate_address`
5. 在「Geocoding」頁籤執行批次查詢
6. 視需要開啟地圖編輯視窗檢視已取得座標的資料
7. 匯出為 `xlsx` 或 `csv`

## Geocoding 輸出欄位

程式會在資料表中維護以下欄位：

- `candidate_address`
- `matched_address`
- `geocode_status`
- `tgos_candidate_address`
- `coord_x`
- `coord_y`
- `coord_system`
- `geocode_match_type`
- `geocode_result_count`
- `geocode_error`

其中：

- `candidate_address` 是實際送去查詢的候選地址
- `matched_address` 是 TGOS 回傳的匹配地址
- `tgos_candidate_address` 是查詢過程中使用的 TGOS 候選內容
- `coord_x`、`coord_y` 為座標
- `geocode_status` 用來表示成功、略過、無結果或錯誤等狀態
- `geocode_error` 用來保存錯誤訊息

## 本地資料與快取

此專案會使用兩種本地資料：

- 工作區狀態：儲存在 Electron `userData` 目錄中的 `workspace.json`
- TGOS 查詢快取：專案根目錄下的 `tgos-cache.sqlite`

`workspace.json` 會保存目前已載入資料集、分頁狀態與清洗腳本。  
`tgos-cache.sqlite` 會保存 TGOS 請求所需的 header / token 快取，降低重複初始化請求。

## 命令列測試腳本

專案根目錄提供 `tgos-query.js`，可直接在命令列快速測試地址查詢邏輯。

```bash
node tgos-query.js
```

執行前可先修改檔案內的 `oAddresses` 陣列。

## 專案結構

```text
.
├── build/                  # 打包圖示與資產
├── src/
│   ├── main/               # Electron 主程序與 preload
│   ├── renderer/           # 前端介面
│   └── services/           # TGOS 查詢服務
├── tgos-query.js           # 命令列查詢測試腳本
├── tgos-cache.sqlite       # TGOS 快取資料庫
└── package.json
```

## 開發說明

- 主程序入口：`src/main/main.js`
- 預載腳本：`src/main/preload.js`
- 主畫面邏輯：`src/renderer/app.js`
- TGOS 查詢服務：`src/services/tgos.js`

TGOS 查詢流程目前會先抓取必要 cookie 與 token，再送出地址查詢請求。若快取失效，程式會自動重新抓取後重試。

## 注意事項

- 本專案會對外連線至 TGOS 與 OpenStreetMap 相關服務
- 請確認你的使用方式符合 TGOS 與相關資料服務的使用條款、頻率限制與授權規範
- 批次查詢目前內建每筆約 `0.3` 秒延遲，以降低過度密集請求
- 若要將本工具用於正式環境或大量資料處理，建議先自行驗證查詢穩定性與資料正確性

## 授權

本專案採用 [MIT License](./LICENSE)。

