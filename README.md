# root

npm install

# src 底下

node start.js

# Structure

start.js 主程式

├── 產生 json 用於比對 檢查輸出資料

└── 產生 inspection.xlsx

readExcel.js 拿來讀取 Excel 輸出 i18n 的 -->可以另外複製出去獨立專案使用

├── 讀取 start.js 產生的 dirPath.json /langList.json 拿來當作設定檔

└── 產生 output folder

## 參數

```js
/** 輸出開關
 *  true => 輸出Excel
 *  false => 讀取Excel 輸出 i18n
 */
EXPORT_EXCEL = true;
```

## 輸出的 Excel 檔案

src/Inspection.xlsx

## 解析產生的 i18n 檔案 src/output
