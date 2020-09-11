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
EXPORT_EXCEL = true
```

## 輸出的 Excel 檔案

src/Inspection\_{日期}.xlsx

key:程式看的

en, zh-cn, zh-tw :語系

rowid: 可以跟 外部人員 核對用

## 解析產生的 i18n 檔案

src/backup/日期/output
src/backup/日期/format

## 實作

新增語系時候要先在 src/i18n/底下複製 en 資料夾 更名成 新的語系資料夾

跑一次 輸出 excel 讓 json 設定檔生成

不然不知道要使用的新的語系代碼是什麼

## 同步更新

### 1.輸出給翻譯人員目前在使用的 Excel

目前僅作手動複製專案的 i18n 資料夾到 這個專案資料夾底下

### 2.匯入翻譯人員給的 Excel

將 Inspection\_{日期}.xlsx 放入 src 下

XPORT_EXCEL 設定 false

在 src 下跑 node start.js
