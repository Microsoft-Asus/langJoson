const path = require('path');
const fs = require('fs');
//hson2xls
const json2xls = require('json2xls');

//Excel JS
const Excel = require('exceljs');
const workbook = new Excel.Workbook();

const readExcel = require('./readExcel.js');

(function () {
  //資料夾名字 backstage ,frontstage
  const i18nDirPath = fs.readdirSync(path.resolve('.', 'i18n'));
  //en, zh-cn, zh-tw
  const langList = [];

  //輸出
  const mapJson = {};

  const jsonFileRegex = new RegExp(`\/([a-z]+)\/([a-z\-]{2,})\/([a-z]+)\.json$`, 'i');
  const jsonFilesPath = walkFilesSync(path.resolve('.', 'i18n'), (fname, dirname) => {
    const fullpath = path.join(dirname, fname);
    return /\.json$/.test(fullpath);
  });
  //以資料當作Key 取出重複的資料
  const repeatAll = {};
  const repeatZhTw = {};
  //組合不重複的語系資料夾名字,做第一階段的過濾,取出全部的語系結構
  i18nDirPath.forEach((dirpath, id) => {
    fs.readdirSync(path.resolve('.', 'i18n', dirpath)).forEach((pathname) => {
      if (langList.indexOf(pathname) < 0 && is_dir(path.resolve('.', 'i18n', dirpath, pathname))) {
        langList.push(pathname);
      }
    });
  });

  // console.log('i18nDirPath:', i18nDirPath);
  // console.log('langList:', langList);
  // console.log(jsonFilesPath[0]);

  jsonFilesPath.forEach((jfPath) => {
    const match = jfPath.split('i18n')[1].match(jsonFileRegex);

    if (match) {
      const dirpath = match[1];
      const lang = match[2];
      const fileString = match[3];

      const rawdata = fs.readFileSync(jfPath, 'utf8');
      const data = JSON.parse(rawdata.toString());
      const flatData = flattenObject(data);

      Object.keys(flatData).forEach((k) => {
        mapJson[`${dirpath}.${fileString}.${k}`] = mapJson[`${dirpath}.${fileString}.${k}`] || {};
        mapJson[`${dirpath}.${fileString}.${k}`][lang] = flatData[k];
      });
    }
  });
  //xls json組合用
  const xlsjson = [];
  //id : key 的map表
  const enumID2Key = {};

  //追加紀錄最大字長作為調整Excel欄位寬度
  const maxWordLength = { key: 0, id: 5 };

  Object.keys(mapJson).forEach((key) => {
    langList.forEach((lang) => {
      mapJson[key][lang] = mapJson[key][lang] || '';
      //追加紀錄最大字長作為調整Excel欄位寬度
      maxWordLength[lang] = maxWordLength[lang] || 0;
      maxWordLength[lang] =
        mapJson[key][lang].length > maxWordLength[lang] ? mapJson[key][lang].length : maxWordLength[lang];
    });
    //id to key map表
    enumID2Key[xlsjson.length] = key;
    enumID2Key[key] = xlsjson.length;
    //追加紀錄最大字長作為調整Excel欄位寬度
    maxWordLength.key = key.length > maxWordLength.key ? key.length : maxWordLength.key;

    //過濾重複ZH-TW 紀錄id
    repeatZhTw[JSON.stringify(mapJson[key]['zh-tw'])] = repeatZhTw[JSON.stringify(mapJson[key]['zh-tw'])] || [];
    repeatZhTw[JSON.stringify(mapJson[key]['zh-tw'])].push(xlsjson.length);

    //過濾全部重語系內容 紀錄id
    repeatAll[JSON.stringify(mapJson[key])] = repeatAll[JSON.stringify(mapJson[key])] || [];
    repeatAll[JSON.stringify(mapJson[key])].push(xlsjson.length);

    xlsjson.push({ key, ...mapJson[key], id: xlsjson.length });
  });

  //過濾出全部的重複內容
  const repeatValue = Object.values(repeatAll).filter((it) => {
    return it.length > 1;
  });
  //針對ZH-TW的過濾
  const repeatZhTwValue = [];
  for (const [key, value] of Object.entries(repeatZhTw)) {
    if (value.length > 1) {
      repeatZhTwValue.push({
        key: key,
        repeatid: value,
        repeatkey: value.map((id) => {
          return enumID2Key[id];
        }),
      });
    }
  }

  //重複內容的Map,將重複的內容抓出去 原本位置先清空
  const repeatMap = repeatValue.map((it) => {
    const repeatArray = [];
    it.forEach((langIndex) => {
      repeatArray.push(xlsjson[langIndex]);
      xlsjson[langIndex] = null;
    });

    return repeatArray;
  });
  //將重複內容塞到陣列最後面
  repeatMap.forEach((it) => it.forEach((langValue) => xlsjson.push(langValue)));

  const count = 0;
  //空出來的位置過濾掉
  const xlsJsonFilter = xlsjson.filter((it) => it !== null);
  const xls = json2xls(xlsJsonFilter);
  //原始XLSX
  fs.writeFileSync('langXls.xlsx', xls, 'binary');
  //檢查輸出的JSON是不是自己要的
  fs.writeFile('dirPath.json', JSON.stringify(i18nDirPath, null, 4), errorHandler);
  fs.writeFile('langList.json', JSON.stringify(langList, null, 4), errorHandler);
  fs.writeFile('mapJson.json', JSON.stringify(mapJson, null, 4), errorHandler);
  fs.writeFile('langXls.json', JSON.stringify(xlsJsonFilter, null, 4), errorHandler);
  fs.writeFile('repeatMap.json', JSON.stringify(repeatMap, null, 4), errorHandler);
  fs.writeFile('repeatZhTw.json', JSON.stringify(repeatZhTwValue, null, 4), errorHandler);
  fs.writeFile('enumID2Key.json', JSON.stringify(enumID2Key, null, 4), errorHandler);

  //產出有合併欄位的 Excels
  const worksheet = workbook.addWorksheet('MySheet');
  const excelColumn = Object.keys(xlsJsonFilter[0]).map((it) => {
    return { header: it, key: it, width: maxWordLength[it] };
  });

  worksheet.columns = excelColumn;
  worksheet.addRows(xlsJsonFilter);
  repeatValue.forEach((repeat) => {
    const rowsIndex = xlsJsonFilter.findIndex((it) => {
      return it.id === repeat[0];
    });

    const letter = String('bcdefghijklmnopqrstuvwxyz').toUpperCase();

    [...letter].slice(0, langList.length).forEach((key) => {
      worksheet.mergeCells(`${key}${rowsIndex + 2}:${key}${rowsIndex + repeat.length - 1 + 2}`);
    });
  });

  (async function () {
    return await workbook.xlsx.writeFile('Inspection.xlsx').then(async () => {
      // console.log(this);
      /** 讀取檢查 */
      readExcel();
    }, errorHandler);
  })();
})();

function walkFilesSync(dirname, filter = undefined) {
  try {
    let files = [];

    fs.readdirSync(dirname).forEach((fname) => {
      const fpath = path.join(dirname, fname);

      if (is_file(fpath)) {
        if ((filter && filter(fname, dirname)) || true) {
          files.push(fpath);
        }
      } else if (is_dir(fpath)) {
        files = files.concat(walkFilesSync(fpath, filter));
      }
    });

    return files;
  } catch (err) {
    throw err;
  }
}

function is_dir(path) {
  const stats = fs.statSync(path);
  return stats.isDirectory();
}

function is_file(path) {
  const stats = fs.statSync(path);
  return stats.isFile();
}

function flattenObject(ob) {
  var toReturn = {};

  for (var i in ob) {
    if (!ob.hasOwnProperty(i)) continue;

    if (typeof ob[i] == 'object' && ob[i] !== null) {
      var flatObject = flattenObject(ob[i]);
      for (var x in flatObject) {
        if (!flatObject.hasOwnProperty(x)) continue;

        toReturn[i + '.' + x] = flatObject[x];
      }
    } else {
      toReturn[i] = ob[i];
    }
  }
  return toReturn;
}

function errorHandler(err) {
  if (err) {
    console.log(err);
    throw err;
  }
}
