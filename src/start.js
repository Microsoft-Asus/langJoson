const path = require('path');
const fs = require('fs');
//hson2xls
const json2xls = require('json2xls');

//Excel JS
const Excel = require('exceljs');
const workbook = new Excel.Workbook();

(function () {
  //資料夾名字 backstage ,frontstage
  const i18nDirPath = fs.readdirSync(path.resolve('.', 'i18n'));
  //en, zh-cn, zh-tw
  const langList = [];
  const objLang = {};

  //fileName
  const fileJson = {};
  //輸出
  const mapJson = {};

  const jsonFileRegex = new RegExp(`\/([a-z]+)\/([a-z\-]{2,})\/([a-z]+)\.json$`, 'i');
  const jsonFilesPath = walkFilesSync(path.resolve('.', 'i18n'), (fname, dirname) => {
    const fullpath = path.join(dirname, fname);
    return /\.json$/.test(fullpath);
  });
  //以資料當作Key 取出重複的資料
  const langValue = {};

  //組合不重複的語系資料夾名字,做第一階段的過濾,取出全部的語系
  i18nDirPath.forEach((dirpath, id) => {
    fs.readdirSync(path.resolve('.', 'i18n', dirpath)).forEach((pathname) => {
      if (!objLang[pathname]) {
        objLang[pathname] = true;
        langList.push(pathname);
      }

      fileJson[dirpath] = fileJson[dirpath] || {};

      fs.readdirSync(path.resolve('.', 'i18n', dirpath, pathname)).forEach((filepath) => {
        const fileString = filepath.replace('.json', '');
        if (!fileJson[dirpath][fileString]) {
          fileJson[dirpath][fileString] = true;
        }
      });
    });
  });

  // console.log('i18nDirPath:', i18nDirPath);
  // console.log('langList:', langList);
  // console.log(JSON.stringify(fileJson));
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
        fileJson[dirpath][fileString] = fileJson[dirpath][fileString] === true ? {} : fileJson[dirpath][fileString];
        fileJson[dirpath][fileString][k] = true;

        mapJson[`${dirpath}.${fileString}.${k}`] = mapJson[`${dirpath}.${fileString}.${k}`] || {};
        mapJson[`${dirpath}.${fileString}.${k}`][lang] = flatData[k];
      });
    }
  });

  const xlsjson = [];

  Object.keys(mapJson).forEach((key) => {
    langList.forEach((lang) => {
      mapJson[key][lang] = mapJson[key][lang] || '';
    });

    langValue[JSON.stringify(mapJson[key])] = langValue[JSON.stringify(mapJson[key])] || [];
    langValue[JSON.stringify(mapJson[key])].push(xlsjson.length);

    xlsjson.push({ key, ...mapJson[key], id: xlsjson.length });
  });
  //重複內容
  const repeatValue = Object.values(langValue).filter((it) => {
    return it.length > 1;
  });

  const repeatMap = repeatValue.map((it) => {
    const repeatArray = [];
    it.forEach((langIndex) => {
      repeatArray.push(xlsjson[langIndex]);
      xlsjson[langIndex] = null;
    });

    return repeatArray;
  });
  repeatMap.forEach((it) => it.forEach((langValue) => xlsjson.push(langValue)));

  const count = 0;
  const xlsJsonFilter = xlsjson.filter((it) => it !== null);
  const xls = json2xls(xlsJsonFilter);
  fs.writeFileSync('langXls.xlsx', xls, 'binary');

  fs.writeFile('fileJson.json', JSON.stringify(fileJson), function (err) {});
  fs.writeFile('mapJson.json', JSON.stringify(mapJson), function (err) {});
  fs.writeFile('langXls.json', JSON.stringify(xlsJsonFilter), function (err) {});
  fs.writeFile('repeatMap.json', JSON.stringify(repeatMap), function (err) {});

  //Excels
  const worksheet = workbook.addWorksheet('MySheet');
  const excelColumn = Object.keys(xlsJsonFilter[0]).map((it) => {
    return { header: it, key: it };
  });
  worksheet.columns = excelColumn;
  worksheet.addRows(xlsJsonFilter);
  repeatValue.forEach((repeat) => {
    const rowsIndex = xlsJsonFilter.findIndex((it) => {
      return it.id === repeat[0];
    });

    const letter = String('bcdefghijklmnopqrstuvwxyz').toUpperCase();

    [...letter].slice(0, langList.length).forEach((key) => {
      console.log(
        rowsIndex,
        '/',
        repeat.length,
        '///',
        `${key}${rowsIndex + 2}:${key}${rowsIndex + repeat.length - 1 + 2}`,
      );
      worksheet.mergeCells(`${key}${rowsIndex + 2}:${key}${rowsIndex + repeat.length - 1 + 2}`);
    });
  });

  // worksheet.mergeCells(`B1161:B1162`);
  // worksheet.mergeCells(`C14:C15`);
  (async function () {
    return await workbook.xlsx.writeFile('Excel.xlsx').then(
      async () => {
        // console.log(this);
      },
      function (err) {
        console.log(err);
      },
    );
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
