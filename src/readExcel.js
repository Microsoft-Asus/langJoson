const fs = require('fs');
const path = require('path');

//Excel JS
const Excel = require('exceljs');

module.exports = function () {
  console.log('readExcelJS');

  const dirPath = JSON.parse(fs.readFileSync('dirPath.json', 'utf8'));
  const langList = JSON.parse(fs.readFileSync('langList.json', 'utf8'));

  /** 預先輸出資料夾 */
  const oupputPath = path.resolve('.', 'output');
  if (!is_dir(oupputPath)) {
    fs.mkdirSync(oupputPath);
  }
  const i18nPath = path.resolve('.', 'output', 'i18n');
  if (!is_dir(i18nPath)) {
    fs.mkdirSync(i18nPath);
  }
  createDir(dirPath, ['.', 'output', 'i18n']);
  Object.values(dirPath).forEach((it) => {
    createDir(langList, ['.', 'output', 'i18n', it]);
  });

  //讀取Inspection.xlsx
  const workbook = new Excel.Workbook();
  workbook.xlsx.readFile('Inspection.xlsx').then(function () {
    //Get sheet by Name
    const worksheet = workbook.getWorksheet('MySheet');
    const langXls = []; //寫檔以後可以跟讀取輸出的langXls.json做交互確認
    const outputJson = {}; //輸出的整理
    worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      // console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
      if (rowNumber > 1) {
        const currRow = worksheet.getRow(rowNumber);

        const rowjson = {};
        rowjson.key = row.values[1];
        langList.forEach((key, index) => {
          rowjson[key] = currRow.getCell(index + 2).value;

          const pathList = row.values[1].split('.');

          const filePath = JSON.stringify(['.', 'output', 'i18n', pathList[0], key, pathList[1] + '.json']);
          const arr = pathList.slice(2, pathList.length);
          const func = function (ar, obj) {
            const k = ar.shift();

            if (ar.length > 0) {
              obj[k] = obj[k] || {};
              func(ar, obj[k]);
            } else {
              obj[k] = (rowjson[key] || '').split('\\n').join('\n');
            }
          };
          outputJson[key] = outputJson[key] || {};
          outputJson[key][filePath] = outputJson[key][filePath] || {};
          func(arr, outputJson[key][filePath]);
        });

        rowjson.rowid = currRow.getCell(worksheet.columnCount).value;
        langXls.push(rowjson);
      }
    });

    // console.log(outputJson['zh-tw']['[".","output","i18n","frontstage","zh-tw","agent.json"]']);
    Object.keys(outputJson).forEach((langkey) => {
      Object.keys(outputJson[langkey]).forEach((writePath) => {
        const resolvePath = JSON.parse(writePath);
        const fileName = resolvePath[resolvePath.length - 1];
        if (fileName === 'undefined.json') {
          return;
        }
        /** 重要!!! 團隊用兩格縮排 所以這裡寫回去要用兩格 不然git會很混亂 */
        fs.writeFile(
          path.resolve(...resolvePath),
          JSON.stringify(outputJson[langkey][writePath], null, 2),
          errorHandler,
        );
      });
    });
  }, errorHandler);
};

function createDir(dirsetting, patharray) {
  Object.values(dirsetting).forEach((it) => {
    const dirpath = path.resolve(...patharray, it);
    const dir = is_dir(dirpath);
    if (!dir) {
      fs.mkdirSync(dirpath);
    }
  });
}

function is_dir(path) {
  try {
    const stats = fs.statSync(path);
    return stats.isDirectory();
  } catch (err) {
    return false;
  }
}

function errorHandler(err) {
  if (err) {
    console.log(err);
    throw err;
  }
}
