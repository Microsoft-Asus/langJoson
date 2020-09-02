const fs = require('fs');
const path = require('path');

//hson2xls
const json2xls = require('json2xls');

//Excel JS
const Excel = require('exceljs');

module.exports = function () {
  console.log('readExcelJS');
  const enumIDKey = JSON.parse(fs.readFileSync('enumID2Key.json', 'utf8'));
  const dirPath = JSON.parse(fs.readFileSync('dirPath.json', 'utf8'));
  const langList = JSON.parse(fs.readFileSync('langList.json', 'utf8'));
  const i18nKeyList = Object.keys(enumIDKey).filter((key) => {
    const FileRegex = new RegExp('^[0-9]+$', 'i');
    return !FileRegex.test(key);
  });
  /** 預先輸出資料夾 */
  const oupputPath = path.resolve('.', 'output');
  const oupputDir = is_dir(oupputPath);
  if (!oupputDir) {
    fs.mkdirSync(oupputPath);
  }
  const i18nPath = path.resolve('.', 'output', 'i18n');
  const i18nDir = is_dir(i18nPath);
  if (!i18nDir) {
    fs.mkdirSync(i18nPath);
  }
  createDir(dirPath, ['.', 'output', 'i18n']);
  Object.values(dirPath).forEach((it) => {
    createDir(langList, ['.', 'output', 'i18n', it]);
  });

  //讀取langXls.xlsx
  const workbook = new Excel.Workbook();
  workbook.xlsx.readFile('Inspection.xlsx').then(
    function () {
      //Get sheet by Name
      const worksheet = workbook.getWorksheet('MySheet');
      const langXls = [];
      worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
        // console.log('Row ' + rowNumber + ' = ' + JSON.stringify(row.values));
        const currRow = worksheet.getRow(rowNumber);
        const rowjson = {};
        rowjson.key = row.values[1];
        langList.forEach((key, index) => {
          rowjson[key] = currRow.getCell(index + 2).value;
        });

        rowjson.id = currRow.getCell(worksheet.columnCount).value;
        langXls.push(rowjson);
      });
    },
    function (err) {
      console.log(err);
    },
  );
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

function is_file(path) {
  const stats = fs.statSync(path);
  return stats.isFile();
}
