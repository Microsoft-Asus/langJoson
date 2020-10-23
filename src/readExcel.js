const path = require('path');

//Excel JS
const Excel = require('exceljs');
const filesJs = require('./files.js');
const extend = require('extend');

module.exports = function () {
  console.log('readExcelJS');

  const dirPath = JSON.parse(filesJs.readFileSync('dirPath.json', 'utf8'));
  const columnKeyList = JSON.parse(filesJs.readFileSync('columnKeyList.json', 'utf8'));

  const langList = Object.keys(columnKeyList).filter((key) => {
    if (key !== 'key' && key !== 'rowid') {
      return true;
    }
  });
  console.log(langList);

  const InspectionXlsx = filesJs.readdirSync(path.resolve('.')).find((file) => {
    return /Inspection_/.test(file);
  });
  //檢核檔案的日期
  const xlsxDate = InspectionXlsx.replace('Inspection_', '').replace('.xlsx', '');
  console.log(InspectionXlsx, xlsxDate);

  /** 預先輸出資料夾 */
  const oupputPath = path.resolve('.', 'backup', xlsxDate, 'output');

  filesJs.delDirSync(oupputPath);

  Object.values(dirPath).forEach((foldstage) => {
    Object.values(langList).forEach((it) => {
      filesJs.createFolderSync(path.resolve('.', 'backup', xlsxDate, 'output', 'i18n', foldstage, it));
      filesJs.createFolderSync(path.resolve('.', 'backup', xlsxDate, 'format', 'i18n', foldstage, it));
    });
  });

  //樣板
  langList.forEach((key, index) => {
    // const langs_filelist = Object.values(filesJs.readdirSync(path.resolve('.', 'langs', key)));
    // console.log(langs_filelist);
    /** 從樣板抓回來 **/
    // const langsJson = JSON.parse(
    //   filesJs.readFileSync(path.resolve('.', 'langs', key, filename), 'utf8'),
    // );

  });

  /** 讀取Inspection.xlsx */
  const workbook = new Excel.Workbook();
  workbook.xlsx.readFile(InspectionXlsx).then(function () {
    //Get sheet by Name
    const worksheet = workbook.getWorksheet('MySheet');
    const langXls = []; //寫檔以後可以跟讀取輸出的langXls.json做交互確認
    const newExcelJson = {}; //輸出的整理
    worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      /** 一列列讀出來 */
      if (rowNumber > 1) {
        const currRow = worksheet.getRow(rowNumber);

        const rowjson = {};
        rowjson.key = row.values[1];
        langList.forEach((key, index) => {
          rowjson[key] = clearFormat(currRow.getCell(index + 2).value);

          const pathList = row.values[1].split('.');

          const filePath = JSON.stringify([pathList[0], key, pathList[1] + '.json']);
          const arr = pathList.slice(2, pathList.length);
          const func = function (ar, obj) {
            const k = ar.shift();

            if (ar.length > 0) {
              obj[k] = obj[k] || {};
              func(ar, obj[k]);
            } else {
              obj[k] = escapeCharacter(rowjson[key] || '');
            }
          };
          newExcelJson[key] = newExcelJson[key] || {};
          newExcelJson[key][filePath] = newExcelJson[key][filePath] || {};
          func(arr, newExcelJson[key][filePath]);
        });

        rowjson.rowid = clearFormat(currRow.getCell(worksheet.columnCount).value);
        langXls.push(rowjson);
      }
    });

    const cloneJson = extend(true, {}, newExcelJson);
    /** 讀出來的JSON結構 依序取出檔案名 */
    Object.keys(newExcelJson).forEach((langkey) => {
      console.log('####',langkey)
      Object.keys(newExcelJson[langkey]).forEach((writePath) => {
        const resolvePath = JSON.parse(writePath);

        const fileName = resolvePath[resolvePath.length - 1];
        if (fileName === 'undefined.json') {
          return;
        }

        /**
         *
         * 如果不管排序直接全塞 上面註解掉走這裡就好 讀取目前最新的i18n檔案
         * 和回來的Excel輸出的JSON做 extend
         */
        try {
          var newi18nFileData = {};
          //[ 'backstage', 'vi', 'memberAccount.json' ]
          const langsetting = resolvePath[1];//語系
          const newi18nFilePath = ['.', 'i18n', ...resolvePath];
          if (filesJs.is_file(path.resolve(...newi18nFilePath))) {
            const newi18nFileContent = filesJs.readFileSync(path.resolve(...newi18nFilePath), 'utf8');
            newi18nFileData = JSON.parse(newi18nFileContent.toString());
          }

          //因為輸出Excel時需要轉換轉譯字元不然會消失,回來時就要反轉回來
          ConvertEscapeCharacters(newExcelJson[langkey][writePath]);

          //從樣板拉
          if (filesJs.is_file(path.resolve('.','langs',langsetting,fileName))) {
             /** 從樣板抓回來 **/
            const langsJson = JSON.parse(
              filesJs.readFileSync(path.resolve('.', 'langs', langsetting, fileName), 'utf8'),
            );

            //這邊要一個刪除的mapping
            mapping(langsJson,newExcelJson[langkey][writePath]);
            //
          }


          const i18nMergeJson = extend(true, {}, newi18nFileData, newExcelJson[langkey][writePath]);

          // /**  extend合併之後輸出的檔案可以藉由git做差異分析 */
          filesJs.createFileSync(
            path.resolve(path.resolve('.', 'backup', xlsxDate, 'format', 'i18n', ...resolvePath)),
            JSON.stringify(i18nMergeJson, null, 2),
            'utf8',
          );
        } catch (err) {
          throw err;
        }
      });
    });
  }, errorHandler);
};

function errorHandler(err) {
  if (err) {
    console.log(err);
    throw err;
  }
}

//取得空格數
function getSpaceCount(line) {
  const ar = [...line];
  for (var i = 0; i < ar.length; i++) {
    if (ar[i] !== ' ') {
      break;
    }
  }
  return i;
}

function getDeepJson(obj, ind, arr) {
  if (ind < arr.length - 1) {
    return getDeepJson(obj[arr[ind]], ind + 1, arr);
  }

  return obj[arr[ind]];
}

function escapeCharacter(value) {
  value = value.split('"').join('\\"');
  value = value.split('\\').join('\\'); //mac Excel再直接編輯時候跳脫字元會隱藏一條

  return value;
}

//返回JSON檔時要轉換回轉譯字元
function ConvertEscapeCharacters(obj) {
  if (typeof obj === 'object') {
    for (const [key, value] of Object.entries(obj)) {
      if (typeof value === 'object') {
        ConvertEscapeCharacters(value);
      } else if (value && typeof value === 'string') {
        obj[key] = value
          .split('\\n')
          .join('\n')
          .split('\\b')
          .join('\b')
          .split('\\t')
          .join('\t')
          .split('\\r')
          .join('\r')
          .split('\\"')
          .join('"');
      }
    }
  } else {
    console.log('ConvertEscapeCharacters ERROR-->', typeof obj);
  }
}

//內文取代成新的
function contentReplace(line, value) {
  const arr = line.split(':');
  const key = arr[0];

  const beforeVal = arr.slice(1, arr.length).join(':');
  const replaceVal = [...beforeVal].slice([...beforeVal].indexOf('"') + 1, [...beforeVal].lastIndexOf('"')).join('');

  return key + ':' + beforeVal.replace(replaceVal, value);
}

function clearFormat(params) {
  if (typeof params === 'object' && params && params.richText) {
    const ar = Object.values(params.richText).map((item) => {
      if (item.text) {
        return item.text;
      }
    });
    return ar.join('');
  }
  return params;
}

function mapping(langs, outputjson) {
  if (typeof langs === typeof outputjson && typeof langs === 'object') {
    Object.keys(langs).forEach((k) => {

      if (typeof langs[k] === 'object') {

        mapping( langs[k], outputjson[k]);
      } else if (langs[k] === outputjson[k] && langs[k].indexOf('@:') == -1 && outputjson[k].indexOf('@:') == -1) {

        delete outputjson[k];//前台前端清除

      }
    });
  }
}