const fs = require('fs');
const path = require('path');

//Excel JS
const Excel = require('exceljs');
const filesJs = require('./files.js');
const extend = require("extend");

module.exports = function () {
  console.log('readExcelJS');

  const dirPath = JSON.parse(fs.readFileSync('dirPath.json', 'utf8'));
  const langList = ['zh-cn', 'zh-tw', 'en', 'th', 'vi'];

  const InspectionXlsx = fs.readdirSync(path.resolve('.')).find(file => {
    return /Inspection_/.test(file);
  });
  //檢核檔案的日期
  const xlsxDate = InspectionXlsx.replace('Inspection_', '').replace('.xlsx', '')
  console.log(InspectionXlsx, xlsxDate)

  /** 預先輸出資料夾 */
  const oupputPath = path.resolve('.', 'backup', xlsxDate, 'output');

  filesJs.delDir(oupputPath);

  Object.values(dirPath).forEach((foldstage) => {
    Object.values(langList).forEach((it) => {
      filesJs.createFolderSync(path.resolve('.', 'backup', xlsxDate, 'output', 'i18n', foldstage, it));
      filesJs.createFolderSync(path.resolve('.', 'backup', xlsxDate, 'format', 'i18n', foldstage, it));
    });
  });


  /** 讀取Inspection.xlsx */
  const workbook = new Excel.Workbook();
  workbook.xlsx.readFile(InspectionXlsx).then(function () {
    //Get sheet by Name
    const worksheet = workbook.getWorksheet('MySheet');
    const langXls = []; //寫檔以後可以跟讀取輸出的langXls.json做交互確認
    const outputJson = {}; //輸出的整理
    worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      /** 一列列讀出來 */
      if (rowNumber > 1) {
        const currRow = worksheet.getRow(rowNumber);

        const rowjson = {};
        rowjson.key = row.values[1];
        langList.forEach((key, index) => {
          rowjson[key] = currRow.getCell(index + 2).value;

          const pathList = row.values[1].split('.');

          const filePath = JSON.stringify([pathList[0], key, pathList[1] + '.json']);
          const arr = pathList.slice(2, pathList.length);
          const func = function (ar, obj) {
            const k = ar.shift();

            if (ar.length > 0) {
              obj[k] = obj[k] || {};
              func(ar, obj[k]);
            } else {
              obj[k] = EscapeCharacter(rowjson[key] || '');
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

    const cloneJson = extend(true, {}, outputJson);
    /** 讀出來的JSON結構 依序取出檔案名 */
    Object.keys(outputJson).forEach((langkey) => {
      Object.keys(outputJson[langkey]).forEach((writePath) => {
        const resolvePath = JSON.parse(writePath);

        const fileName = resolvePath[resolvePath.length - 1];
        if (fileName === 'undefined.json') {
          return;
        }



        /** 讀取輸出日期的模板  而且因為ZH_TW是基準所以用ZH_TW來做會比較完整 */
        const modulePath = ['.', 'backup', xlsxDate, 'i18n', resolvePath[0], 'zh-tw', fileName];
        fs.readFile(path.resolve(...modulePath), 'utf8', function (err, data) {
          const KeyList = [];
          /** 寫的位置 */
          const logger = fs.createWriteStream(path.resolve('.', 'backup', xlsxDate, 'output', 'i18n', ...resolvePath), {
            flags: 'a', // 'a' means appending (old data will be preserved)
          });
          const dataArray = data.split('\n');
          const spaceCondition = [];
          dataArray.forEach((line) => {
            const spaceCount = getSpaceCount(line);

            if (spaceCount !== 0 && spaceCondition.indexOf(spaceCount) < 0) {
              spaceCondition.push(spaceCount);
            }
          });

          dataArray.forEach((line, index, data) => {
            if (line === '\n' || !line.trim()) {
              return;
            }

            if (line.trim() == '{' || line.trim() == '}' || line.trim() == '},') {
              // console.log(line); //直接寫
            } else {
              const spaceCount = getSpaceCount(line);
              const writeLine = funcRegex(line);

              const findIndex = spaceCondition.indexOf(spaceCount);
              if (line.indexOf(':') != -1) {
                KeyList[findIndex] = String(line.split(':')[0].split('"').join('')).trim();
                KeyList.length = findIndex + 1;
              }
              const newValue = getDeepJson(cloneJson[langkey][writePath], 0, KeyList);

              if (writeLine === true) {

                if (typeof newValue === 'object') {
                  const regxLine = Object.keys(newValue).join('');

                  if (/^[0-9]+$/.test(regxLine) && !/[\[]/.test(line) && !/[\]]/.test(line)) {
                    const lineKey = Object.keys(newValue).find((arrKey) => {
                      if (newValue[arrKey] !== false) {
                        return true;
                      }
                    });

                    const characterArray = [' '.repeat(spaceCount), '"', ...EscapeCharacter(newValue[lineKey])];
                    characterArray.push(line.indexOf(',') < 0 ? '"' : '",');
                    line = characterArray.join('').split('\n').join('');

                    newValue[lineKey] = false;
                  }
                }

                // console.log(line); //直接寫
              } else {
                // console.log('########', KeyList);
                line = contentReplace(line, newValue);
              }
            }

            if (index < data.length - 1) {
              line = line + '\n';
            }
            logger.write(line);
          });

          logger.end();
        });

        /** 如果不管排序直接全塞 ˋ上面註解掉走這裡就好 */
        fs.writeFile(
          path.resolve(path.resolve('.', 'backup', xlsxDate, 'format', 'i18n', ...resolvePath)),
          JSON.stringify(outputJson[langkey][writePath], null, 2),
          errorHandler,
        );
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

function EscapeCharacter(value) {
  value = value.split('"').join('\\"');
  value = value.split('\\').join('\\'); //mac Excel再直接編輯時候跳脫字元會隱藏一條

  return value;
}

function funcRegex(line) {
  if (/[\"]{1,2}/.test(line.split(':')[1])) {
    return false;
  } else {
    return true; //直接寫
  }
}


function contentReplace(line, value) {
  const arr = line.split(':');
  const key = arr[0];

  const beforeVal = arr.slice(1, arr.length).join(':');
  const replaceVal = [...beforeVal]
    .slice([...beforeVal].indexOf('"') + 1, [...beforeVal].lastIndexOf('"'))
    .join('');

  return key + ':' + beforeVal.replace(replaceVal, value);
};