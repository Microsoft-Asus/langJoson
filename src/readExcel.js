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

  delDir(oupputPath);

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

    // console.log(outputJson['zh-tw']['[".","output","i18n","frontstage","zh-tw","agent.json"]']);
    Object.keys(outputJson).forEach((langkey) => {
      Object.keys(outputJson[langkey]).forEach((writePath) => {
        const resolvePath = JSON.parse(writePath);
        const fileName = resolvePath[resolvePath.length - 1];
        if (fileName === 'undefined.json') {
          return;
        }

        const func = function (line) {
          if (/[\"]{1,2}/.test(line.split(':')[1])) {
            return false;
          } else {
            return true; //直接寫
          }
        };

        const funcReplace = function (line, value) {
          const arr = line.split(':');
          const key = arr[0];

          const beforeVal = arr.slice(1, arr.length).join(':');
          const replaceVal = [...beforeVal]
            .slice([...beforeVal].indexOf('"') + 1, [...beforeVal].lastIndexOf('"'))
            .join('');

          return key + ':' + beforeVal.replace(replaceVal, value);
        };
        //因為ZH_TW是基準所以用ZH_TW來做會比較完整
        const modulePath = ['i18n', resolvePath[3], 'zh-tw', fileName];
        fs.readFile(path.resolve(...modulePath), 'utf8', function (err, data) {
          const KeyList = [];

          const logger = fs.createWriteStream(path.resolve(...resolvePath), {
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

          dataArray.forEach((line) => {
            if (line === '\n' || line === ' ' || !line || line === '}') {
              return;
            }

            if (/^[\{]{1}$/.test(line) || /[\}]{1}[\,]{0,1}$/.test(line)) {
              // console.log(line); //直接寫
            } else {
              const spaceCount = getSpaceCount(line);
              const writeLine = func(line);
              const findIndex = spaceCondition.indexOf(spaceCount);
              KeyList[findIndex] = String(line.split(':')[0].split('"').join('')).trim();
              KeyList.length = findIndex + 1;
              // console.log(KeyList, findIndex, '///', line);

              if (writeLine === true) {
                // console.log(line);//直接寫
              } else {
                // console.log('########', KeyList);
                line = funcReplace(line, getDeepJson(outputJson[langkey][writePath], 0, KeyList));
              }
            }

            logger.write(line + '\n');
          });

          logger.write('}');
          logger.end();
        });

        // fs.writeFile(
        //   path.resolve(...resolvePath),
        //   JSON.stringify(outputJson[langkey][writePath], null, 2),
        //   errorHandler,
        // );
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

function delDir(path) {
  let files = [];
  if (fs.existsSync(path)) {
    files = fs.readdirSync(path);
    files.forEach((file, index) => {
      let curPath = path + '/' + file;
      if (fs.statSync(curPath).isDirectory()) {
        delDir(curPath); //遞迴刪除資料夾
      } else {
        fs.unlinkSync(curPath); //刪除檔案
      }
    });
    fs.rmdirSync(path);
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
