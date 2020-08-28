const path = require('path');
const fs = require('fs');

{
  //資料夾名字 backstage ,frontstage
  const i18nDirPath = fs.readdirSync(path.resolve('.', 'i18n'));
  //en, zh-cn, zh-tw
  const langList = [];
  const objLang = {};

  //fileName
  const fileJson = {};
  //輸出
  const langs = {};

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

  console.log('i18nDirPath:', i18nDirPath);
  console.log('langList:', langList);
  // console.log(JSON.stringify(fileJson));

  const jsonFileRegex = new RegExp(`\/([a-z]+)\/([a-z\-]{2,})\/([a-z]+)\.json$`, 'i');
  const jsonFilesPath = walkFilesSync(path.resolve('.', 'i18n'), (fname, dirname) => {
    const fullpath = path.join(dirname, fname);

    return /\.json$/.test(fullpath);
  });

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

        langs[`${dirpath}.${fileString}.${k}`] = langs[`${dirpath}.${fileString}.${k}`] || {};
        langs[`${dirpath}.${fileString}.${k}`][lang] = flatData[k];
      });
    }
  });

  fs.writeFile('fileJson.json', JSON.stringify(fileJson), function (err) {});
  fs.writeFile('langs.json', JSON.stringify(langs), function (err) {});
}

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
