const fs = require('fs-extra');
const path = require('path');

exports = module.exports = function (name) {};

(function () {
  for (const [key, value] of Object.entries(fs)) {
    exports[key] = value;
  }
})();

exports.is_file = function (path) {
  try {
    const stats = fs.statSync(path);
    return stats.isFile();
  } catch (err) {
    return false;
  }
};

exports.is_dir = function (path) {
  try {
    const stats = fs.statSync(path);
    return stats.isDirectory();
  } catch (err) {
    return false;
  }
};

exports.delDirSync = function (path) {
  let files = [];
  if (fs.existsSync(path)) {
    files = fs.readdirSync(path);
    files.forEach((file, index) => {
      let curPath = path + '/' + file;
      if (fs.statSync(curPath).isDirectory()) {
        this.delDirSync(curPath); //遞迴刪除資料夾
      } else {
        fs.unlinkSync(curPath); //刪除檔案
      }
    });
    fs.rmdirSync(path);
  }
};

exports.copyFolder = function (from, to, complete) {
  fs.copy(from, to)
    .then(() => {
      if (complete) {
        complete();
      }
    })
    .catch((err) => {
      return console.error(err);
    });
};

exports.createFolderSync = function (dir) {
  fs.ensureDirSync(dir);
};

exports.createFileSync = function (path, data, opt) {
  fs.outputFileSync(path, data, opt);
};

exports.removeFile = function (path) {
  fs.remove(path)
    .then(() => {
      // console.log('success!');
    })
    .catch((err) => {
      console.error(err);
    });
};

exports.copyFile = function (from, to, complete) {
  fs.copy(from, to, complete);
};
