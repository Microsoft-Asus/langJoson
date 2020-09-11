const fs = require('fs');
const path = require('path');
exports = module.exports = function (name) { };

exports.is_file = function (path) {
  const stats = fs.statSync(path);
  return stats.isFile();
};

exports.is_dir = function (path) {
  try {
    const stats = fs.statSync(path);
    return stats.isDirectory();
  } catch (err) {
    return false;
  }
};

exports.delDir = function (path) {
  let files = [];
  if (fs.existsSync(path)) {
    files = fs.readdirSync(path);
    files.forEach((file, index) => {
      let curPath = path + '/' + file;
      if (fs.statSync(curPath).isDirectory()) {
        this.delDir(curPath); //遞迴刪除資料夾
      } else {
        fs.unlinkSync(curPath); //刪除檔案
      }
    });
    fs.rmdirSync(path);
  }
};


exports.copyFolderSync = function (from, to) {
  fs.mkdirSync(to);
  fs.readdirSync(from).forEach(element => {
    if (fs.lstatSync(path.join(from, element)).isFile()) {
      fs.copyFileSync(path.join(from, element), path.join(to, element));
    } else {
      this.copyFolderSync(path.join(from, element), path.join(to, element));
    }
  });
}