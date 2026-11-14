// 打包 Univer Pro 全功能 bundle
//   node tools/build_pro_bundle.js
// 前置：../poc-compare/node_modules 已装 @univerjs/preset-sheets-advanced 等 Pro 包
const path = require('path');
const fs = require('fs');

const poc = path.join(__dirname, '..', '..', 'poc-compare');
const entry = path.join(poc, 'univer_pro_entry.js');
const outfile = path.join(__dirname, '..', 'vendor', 'univer-pro.bundle.js');

if (!fs.existsSync(entry)) {
  console.error('找不到入口：', entry);
  process.exit(1);
}
console.log('entry  :', entry);
console.log('outfile:', outfile);

require(path.join(poc, 'node_modules', 'esbuild')).build({
  entryPoints: [entry],
  bundle: true,
  format: 'iife',
  target: 'chrome100',
  outfile,
  absWorkingDir: poc,
  assetNames: 'assets/[name]-[hash]',
  logLevel: 'warning',
  logLimit: 20,
  minify: true,
}).then((r) => {
  if (r.errors && r.errors.length) {
    console.error('构建有错误：', r.errors.length);
    process.exit(1);
  }
  const sz = fs.statSync(outfile).size / 1048576;
  console.log('构建完成，' + sz.toFixed(1) + ' MB，警告 ' + (r.warnings || []).length + ' 条');
}).catch((e) => {
  console.error('构建失败', e);
  process.exit(1);
});
