// 重建 Univer 侧 bundle（vendor/univer.bundle.js + .css）
//   node build_bundle.js    （在 poc-compare 目录下）
//
// 为什么要有这个脚本：entry 里注册哪些插件、用哪些参数打包，是"换模块对比时最容易忘"
// 的一步。忘记注册 = 方法全部 undefined，看起来像产品不支持。参数固定在下面，别改动口径。
const path = require('path');

const entry = path.join(__dirname, 'univer_compare_entry.js');
const outfile = path.join(__dirname, '..', '测试工程', 'vendor', 'univer.bundle.js');

console.log('entry  :', entry);
console.log('outfile:', outfile);

require('esbuild').build({
  entryPoints: [entry],
  bundle: true,
  format: 'iife',
  target: 'chrome100',
  outfile,
  assetNames: 'assets/[name]-[hash]',
  logLevel: 'warning',
  logLimit: 10,
}).then((r) => {
  if (r.errors && r.errors.length) {
    console.error('构建有错误：', r.errors.length);
    process.exit(1);
  }
  console.log('构建完成，警告', (r.warnings || []).length, '条');
}).catch((e) => {
  console.error('构建失败', e);
  process.exit(1);
});
