#!/usr/bin/env node
/**
 * bootstrap.js —— 新机器环境体检（table-sdk-compare 首次安装后跑一次）
 *
 * 只读检查，不改任何东西。退出码 0=可用，1=有阻断项。
 *
 * 检查项：Node 版本 / Python 及三个包 / Excel COM / Chrome / Node 依赖 / MCP 配置 / 技能文件完整性
 *
 * 用法：
 *   node scripts/bootstrap.js
 *   node scripts/bootstrap.js --root D:\work\my-compare-project   # 顺便查该工程的 node 依赖
 *   node scripts/bootstrap.js --json
 */

const fs = require('fs');
const os = require('os');
const path = require('path');
const { spawnSync } = require('child_process');

const args = process.argv.slice(2);
let ROOT = process.cwd();
let JSON_OUT = false;
for (let i = 0; i < args.length; i++) {
  if (args[i] === '--root') ROOT = path.resolve(args[++i] || '.');
  else if (args[i] === '--json') JSON_OUT = true;
}

const checks = [];
function add(name, status, detail, hint) {
  checks.push({ name, status, detail: detail || '', hint: hint || '' });
}

function run(cmd, argv, opts) {
  try {
    const r = spawnSync(cmd, argv, Object.assign({ encoding: 'utf8', timeout: 60000 }, opts));
    if (r.error) return { ok: false, out: String(r.error.message || r.error) };
    return { ok: r.status === 0, out: (r.stdout || '').trim(), err: (r.stderr || '').trim(), status: r.status };
  } catch (e) {
    return { ok: false, out: String((e && e.message) || e) };
  }
}

// 1. Node
(function checkNode() {
  const major = parseInt(process.version.replace('v', '').split('.')[0], 10);
  if (major >= 18) add('Node.js', 'PASS', process.version);
  else add('Node.js', 'FAIL', process.version + ' 过旧', '装 Node 18 或更高');
})();

// 2. Python + 包
(function checkPython() {
  const candidates = process.platform === 'win32' ? ['python', 'py'] : ['python3', 'python'];
  let py = null;
  for (const c of candidates) {
    const r = run(c, ['--version']);
    if (r.ok && /\d+\.\d+/.test(r.out)) {
      py = c;
      add('Python', 'PASS', r.out + ' (' + c + ')');
      break;
    }
  }
  if (!py) {
    add('Python', 'FAIL', '未找到 python', '安装 Python 3.9+ 并加入 PATH');
    return;
  }
  const libs = ['openpyxl', 'PIL', 'numpy'];
  const probe = libs.map(l => 'import ' + l).join('; ');
  const r = run(py, ['-c', probe]);
  if (r.ok) add('Python 包', 'PASS', libs.join(' / '));
  else {
    const missing = libs.filter(l => !run(py, ['-c', 'import ' + l]).ok);
    add('Python 包', 'FAIL', '缺 ' + (missing.join(', ') || '未知'), 'pip install ' + missing.map(m => (m === 'PIL' ? 'pillow' : m)).join(' '));
  }
})();

// 3. Excel COM（Windows only）
(function checkExcel() {
  if (process.platform !== 'win32') {
    add('Excel', 'WARN', '非 Windows，无法做 Excel 基准', '只能做 API 层对比，报告须注明"无 Excel 基准列"');
    return;
  }
  const ps = [
    "$ErrorActionPreference='Stop'",
    '$x = New-Object -ComObject Excel.Application',
    '$v = $x.Version',
    '$x.Quit()',
    '[void][Runtime.InteropServices.Marshal]::ReleaseComObject($x)',
    'Write-Output $v',
  ].join('; ');
  const r = run('powershell', ['-NoProfile', '-NonInteractive', '-Command', ps], { timeout: 90000 });
  if (r.ok && /\d+\.\d+/.test(r.out)) add('Excel COM', 'PASS', '版本 ' + r.out.match(/\d+\.\d+/)[0]);
  else {
    const msg = r.err || r.out || '调用失败';
    add(
      'Excel COM',
      'FAIL',
      msg.split('\n')[0].slice(0, 120),
      '需要桌面版 Excel；若弹出"正在配置 Office"先手动打开一次 Excel 完成初始化'
    );
  }
})();

// 4. Chrome（playwright channel:'chrome' 依赖本机 Chrome）
(function checkChrome() {
  const paths = [];
  if (process.platform === 'win32') {
    const pf = process.env['ProgramFiles'] || 'C:\\Program Files';
    const pf86 = process.env['ProgramFiles(x86)'] || 'C:\\Program Files (x86)';
    const lad = process.env['LOCALAPPDATA'] || path.join(os.homedir(), 'AppData', 'Local');
    paths.push(path.join(pf, 'Google', 'Chrome', 'Application', 'chrome.exe'));
    paths.push(path.join(pf86, 'Google', 'Chrome', 'Application', 'chrome.exe'));
    paths.push(path.join(lad, 'Google', 'Chrome', 'Application', 'chrome.exe'));
  } else if (process.platform === 'darwin') {
    paths.push('/Applications/Google Chrome.app/Contents/MacOS/Google Chrome');
  } else {
    paths.push('/usr/bin/google-chrome', '/usr/bin/chromium');
  }
  const hit = paths.find(p => fs.existsSync(p));
  if (hit) add('Chrome', 'PASS', hit);
  else add('Chrome', 'FAIL', '未在常见路径找到', '装 Google Chrome；或改用 playwright 自带 chromium（截图口径会变，需在报告注明）');
})();

// 5. Node 依赖（先看 ROOT，再看本脚本所在目录）
(function checkNodeDeps() {
  const roots = [ROOT, __dirname, path.join(__dirname, '..')];
  const mods = ['playwright', 'pngjs', 'esbuild'];
  const ok = {};
  for (const m of mods) {
    ok[m] = roots.some(r => {
      try {
        require.resolve(m, { paths: [r, path.join(r, 'node_modules')] });
        return true;
      } catch (e) {
        return false;
      }
    });
  }
  const miss = mods.filter(m => !ok[m]);
  if (!miss.length) add('Node 依赖', 'PASS', mods.join(' / '));
  else add('Node 依赖', 'FAIL', '缺 ' + miss.join(', '), '在工程根执行 npm i ' + miss.join(' '));
})();

// 6. MCP 配置
(function checkMcp() {
  const p = path.join(os.homedir(), '.workbuddy', 'mcp.json');
  if (!fs.existsSync(p)) {
    add('MCP', 'FAIL', '找不到 ' + p, '按 README 第 3 节配置 spreadjs-mcp（必须是自己的 token）');
    return;
  }
  let j;
  try {
    j = JSON.parse(fs.readFileSync(p, 'utf8'));
  } catch (e) {
    add('MCP', 'FAIL', 'mcp.json 不是合法 JSON', '修好语法再试');
    return;
  }
  const servers = (j && j.mcpServers) || {};
  const sjs = servers['spreadjs-mcp'];
  const hasToken = sjs && sjs.headers && sjs.headers.token && String(sjs.headers.token).length > 10;
  if (sjs && hasToken) add('MCP spreadjs-mcp', 'PASS', sjs.url || '(已配置)');
  else if (sjs) add('MCP spreadjs-mcp', 'FAIL', '已配置但缺 token', 'token 找 GrapeCity 申请，不要用别人的');
  else add('MCP spreadjs-mcp', 'FAIL', '未配置', '按 README 第 3 节加，并在连接器页点"信任"');
  add('MCP gcexcel-mcp', servers['gcexcel-mcp'] ? 'PASS' : 'WARN', servers['gcexcel-mcp'] ? '已配置' : '未配置（可选）');
})();

// 7. 技能文件完整性
(function checkSkill() {
  const base = __dirname.replace(/[\\/]scripts[\\/]?$/, '');
  const must = ['SKILL.md', 'README.md', 'VERSION', 'CHANGELOG.md', path.join('scripts', 'preflight_univer.js')];
  const miss = must.filter(f => !fs.existsSync(path.join(base, f)));
  const refs = path.join(base, 'references');
  const refCount = fs.existsSync(refs) ? fs.readdirSync(refs).length : 0;
  if (!miss.length && refCount >= 8) add('技能文件', 'PASS', 'references ' + refCount + ' 份');
  else add('技能文件', 'FAIL', '缺 ' + (miss.join(', ') || 'references'), '重新 clone 仓库，不要手动挑文件');
})();

// 输出
const fail = checks.filter(c => c.status === 'FAIL');
const warn = checks.filter(c => c.status === 'WARN');
const pass = fail.length === 0;

if (JSON_OUT) {
  console.log(JSON.stringify({ time: new Date().toISOString(), pass, checks }, null, 2));
} else {
  console.log('');
  console.log('table-sdk-compare 环境体检  ' + (pass ? 'PASS' : 'FAIL'));
  console.log('─'.repeat(64));
  for (const c of checks) {
    const mark = c.status === 'PASS' ? '✓' : c.status === 'WARN' ? '!' : '✗';
    console.log(mark + ' ' + c.name.padEnd(18, ' ') + c.detail);
    if (c.hint && c.status !== 'PASS') console.log('  → ' + c.hint);
  }
  console.log('─'.repeat(64));
  if (!pass) console.log('有 ' + fail.length + ' 项阻断，按提示修完再跑测试。');
  else if (warn.length) console.log('可用，但有 ' + warn.length + ' 项警告（见上）。');
  else console.log('全部就绪。下一步：node scripts/preflight_univer.js 查具体模块的包与插件。');
  console.log('');
}

process.exit(pass ? 0 : 1);
