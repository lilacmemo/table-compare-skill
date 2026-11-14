#!/usr/bin/env node
/**
 * preflight_univer.js —— Univer 侧跑批前置门禁（SKILL.md 第 0 步）
 *
 * 三类误判（背景图 / 焦点高亮 / 查找替换 / 行分组 / 批注 / 图表）全部来自
 * "包没装或插件没注册 → 方法 undefined → 被写成对方缺陷"。本脚本把这件事
 * 从"事后回溯"前移成"事前断言"：不过门禁不许跑批，也不许把结果记成 FAIL。
 *
 * 用法：
 *   node preflight_univer.js                                  # 只查环境
 *   node preflight_univer.js --module comment,filter           # 查指定能力所需的包
 *   node preflight_univer.js --facade "univerAPI.getActiveWorkbook().getActiveSheet().getComments"
 *   node preflight_univer.js --url http://127.0.0.1:8124/对比工程/index.html   # 加查插件自证
 *   node preflight_univer.js --json                            # 机器可读，给 CI/跑批链用
 *
 * 退出码：0 全绿 / 1 有硬失败 / 2 用法错误
 */

const fs = require('fs');
const path = require('path');

// ---- 依赖清单：只收录本项目实测装过的真实包名，不凭记忆补 ----
// required = 缺了必失败（功能包 + 对应 UI 包 + facade 同源包）
// optional = 装了更好，缺了只提示
const MODULE_DEPS = {
  core: {
    required: ['@univerjs/core', '@univerjs/sheets', '@univerjs/sheets-ui', '@univerjs/ui'],
    optional: [],
  },
  comment: {
    required: [
      '@univerjs/thread-comment',
      '@univerjs/thread-comment-ui',
      '@univerjs/sheets-thread-comment',
      '@univerjs/sheets-thread-comment-ui',
    ],
    optional: [],
    note: '批注 = 4 个包 + 2 个 facade；漏装表现是 sheet.getComments 为 undefined',
  },
  note: {
    required: ['@univerjs/sheets-note', '@univerjs/sheets-note-ui'],
    optional: [],
  },
  filter: {
    required: ['@univerjs/sheets-filter', '@univerjs/sheets-filter-ui'],
    optional: [],
  },
  'data-validation': {
    required: ['@univerjs/sheets-data-validation', '@univerjs/sheets-data-validation-ui'],
    optional: [],
  },
  table: {
    required: ['@univerjs/sheets-table', '@univerjs/sheets-table-ui'],
    optional: [],
  },
  'conditional-formatting': {
    required: [
      '@univerjs/sheets-conditional-formatting',
      '@univerjs/sheets-conditional-formatting-ui',
    ],
    optional: [],
    note: '漏 facade 时 newConditionalFormattingRule 为 undefined（格式模块 95→63 的事故源）',
  },
  'find-replace': {
    required: ['@univerjs/find-replace', '@univerjs/sheets-find-replace'],
    optional: [],
    note: '两个包都要：一个是引擎、一个是表格侧命令；只装一个 createTextFinderAsync 为 undefined',
  },
  crosshair: {
    required: ['@univerjs/sheets-crosshair-highlight'],
    optional: [],
    note: '不在 Core Preset 里，必须单独注册',
  },
  drawing: {
    required: ['@univerjs/sheets-drawing'],
    optional: ['@univerjs/sheets-drawing-ui'],
  },
  hyperlink: {
    required: ['@univerjs/sheets-hyper-link'],
    optional: ['@univerjs/sheets-hyper-link-ui'],
  },
  sort: {
    required: ['@univerjs/sheets-sort'],
    optional: ['@univerjs/sheets-sort-ui'],
  },
  outline: {
    required: ['@univerjs-pro/sheets-outline'],
    optional: [],
    note: 'Pro 专属，OSS bundle 里没有',
  },
};

// ---- 参数解析 ----
function parseArgs(argv) {
  const out = { modules: [], facade: [], url: '', root: process.cwd(), json: false };
  for (let i = 0; i < argv.length; i++) {
    const a = argv[i];
    if (a === '--module' || a === '--modules') {
      out.modules = String(argv[++i] || '').split(',').map(s => s.trim()).filter(Boolean);
    } else if (a === '--facade') {
      out.facade = String(argv[++i] || '').split(',').map(s => s.trim()).filter(Boolean);
    } else if (a === '--url') {
      out.url = String(argv[++i] || '').trim();
    } else if (a === '--root') {
      out.root = path.resolve(String(argv[++i] || '.'));
    } else if (a === '--json') {
      out.json = true;
    } else if (a === '--help' || a === '-h') {
      console.log(
        '用法: node preflight_univer.js [--root <dir>] [--module a,b] [--facade "expr1,expr2"] [--url <page>] [--json]'
      );
      process.exit(0);
    } else {
      console.error('未知参数: ' + a);
      process.exit(2);
    }
  }
  return out;
}

// ---- 工具 ----
function readPkg(root) {
  const p = path.join(root, 'package.json');
  if (!fs.existsSync(p)) return null;
  try {
    return JSON.parse(fs.readFileSync(p, 'utf8'));
  } catch (e) {
    return null;
  }
}

function declaredNames(pkg) {
  const d = Object.assign({}, pkg.dependencies, pkg.devDependencies, pkg.optionalDependencies);
  return new Set(Object.keys(d));
}

function installedVersion(root, name) {
  const p = path.join(root, 'node_modules', name, 'package.json');
  if (!fs.existsSync(p)) return null;
  try {
    return JSON.parse(fs.readFileSync(p, 'utf8')).version || '0.0.0';
  } catch (e) {
    return null;
  }
}

// 从 package.json 里声明的包出发，遍历 node_modules 内的依赖闭包。
// 目的：区分"合法间接依赖"（父包声明了它，npm i 会重装）与"孤立安装"
// （没人声明，npm i 会当垃圾清掉）——只有后者才是坑。
function depClosure(root, roots) {
  const seen = new Set();
  const queue = Array.from(roots);
  let guard = 0;
  while (queue.length && guard++ < 20000) {
    const name = queue.shift();
    if (seen.has(name)) continue;
    seen.add(name);
    const pj = path.join(root, 'node_modules', name, 'package.json');
    if (!fs.existsSync(pj)) continue;
    let j;
    try {
      j = JSON.parse(fs.readFileSync(pj, 'utf8'));
    } catch (e) {
      continue;
    }
    for (const kind of ['dependencies', 'peerDependencies', 'optionalDependencies']) {
      for (const d of Object.keys(j[kind] || {})) if (!seen.has(d)) queue.push(d);
    }
  }
  return seen;
}

// ---- 检查 1：依赖 ----
function checkDeps(root, modules) {
  const res = { name: 'deps', pass: true, errors: [], warnings: [], details: [] };
  const pkg = readPkg(root);
  if (!pkg) {
    res.pass = false;
    res.errors.push('找不到 ' + path.join(root, 'package.json') + '，先用 --root 指定工程根目录');
    return res;
  }
  const declared = declaredNames(pkg);
  // 依赖闭包只算一次，供"孤立安装"判定用
  const reachable = depClosure(root, declared);

  const wanted = new Map(); // pkg -> {required:bool, modules:[]}
  const mods = modules.length ? modules : [];
  if (!mods.length) {
    res.warnings.push('未指定 --module，只做全量 declared/installed 一致性体检');
  }
  for (const m of mods) {
    const spec = MODULE_DEPS[m];
    if (!spec) {
      res.warnings.push('未知模块 "' + m + '"，跳过其依赖检查（可用：' + Object.keys(MODULE_DEPS).join(', ') + '）');
      continue;
    }
    for (const n of spec.required) {
      if (!wanted.has(n)) wanted.set(n, { required: true, modules: [] });
      wanted.get(n).modules.push(m);
    }
    for (const n of spec.optional || []) {
      if (!wanted.has(n)) wanted.set(n, { required: false, modules: [] });
      wanted.get(n).modules.push(m);
    }
  }

  for (const [name, meta] of wanted) {
    const inst = installedVersion(root, name);
    const decl = declared.has(name);
    const row = { pkg: name, declared: decl, installed: inst, required: meta.required, modules: meta.modules };
    res.details.push(row);
    if (!inst) {
      if (meta.required) {
        res.pass = false;
        res.errors.push('缺包 ' + name + '（' + meta.modules.join('/') + ' 需要）');
      } else {
        res.warnings.push('可选包未安装 ' + name + '（' + meta.modules.join('/') + '）');
      }
    } else if (!decl && !reachable.has(name)) {
      // 孤立安装：没有任何已声明包依赖它，npm i 会当垃圾清掉
      const msg =
        '包 ' + name + ' 是孤立安装（未声明、也无已声明包依赖它）—— npm i 会把它删掉';
      if (meta.required) {
        res.pass = false;
        res.errors.push(msg);
      } else {
        res.warnings.push(msg);
      }
    }
  }
  return res;
}

// ---- 检查 2/3：页面内插件自证 + facade 断言 ----
async function checkPage(url, facadeExprs, root) {
  const res = { name: 'page', pass: true, errors: [], warnings: [], details: {} };

  // playwright 优先从被测工程自己的 node_modules 解析（本脚本常被复制到工程外运行）
  const tryResolve = name => {
    const tried = [name, path.join(root, 'node_modules', name)];
    for (const t of tried) {
      try {
        return require(t);
      } catch (e) {
        /* next */
      }
    }
    return null;
  };
  const pw = tryResolve('playwright') || tryResolve('playwright-core');
  if (!pw || !pw.chromium) {
    // 指定了 --url 却无法开页 = 门禁没有真正断言，不能算通过
    res.pass = false;
    res.errors.push(
      '指定了 --url 但找不到 playwright（已试 ' + path.join(root, 'node_modules', 'playwright') + '），无法做插件自证'
    );
    return res;
  }
  const { chromium } = pw;
  const browser = await chromium.launch({ channel: 'chrome', headless: true });
  try {
    // 对比工程用自签 HTTPS（node tools/serve.js），必须放行证书
    const ctx = await browser.newContext({ ignoreHTTPSErrors: true, viewport: { width: 1600, height: 1000 } });
    const page = await ctx.newPage();
    let resp = null;
    try {
      resp = await page.goto(url, { waitUntil: 'networkidle', timeout: 60000 });
    } catch (e) {
      res.pass = false;
      res.errors.push('页面打不开：' + url + ' → ' + String((e && e.message) || e).split('\n')[0]);
      return res;
    }
    if (!resp || !resp.ok()) {
      res.pass = false;
      res.errors.push('页面打不开：' + url + '（HTTP ' + (resp ? resp.status() : 'no-response') + '）');
      return res;
    }
    await page.waitForTimeout(1500);

    // 2. 插件自证
    // 三类取值要分开看：ok=注册成功；"duplicated/already registered"=preset 已注册过、
    // 能力其实在（判 FAIL 会误伤）；其余 skip=真缺件。
    const plugins = await page.evaluate(() => window.__proExtraPlugins || null);
    if (!plugins) {
      res.pass = false;
      res.errors.push('window.__proExtraPlugins 不存在 —— entry 没写自证清单，等于无法证明插件装齐');
    } else {
      const keys = Object.keys(plugins);
      const benign = k => plugins[k] === 'ok' || /duplicat|already/i.test(String(plugins[k]));
      const bad = keys.filter(k => !benign(k));
      const dup = keys.filter(k => plugins[k] !== 'ok' && benign(k));
      res.details.plugins = { total: keys.length, bad, duplicated: dup };
      for (const k of dup) res.warnings.push('插件 ' + k + ' 被 preset 重复注册（能力在，不阻塞）：' + plugins[k]);
      if (bad.length) {
        res.pass = false;
        for (const k of bad) res.errors.push('插件未就绪 ' + k + ' → ' + plugins[k]);
      }
    }

    // 3. facade 断言
    if (facadeExprs.length) {
      res.details.facade = [];
      for (const expr of facadeExprs) {
        let v = null;
        let err = null;
        try {
          v = await page.evaluate(x => {
            /* eslint-disable no-new-func */
            const fn = new Function('return (' + x + ')');
            return typeof fn();
          }, expr);
        } catch (e) {
          err = String((e && e.message) || e).slice(0, 160);
        }
        const ok = v !== 'undefined' && v !== null && !err;
        res.details.facade.push({ expr, type: v, ok, err });
        if (!ok) {
          res.pass = false;
          res.errors.push('facade 不存在：' + expr + (err ? '（' + err + '）' : '（typeof undefined）'));
        }
      }
    }
  } finally {
    await browser.close().catch(() => {});
  }
  return res;
}

// ---- 主流程 ----
(async () => {
  const args = parseArgs(process.argv.slice(2));
  const report = { root: args.root, time: new Date().toISOString(), checks: [], pass: true };

  const deps = checkDeps(args.root, args.modules);
  report.checks.push(deps);

  if (args.url) {
    const page = await checkPage(args.url, args.facade, args.root);
    report.checks.push(page);
  } else if (args.facade.length) {
    report.checks.push({
      name: 'page',
      pass: true,
      skipped: true,
      warnings: ['指定了 --facade 但没给 --url，跳过页面内断言'],
      errors: [],
    });
  }

  report.pass = report.checks.every(c => c.pass);

  if (args.json) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    const line = s => console.log(s);
    line('');
    line('Univer 前置门禁  ' + (report.pass ? 'PASS' : 'FAIL'));
    line('root: ' + args.root);
    line('─'.repeat(60));
    for (const c of report.checks) {
      line('[' + c.name + '] ' + (c.skipped ? 'SKIP' : c.pass ? 'ok' : 'FAIL'));
      for (const e of c.errors || []) line('  ✗ ' + e);
      for (const w of c.warnings || []) line('  ! ' + w);
      if (c.details && Array.isArray(c.details)) {
        for (const r of c.details) {
          line(
            '  ' + (r.installed ? '✓' : '✗') + ' ' + r.pkg + (r.installed ? '@' + r.installed : ' 未安装')
          );
        }
      }
      if (c.details && c.details.plugins) {
        const p = c.details.plugins;
        line('  插件 ' + p.total + ' 个，缺件 ' + p.bad.length + ' 个，重复注册 ' + (p.duplicated || []).length + ' 个');
      }
      if (c.details && c.details.facade) {
        for (const f of c.details.facade) line('  ' + (f.ok ? '✓' : '✗') + ' ' + f.expr + ' → ' + (f.type || f.err));
      }
    }
    line('─'.repeat(60));
    if (report.pass) {
      line('门禁通过，可以跑批。注意：判定为 FAIL 后仍要回看本条记录，确认不是装配问题。');
    } else {
      line('门禁未过 —— 不许开跑，也不许把当前结果记成 FAIL。先修环境。');
    }
    line('');
  }

  process.exit(report.pass ? 0 : 1);
})().catch(e => {
  console.error('preflight 异常: ' + ((e && e.stack) || e));
  process.exit(1);
});
