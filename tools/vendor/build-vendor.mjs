#!/usr/bin/env node
// ==========================================================================
// build-vendor.mjs — ライブラリを npm から取りこんで自己ホストする。
//
// ⚠️ 正本。GIGAyama.github.io/standards/vendor/ を直してから配ること。
//
// 使い方:  node tools/vendor/build-vendor.mjs   （vendor.config.json を読む）
//
// なぜ要るか（MIRAI-Compass で実際に起きたこと）:
//   学校のネットワークは cdn.jsdelivr.net を塞いでいることがある。
//   塞がれた状態で開くと、白い画面ではなく「Bootstrap が当たっていない
//   素の HTML が半分だけ動く」という壊れ方をする。ローディング画面が消えず、
//   d-none が効かないので児童画面と先生用ボタンが同時に出て、
//   Swal / Chart / Sortable がすべて undefined になる。
//   児童からは「壊れている」としか見えず、原因はアプリの外にあるので
//   先生が調べても分からない。
//
//   そこで、実行コードは 1 バイトも外から取らない。npm で版を固定し、
//   ここで包んでから配る。
//
// ⚠️ アイコンは webfont を取りこまない。使っている分だけ SVG マスクにする。
//   bootstrap-icons を丸ごと持つと CSS 98KB + woff2 131KB = 229KB になるが、
//   実際に使うのは全 2000 種類のうち数十種類でしかない。
//   マークアップ（<i class="bi bi-compass-fill">）は変えず、
//   background-color: currentColor なので text-danger などの色指定も効く。
//
// ⚠️⚠️ 走査の取りこぼしは「その画面だけ絵が消える」という形で出る
//   MIRAI-Compass では js_worksheet.html が一覧から抜けていて、
//   ワークシート画面のツールバーだけアイコン無しになっていた。
//   ビルドは通り、他の画面は正しく出るので、**開いて見るまで気づけない。**
//   だから既定の走査対象は「リポジトリ内の .html / .js 全部」にしてある。
//   vendor.config.json の scan で狭めるときは、その危険を承知で狭めること。
// ==========================================================================
import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

export const DEFAULT_CONFIG = {
  // 走査する場所。既定はリポジトリ全体（取りこぼすより重いほうがまし）
  scan: ['.'],
  scanExt: ['.html', '.js', '.jsx', '.ts', '.tsx'],
  // 歩かない置き場
  skipDirs: ['node_modules', '.git', 'dist', 'build', 'vendor', 'fonts', 'legacy'],
  targets: [],
};

/** 生成物の頭に必ず置く注意書き。手で編集されると原本と食い違う。 */
export function banner(what, generator) {
  return `<!--
  ⚠️ このファイルは ${generator} が生成しています。手で編集しないでください。
     直す場所は package.json（版）と vendor.config.json（組み立て方）です。
     直したら必ず \`npm run build\` を走らせてから push してください。
     中身: ${what}
-->
`;
}

/** node_modules の中身を読む。無ければ「npm ci して」と言って止まる。 */
export function readDep(nodeModules, rel) {
  const file = path.join(nodeModules, ...rel.split('/'));
  if (!fs.existsSync(file)) {
    throw new Error(
      `${rel} がありません。先に \`npm ci\`（または npm install）を実行してください。`,
    );
  }
  return fs.readFileSync(file, 'utf8');
}

/**
 * 走査して、使われているアイコンの名前を集める。
 *
 * prefix は 'bi' や 'mdi' など。`bi-arrow-left` の `arrow-left` を返す。
 *
 * ⚠️ 生成物は必ず外すこと。
 *   生成した vendor_icons.html には `--bi-i` という変数名が入っている。
 *   それを次の走査で拾うと `bi-i` というアイコンが使われていることになり、
 *   走らせるたびに中身が変わる（1 回目 1 個 → 2 回目 2 個）。
 *   入力から出力が決まらなくなるので、verify-generated が常に落ちる。
 *   fonts.css で同じことが起きている（standards/fonts/build-fonts.mjs の
 *   collectSources を見ること）。
 */
export function collectIconNames(repoRoot, cfg, prefix, deps = {}) {
  const { readdir = fs.readdirSync, read = fs.readFileSync, stat = fs.statSync } = deps;
  const used = new Set();
  // ⚠️ _ を入れること。Material Symbols の名前は check_circle のように _ を使う。
  //    入れ忘れると ms-check_circle が "check" として拾われ、**別の絵が出るか、
  //    何も出ない**。実際に踏んだ（2026-08-28）。
  const re = new RegExp(`\\b${prefix}-([a-z0-9_-]+)`, 'g');
  const generated = new Set(
    (cfg.targets || []).map((t) => path.resolve(repoRoot, t.out)),
  );
  /* ⚠️ 生成物の「名前」も落とす。
     index.html に <link href="./vendor/mdi-icons.css"> と書いてあると、
     その文字列から `mdi-icons` というアイコンが使われていると読んでしまう。
     ファイルを走査対象から外すだけでは足りない（読む側ではなく、
     読まれる側に名前が出てくるため）。2026-08-28、SekigaeMaker で実測。 */
  /* ⚠️ 拡張子まで含めた「ファイル名」で落とすこと。
     拡張子を外した幹（"ms"）で落とすと、`class="ms ms-search"` の中の
     ms まで消えてアイコンが 1 つも見つからなくなる。実際に踏んだ。 */
  const generatedNames = (cfg.targets || []).map((t) => path.basename(t.out));
  const dropNames = (text) => {
    let out = text;
    for (const n of generatedNames) out = out.split(n).join('');
    return out;
  };
  const visit = (p) => {
    if (generated.has(path.resolve(p))) return;
    let st;
    try {
      st = stat(p);
    } catch {
      return;
    }
    if (st.isDirectory()) {
      for (const name of readdir(p)) {
        if (cfg.skipDirs.includes(name)) continue;
        visit(path.join(p, name));
      }
      return;
    }
    if (!cfg.scanExt.includes(path.extname(p))) return;
    let text;
    try {
      text = read(p, 'utf8');
    } catch {
      return;
    }
    for (const m of dropNames(text).matchAll(re)) used.add(m[1]);
  };
  for (const rel of cfg.scan) visit(path.join(repoRoot, rel));
  return [...used].sort();
}

/** SVG 1 枚を、CSS の mask-image に載る data: URI にする */
export function svgToDataUri(svg) {
  const body = svg
    .replace(/<\?xml[^>]*\?>/g, '')
    .replace(/<!--[\s\S]*?-->/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    // currentColor で塗るので、SVG 側の色指定は落として統一する
    .replace(/\sfill="[^"]*"/g, '')
    /* class / width / height も落とす。
       マスクの大きさは CSS（1em・center/contain）が決めるので効かないうえ、
       bootstrap-icons の SVG は class="bi bi-rulers" のように**自分の名前**を
       持っている。それが data: URI の中に残ると、走査で「そのアイコンも
       使われている」と読めてしまう。実際に MIRAI-Compass で
       bi-pencil-ruler の中の bi-rulers を拾った（2026-08-28）。 */
    .replace(/\s(?:class|width|height)="[^"]*"/g, '')
    .replace('<svg ', '<svg fill="currentColor" ');
  const esc = body
    .replace(/"/g, "'")
    .replace(/[<>#%{}|\\^~[\]`]/g, (c) => '%' + c.charCodeAt(0).toString(16).toUpperCase());
  return `data:image/svg+xml,${esc}`;
}

/** アイコンの CSS を組み立てる。使っている名前のぶんだけ。 */
export function renderIconCss(names, { prefix, baseClass, resolve }) {
  const rules = [];
  const missing = [];
  for (const name of names) {
    const svg = resolve(name);
    if (svg == null) {
      missing.push(name);
      continue;
    }
    rules.push(`.${prefix}-${name}{--${prefix}-i:url("${svgToDataUri(svg)}")}`);
  }
  const base = [
    `/* ${prefix}-* のうち、このアプリが実際に使っている分だけ。`,
    `   マスク方式なので currentColor / font-size(1em) がそのまま効く。 */`,
    `.${baseClass}{display:inline-block;width:1em;height:1em;vertical-align:-.125em;flex-shrink:0;`,
    `background-color:currentColor;`,
    `-webkit-mask:var(--${prefix}-i) center/contain no-repeat;mask:var(--${prefix}-i) center/contain no-repeat}`,
    `/* 高コントラストモードでは mask が消える。絵が無くても意味が失われないよう、`,
    `   アイコンだけのボタンには aria-label を付けること。 */`,
    `@media (forced-colors: active){.${baseClass}{background-color:CanvasText}}`,
  ].join('\n');
  return { css: `${base}\n${rules.join('\n')}`, count: rules.length, missing };
}

/** JavaScript を <script> に包めるようにする */
export function wrapJs(js) {
  return (
    js
      .replace(/\/\/#\s*sourceMappingURL=.*$/gm, '')
      // </script> が文字列の中に現れると、そこで <script> が閉じてしまう
      .replace(/<\/script>/gi, '<\\/script>')
      .trim()
  );
}

/** CSS を <style> に包めるようにする */
export function wrapCss(css) {
  // sourceMappingURL が残っていると、開発者ツールが取りに行って 404 を出す
  return css.replace(/\/\*#\s*sourceMappingURL=.*?\*\//g, '').trim();
}

export function loadConfig(repoRoot, readFile = fs.readFileSync) {
  const p = path.join(repoRoot, 'vendor.config.json');
  let raw;
  try {
    raw = readFile(p, 'utf8');
  } catch {
    throw new Error(`vendor.config.json が無い: ${p}`);
  }
  const cfg = { ...DEFAULT_CONFIG, ...JSON.parse(raw) };
  if (!Array.isArray(cfg.targets) || cfg.targets.length === 0) {
    throw new Error('vendor.config.json の targets が空');
  }
  for (const t of cfg.targets) {
    if (!t.out) throw new Error('vendor.config.json の targets に out がない');
    if (!t.css && !t.js && !t.icons) {
      throw new Error(`vendor.config.json の ${t.out} に css / js / icons のどれも無い`);
    }
  }
  return cfg;
}

/** 生成物の一覧（verify-generated が使う） */
export function outputs(cfg) {
  return cfg.targets.map((t) => t.out);
}

// --- アイコンの出どころ ----------------------------------------------------
// パッケージごとに SVG の置き場と読み替えが違う。ここに 1 か所だけ持つ。
export const ICON_PACKS = {
  'bootstrap-icons': {
    prefix: 'bi',
    baseClass: 'bi',
    dir: 'bootstrap-icons/icons',
    // 存在しない名前の読み替え。pencil-ruler は無く、いま何も描かれていない
    alias: { 'pencil-ruler': 'rulers' },
  },
  '@mdi/svg': {
    prefix: 'mdi',
    baseClass: 'mdi',
    dir: '@mdi/svg/svg',
    alias: {},
  },
  /* Material Symbols。合字で絵を出す書体としても配れるが、そちらは
     「その字で綴れる絵ぜんぶ」が返ってくるので重い（実測 289KB）。
     使っている分の SVG だけを取り出せば 10KB 前後で済む。
     FILL@1 の見た目に合わせて -fill.svg を採る。 */
  '@material-symbols/svg-400': {
    prefix: 'ms',
    baseClass: 'ms',
    dir: '@material-symbols/svg-400/rounded',
    suffix: '-fill',
    alias: {},
  },
};

export async function buildVendor(repoRoot, { log = console.log, warn = console.warn } = {}) {
  const cfg = loadConfig(repoRoot);
  const nm = path.join(repoRoot, 'node_modules');
  const generator = cfg.generator || 'tools/vendor/build-vendor.mjs';
  const kb = (s) => (Buffer.byteLength(s, 'utf8') / 1024).toFixed(1) + ' KB';
  const wrap = cfg.wrap !== 'none';
  const results = [];

  for (const t of cfg.targets) {
    let body = '';
    let what = '';
    let missingIcons = [];

    if (t.css) {
      const parts = t.css.map((rel) => wrapCss(readDep(nm, rel)));
      what = t.css.join(' + ');
      body = wrap ? `<style>\n${parts.join('\n')}\n</style>\n` : parts.join('\n') + '\n';
    } else if (t.icons) {
      const pack = ICON_PACKS[t.icons];
      if (!pack) throw new Error(`知らないアイコンの出どころ: ${t.icons}`);
      /* ⚠️ パッケージが入っていないときは、警告ではなく止める。
         2026-08-28、CI に npm ci が無いリポジトリでこれをやったところ、
         「SVG が見つからないアイコン: （47 個ぜんぶ）」と警告を出したうえで
         **アイコンが 0 個の 1KB の CSS を書き出して成功した**。
         配れば画面から絵が全部消えるのに、ビルドは緑になる。
         1 個も見つからないのは打ち間違いではなく、取りこみ元が無いということ。 */
      if (!fs.existsSync(path.join(nm, pack.dir))) {
        throw new Error(
          `${t.icons} がありません（${pack.dir}）。先に \`npm ci\`（または npm install）を実行してください。`,
        );
      }
      /* ⚠️ 実行時に決まるアイコン（`ms-${cond ? 'a' : 'b'}` のような書き方）は
         走査では見つからない。設定の extra に並べること。
         並べ忘れると、その場面でだけ絵が消える。 */
      const names = [...new Set([...collectIconNames(repoRoot, cfg, pack.prefix), ...(t.extra || [])])].sort();
      const { css, count, missing } = renderIconCss(names, {
        prefix: pack.prefix,
        baseClass: pack.baseClass,
        resolve: (name) => {
          /* 読み替えは 2 段。出どころが持つ表（版が変わって名前が消えたもの）と、
             リポジトリが vendor.config.json に書いた表（そのアプリの事情）。
             後者が優先。読み替えの理由は設定のそばに書けるようにしてある。 */
          const real = (t.alias && t.alias[name]) || pack.alias[name] || name;
          const file = path.join(nm, pack.dir, real + (pack.suffix || '') + '.svg');
          return fs.existsSync(file) ? fs.readFileSync(file, 'utf8') : null;
        },
      });
      if (missing.length && count === 0 && names.length > 0) {
        // 1 個も引けなかった。打ち間違いが全部に起きることは無いので、
        // 出どころのほうが壊れている（版が違う／中身が空）。
        throw new Error(
          `${t.icons} から 1 個も引けませんでした（${names.length} 個ぜんぶ）。` +
            `版が合っているか、${pack.dir} の中身を確かめてください。`,
        );
      }
      if (missing.length) {
        // 落とさない。名前の打ち間違いでビルドを止めると、直せる人がいないときに
        // 何も配れなくなる。ただし黙って通さない。
        warn(`[build-vendor] SVG が見つからないアイコン: ${missing.join(', ')}`);
      }
      missingIcons = missing;
      what = `${t.icons} から ${count} 個（使用分のみ / 走査 ${names.length} 名）`;
      body = wrap ? `<style>\n${css}\n</style>\n` : css + '\n';
    } else {
      const parts = t.js.map((rel) => [path.basename(rel), wrapJs(readDep(nm, rel))]);
      what = parts.map(([n]) => n).join(' + ');
      body = wrap
        ? parts.map(([n, js]) => `<!-- ${n} -->\n<script>\n${js}\n</script>\n`).join('')
        : parts.map(([n, js]) => `/* ${n} */\n${js}\n`).join('\n');
    }

    const head = wrap
      ? banner(what, generator)
      : `/* ${what}\n   ${generator} が生成。手で編集しないこと。 */\n`;
    const text = head + body;
    const dest = path.join(repoRoot, t.out);
    fs.mkdirSync(path.dirname(dest), { recursive: true });
    fs.writeFileSync(dest, text);
    log(`${t.out.padEnd(20)} ${kb(text)}  ${what}`);
    results.push({ out: t.out, bytes: Buffer.byteLength(text, 'utf8'), missingIcons });
  }

  log('\n[build-vendor] 完了。生成物は手で編集しないこと。');
  return results;
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  buildVendor(process.cwd()).catch((err) => {
    console.error(`❌ ${err.message}`);
    process.exit(1);
  });
}
