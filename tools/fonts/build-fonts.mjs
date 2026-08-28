#!/usr/bin/env node
// ==========================================================================
// build-fonts.mjs — Google Fonts を必要な字だけ取り寄せて自己ホストする。
//
// ⚠️ 正本。GIGAyama.github.io/standards/fonts/ を直してから配ること。
//
// 使い方:  node tools/fonts/build-fonts.mjs      （fonts.config.json を読む）
//
// ⚠️ --check は無い。生成物のずれは各リポジトリの品質ゲートと SW の版づけが見る。
//    ここに書いていない検査は「していない」と読むこと。
//
// なぜ要るか:
//   実行時に fonts.googleapis.com を読むと、学校のフィルタが塞いだ日に
//   画面が出ない。塞がり方が「拒否」ではなく「握ったまま返さない」だと、
//   スタイルシートの読み込みが終わらず script が動き出さないため、
//   真っ白のまま何も起きない（Quarto で 2026-08-23 に実際に起きた）。
//   ビルド時にここで取り寄せてしまえば、配信物からは外部通信が消える。
//
// ⚠️⚠️ text= は約 800 字で黙って効かなくなる（2026-08-28 実測）
//   Google Fonts の CSS API は、text= が長すぎると **エラーを返さない**。
//   HTTP 200 のまま、字を絞っていない 122 面ぶんの CSS を返す。
//
//       800 字 / URL 6,581 B → @font-face 1 個（100 KB）
//       806 字 / URL 6,635 B → @font-face 122 個（＝ text= が無視された）
//
//   このとき最初の面だけを採ると、ごく狭い範囲の 1 KB ほどのフォントが
//   書き出され、画面のほとんどが端末フォントに落ちる。**それでもビルドは
//   成功したように見える。** だから字を束に割り、1 リクエスト＝1 面に
//   なっていることを毎回確かめる（assertSingleFace）。
//   束ごとの unicode-range は互いに素に返るので、ブラウザは要る束だけ取る。
// ==========================================================================
import fs from 'node:fs';
import path from 'node:path';
import { pathToFileURL } from 'node:url';

import { buildCharset, stripComments } from './chars.mjs';
import { oflText } from './ofl.mjs';

// Google Fonts の CSS API は User-Agent で返す形式を変える。
// woff2 を受け取るために、新しめの Chrome を名乗る。
export const UA =
  'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36';

/** 1 リクエストに載せる字数の上限。実測の崖（約 800）より少し内側に置く。 */
export const DEFAULT_BUCKET_SIZE = 780;

export const DEFAULT_CONFIG = {
  family: null, // 例: "Zen Maru Gothic"（families を使わないなら必須）
  weights: [400, 700],
  families: null, // 複数の書体を使うとき: [{ family, weights }, …]
  grades: [1, 2], // 学年別漢字のどこまでを入れるか
  extra: '', // 追加で必ず入れたい字
  scan: [], // 画面に出る字を拾うファイル / ディレクトリ
  scanExt: ['.html', '.css', '.js', '.jsx', '.ts', '.tsx'],
  // コメントの字を収録しない。既定で入れる（画面に出ない字を配る理由がない）
  stripComments: true,
  outDir: 'fonts', // woff2 の置き場（embed: true のときは使わない）
  cssPath: 'fonts.css', // 生成する CSS
  hrefPrefix: './fonts/', // CSS から woff2 をどう指すか
  bucketSize: DEFAULT_BUCKET_SIZE,
  embed: false, // true: woff2 を base64 の data: URI で CSS に埋める（GAS 用）
  // 出力を <style> で包むか。既定は cssPath が .html なら包む。
  // GAS の include() は中身を **HTML として読み直して組み立て直す**ので、
  // 素の CSS を置くと文字として扱われる。.html に出す以上は必ず包むこと。
  wrapStyle: null,
  slug: null, // ファイル名の頭。既定は family から作る
  license: '', // CSS の頭に書く著作権表記
  copyright: '', // OFL.txt の先頭に置く著作権表示（例: "Copyright 2021 The …"）
  oflPath: null, // OFL 全文の置き場。既定は outDir/OFL.txt。embed のときは必須
  generator: 'tools/fonts/build-fonts.mjs', // CSS に「作り直し方」として書く道具の場所
};

/** family から woff2 のファイル名の頭を作る（"Zen Maru Gothic" → "zen-maru-gothic"） */
export function slugify(family) {
  return String(family)
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, '-')
    .replace(/^-|-$/g, '');
}

/** 字を束に割る。1 束が崖（約 800 字）を越えないようにするためのもの。 */
export function splitBuckets(chars, size = DEFAULT_BUCKET_SIZE) {
  if (!Number.isInteger(size) || size < 1) throw new Error(`bucketSize が不正: ${size}`);
  if (size > 800) {
    // ここを越えると text= が黙って無視される。設定で踏めてしまうと
    // 「通ったのに何も絞れていない」に戻るので、設定の時点で止める。
    throw new Error(
      `bucketSize は 800 以下にすること（指定: ${size}）。` +
        'Google Fonts の text= は約 800 字で黙って効かなくなる。',
    );
  }
  const all = [...chars];
  const out = [];
  for (let i = 0; i < all.length; i += size) out.push(all.slice(i, i + size).join(''));
  return out;
}

/**
 * 返ってきた CSS が「1 束ぶん」になっていることを確かめる。
 *
 * ⚠️ この検査がこの道具の要。落とさないこと。
 *    text= が無視されると 122 面が返るが、HTTP は 200 のままなので
 *    ここで見なければ気づけない。
 */
export function assertSingleFace(css, label) {
  const faces = (css.match(/@font-face/g) || []).length;
  if (faces !== 1) {
    throw new Error(
      `${label}: Google Fonts が @font-face を ${faces} 個返した（1 個のはず）。` +
        'text= が長すぎて無視された可能性が高い。fonts.config.json の ' +
        'bucketSize を小さくするか、grades を減らすこと。',
    );
  }
}

/** CSS から woff2 の URL と unicode-range を取り出す */
export function parseFace(css, label) {
  assertSingleFace(css, label);
  const src = css.match(/src:\s*url\(([^)]+)\)/);
  if (!src) throw new Error(`${label}: woff2 の URL が見つからない`);
  const range = css.match(/unicode-range:\s*([^;]+);/);
  if (!range) throw new Error(`${label}: unicode-range が見つからない`);
  return { url: src[1], unicodeRange: range[1].trim() };
}

/** 取り寄せる URL を組み立てる */
export function cssApiUrl(family, weight, text) {
  // ビルド時にここから取り寄せて自己ホストする側。配信物には外部通信が残らない。
  // giga-lint-ignore-next-line
  const url = new URL('https://fonts.googleapis.com/css2');
  url.searchParams.set('family', `${family}:wght@${weight}`);
  url.searchParams.set('text', text);
  url.searchParams.set('display', 'swap');
  return url;
}

/** @font-face を並べた CSS を組み立てる */
export function renderCss(faces, { family, license = '', generator = 'tools/fonts/build-fonts.mjs' }) {
  const head = `/* ==========================================================================
 * ${family}
${license ? ` * ${license}\n` : ''} *
 * このファイルは ${generator} が生成する。**直接編集しないこと。**
 * 作り直すには: node ${generator}
 *
 * unicode-range は Google Fonts が返した値をそのまま使っている。ここに無い字
 * （児童が入力した珍しい漢字など）は、このフォントを飛ばして端末内蔵フォントへ
 * 落ちる。書体が変わるだけで、□（豆腐）にはならない。
 *
 * 面が複数あるのは、1 リクエストに載せられる字数に上限があるため。
 * ブラウザは unicode-range を見て、実際に要る面だけを取りにいく。
 * ========================================================================== */
`;
  const body = faces
    .map(
      (f) => `@font-face {
  font-family: '${f.family || family}';
  font-style: normal;
  font-weight: ${f.weight};
  font-display: swap;
  src: url('${f.href}') format('woff2');
  unicode-range: ${f.unicodeRange};
}`,
    )
    .join('\n\n');
  return `${head}\n${body}\n`;
}

// --- 設定と入力 ------------------------------------------------------------

export function loadConfig(repoRoot, readFile = fs.readFileSync) {
  const p = path.join(repoRoot, 'fonts.config.json');
  let raw;
  try {
    raw = readFile(p, 'utf8');
  } catch {
    throw new Error(`fonts.config.json が無い: ${p}`);
  }
  const cfg = { ...DEFAULT_CONFIG, ...JSON.parse(raw) };
  // 1 書体でも複数でも、内部では families の並びとして扱う
  if (!Array.isArray(cfg.families) || cfg.families.length === 0) {
    if (!cfg.family) throw new Error('fonts.config.json に family も families もない');
    cfg.families = [{ family: cfg.family, weights: cfg.weights }];
  }
  for (const f of cfg.families) {
    if (!f.family) throw new Error('fonts.config.json の families に family がない');
    if (!Array.isArray(f.weights) || f.weights.length === 0) {
      throw new Error(`fonts.config.json の ${f.family} の weights が空`);
    }
    f.slug = f.slug || slugify(f.family);
  }
  cfg.slug = cfg.slug || cfg.families[0].slug;
  return cfg;
}

/**
 * 画面に出る字を、設定された置き場から拾う。
 *
 * ⚠️ 生成物（cssPath と outDir）は必ず外すこと。
 *    生成した fonts.css には、この道具が書く日本語のコメントが入っている。
 *    それを次の走査で拾うと、そのコメントに出てくる漢字が収録対象に増える。
 *    実測で 744 字 → 768 字に増えた（2026-08-28）。入力から出力が決まるはずが、
 *    出力が入力に混ざるので、走らせるたびに収録内容が変わりうる。
 */
export function collectSources(repoRoot, cfg, deps = {}) {
  const { readdir = fs.readdirSync, read = fs.readFileSync, stat = fs.statSync } = deps;
  const chunks = [];
  const generated = new Set(
    [cfg.cssPath, cfg.outDir, cfg.oflPath]
      .filter(Boolean)
      .map((rel) => path.resolve(repoRoot, rel)),
  );
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
        // 配信しない置き場と、生成物の巣は歩かない
        if (['node_modules', '.git', 'dist', 'vendor', 'fonts'].includes(name)) continue;
        visit(path.join(p, name));
      }
      return;
    }
    if (!cfg.scanExt.includes(path.extname(p))) return;
    try {
      const text = read(p, 'utf8');
      chunks.push(cfg.stripComments === false ? text : stripComments(text));
    } catch {
      /* 読めないものは飛ばす */
    }
  };
  for (const rel of cfg.scan) visit(path.join(repoRoot, rel));
  return chunks.join('');
}

// --- 本体 ------------------------------------------------------------------

export async function buildFonts(repoRoot, { fetchImpl = fetch, log = console.log } = {}) {
  const cfg = loadConfig(repoRoot);
  const sources = collectSources(repoRoot, cfg);
  const chars = buildCharset({ grades: cfg.grades, extra: cfg.extra, sources });
  const buckets = splitBuckets(chars, cfg.bucketSize);

  log(
    `書体: ${cfg.families.map((f) => `${f.family}(${f.weights.join('/')})`).join('  ')}`,
  );
  log(`収録する字: ${chars.length} 字 → ${buckets.length} 束（1 束 ${cfg.bucketSize} 字まで）`);

  const outDir = path.join(repoRoot, cfg.outDir);
  if (!cfg.embed) fs.mkdirSync(outDir, { recursive: true });

  const faces = [];
  let total = 0;
  for (const fam of cfg.families) {
  for (const weight of fam.weights) {
    for (const [i, text] of buckets.entries()) {
      const label = `${fam.family} ${weight} 束${i + 1}`;
      const cssRes = await fetchImpl(cssApiUrl(fam.family, weight, text), {
        headers: { 'User-Agent': UA },
      });
      if (!cssRes.ok) throw new Error(`${label}: CSS の取得に失敗 (${cssRes.status})`);
      const { url, unicodeRange } = parseFace(await cssRes.text(), label);

      const fontRes = await fetchImpl(url, { headers: { 'User-Agent': UA } });
      if (!fontRes.ok) throw new Error(`${label}: woff2 の取得に失敗 (${fontRes.status})`);
      const buf = Buffer.from(await fontRes.arrayBuffer());
      total += buf.length;

      let href;
      if (cfg.embed) {
        // GAS は HtmlService でバイナリを配れないので、CSS に焼きこむ
        href = `data:font/woff2;base64,${buf.toString('base64')}`;
      } else {
        const file = `${fam.slug}-${weight}-${i + 1}.woff2`;
        fs.writeFileSync(path.join(outDir, file), buf);
        href = `${cfg.hrefPrefix}${file}`;
        log(`  ${file}  ${(buf.length / 1024).toFixed(1)} KB`);
      }
      faces.push({ family: fam.family, weight, bucket: i + 1, href, unicodeRange });
    }
  }
  }

  // OFL のフォントを自分のところから配る条件。フォントだけ置いて本文を置き忘れると
  // ライセンス違反になるので、道具のほうで必ず書き出す。
  const oflPath = cfg.oflPath || path.join(cfg.embed ? '.' : cfg.outDir, 'OFL.txt');
  fs.mkdirSync(path.dirname(path.join(repoRoot, oflPath)), { recursive: true });
  fs.writeFileSync(path.join(repoRoot, oflPath), oflText(cfg.copyright || cfg.license));

  const css = renderCss(faces, {
    family: cfg.families.map((f) => f.family).join(' / '),
    license: cfg.license,
    generator: cfg.generator,
  });
  const wrap =
    cfg.wrapStyle === null || cfg.wrapStyle === undefined
      ? path.extname(cfg.cssPath) === '.html'
      : cfg.wrapStyle;
  fs.writeFileSync(path.join(repoRoot, cfg.cssPath), wrap ? `<style>\n${css}</style>\n` : css);
  log(
    `✅ ${cfg.cssPath} と ${oflPath} を更新した（@font-face ${faces.length} 面 / ` +
      `フォント計 ${(total / 1024).toFixed(1)} KB${cfg.embed ? '・CSS に埋め込み' : ''}）`,
  );
  return { chars, buckets: buckets.length, faces, bytes: total, oflPath };
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  buildFonts(process.cwd()).catch((err) => {
    console.error(`❌ ${err.message}`);
    process.exit(1);
  });
}
