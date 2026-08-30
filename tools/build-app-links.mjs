#!/usr/bin/env node
/**
 * 正本の giga-app-links.js を、GAS が配れる形（giga_links.html）へ焼き込む。
 *
 *   node tools/build-app-links.mjs           作り直す
 *   node tools/build-app-links.mjs --check   コミットされている中身と食い違わないか見る
 *
 * ── なぜ焼き込みが要るのか ──────────────────────────
 *
 * Apps Script が持てるのは .gs と .html だけで、.js のファイルは置けない。
 * web/giga-app-links.js は GAS 側に存在しないので <script src> では読めず、
 * 相対パスで書くと script.google.com の側を探して 404 になる。
 * vendor_js.html / js.html と同じように、中身を <script> で囲んだ .html にしておき、
 * index.html から HtmlService.createHtmlOutputFromFile で差し込む。
 *
 * ⚠️ 差し込む 1 行を index.html に足すこと。忘れても**エラーにはならない**ので、
 *    画面を開いて見るまで気づけない（リンクが出ないだけ）。
 *
 * ── 正本を直接いじらないこと ────────────────────────
 *
 * web/giga-app-links.js は正本 standards/web/giga-app-links.js の写しで、
 * 42 本すべてに同じものが配られる。ここを直しても他へは届かず、
 * check-drift が赤くなる。直すときはポータル側の正本を直して配ること。
 */
import { readFileSync, writeFileSync } from 'node:fs';
import { join, dirname } from 'node:path';
import { fileURLToPath } from 'node:url';

const ROOT = join(dirname(fileURLToPath(import.meta.url)), '..');
const SRC = join(ROOT, 'web', 'giga-app-links.js');
const OUT = join(ROOT, 'giga_links.html');
const SLUG = 'shared-folder-sync';

function build() {
  const code = readFileSync(SRC, 'utf8');
  return `<!--
  この HTML は tools/build-app-links.mjs が作ります。手で編集しないこと。
  中身は正本 standards/web/giga-app-links.js（配布物 web/giga-app-links.js）です。

  画面から「つかいかた・利用規約・プライバシー」へつなぐリンクを、
  <span data-giga-links> の中に出します（index.html のフッター）。

  ⚠️ slug をここで渡しているのは、GAS の画面が script.google.com（または
     googleusercontent.com）で動いていて、ホスト名から自分が何のアプリか
     分からないためです。渡さないと、部品は何も出しません
     （当てずっぽうに組むと、存在しないアプリの利用規約へ飛ばすことになる）。
-->
<script>window.GIGA_APP_LINKS = { slug: '${SLUG}', links: 'terms,privacy' };</script>
<script>
${code}</script>
`;
}

const html = build();
if (process.argv.includes('--check')) {
  const now = readFileSync(OUT, 'utf8');
  if (now !== html) {
    console.error('❌ giga_links.html が web/giga-app-links.js と食い違っています。');
    console.error('   node tools/build-app-links.mjs で作り直してコミットしてください。');
    process.exit(1);
  }
  console.log('✅ giga_links.html は正本と一致しています');
} else {
  writeFileSync(OUT, html);
  console.log(`giga_links.html  ${(html.length / 1024).toFixed(1)} KB`);
}
