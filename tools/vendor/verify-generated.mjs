#!/usr/bin/env node
// ==========================================================================
// verify-generated.mjs — 生成物が原本と食い違っていないか確かめる。
//
// ⚠️ 正本。GIGAyama.github.io/standards/vendor/ を直してから配ること。
//
// なぜ要るか:
//   vendor_*.html は生成物なので、原本（package.json / vendor.config.json）を
//   直したのに `npm run build` を忘れて push すると、リポジトリの中身だけが
//   古いまま残る。ビルドも静的解析も通ってしまい、動かすまで気づけない。
//   CI でここを踏むと、その取りこぼしが PR の時点で止まる。
//
//   検査する一覧は vendor.config.json の targets から読む。決め打ちにすると、
//   生成物を 1 つ足したときに、その 1 つだけ見られないまま緑になる。
// ==========================================================================
import fs from 'node:fs';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { pathToFileURL } from 'node:url';

import { buildVendor, loadConfig, outputs } from './build-vendor.mjs';

export const hashOf = (file) =>
  createHash('sha256').update(fs.readFileSync(file)).digest('hex').slice(0, 16);

export async function verifyGenerated(repoRoot, { log = console.log, err = console.error } = {}) {
  const files = outputs(loadConfig(repoRoot));

  const before = {};
  for (const f of files) {
    const p = path.join(repoRoot, f);
    if (!fs.existsSync(p)) {
      err(`❌ ${f} がありません。\`npm run build\` を実行してください。`);
      return 1;
    }
    before[f] = hashOf(p);
  }

  await buildVendor(repoRoot, { log: () => {} });

  let bad = 0;
  for (const f of files) {
    const after = hashOf(path.join(repoRoot, f));
    if (before[f] !== after) {
      err(`❌ ${f} が原本と食い違っています（${before[f]} → ${after}）。`);
      bad++;
    } else {
      log(`✅ ${f} は最新です（${after}）。`);
    }
  }

  if (bad) {
    err('\n原本を直したら `npm run build` を走らせてから push してください。');
    return 1;
  }
  return 0;
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  verifyGenerated(process.cwd())
    .then((code) => process.exit(code))
    .catch((e) => {
      console.error(`❌ ${e.message}`);
      process.exit(1);
    });
}
