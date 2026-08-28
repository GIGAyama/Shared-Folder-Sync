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
//
// ⚠️ 比べる相手は「git がコミットしている中身」であって、作業ツリーの
//    ファイルではない。
//
//    前の版は作業ツリーのファイルを控えてから作り直して比べていた。
//    これだと `npm run build` を先に走らせるだけで、控える相手が
//    作り直したものになり、何を壊しても必ず緑になる。
//    2026-08-28 に実測した。`ci = npm run build && npm run verify` という
//    書き方をしていた 10 本すべてで、わざと壊した vendor_icons.html が
//    素通りした（build を外すと同じ検査が赤くなった）。
//
//    順番で結果が変わる検査は、いつか必ず順番のほうを間違える。
//    コミットされている中身と比べるなら、前に build が何回走っていても
//    結果は変わらない。
// ==========================================================================
import fs from 'node:fs';
import path from 'node:path';
import { createHash } from 'node:crypto';
import { execFileSync } from 'node:child_process';
import { pathToFileURL } from 'node:url';


export const hashOf = (file) =>
  createHash('sha256').update(fs.readFileSync(file)).digest('hex').slice(0, 16);

const sha = (buf) => createHash('sha256').update(buf).digest('hex').slice(0, 16);

/**
 * git がコミットしている中身（HEAD）を返す。取れないときは null。
 * repoRoot が git リポジトリの途中の階層でも動くよう、接頭辞を足して引く。
 */
export function committedBytes(repoRoot, rel, { run = execFileSync } = {}) {
  try {
    const prefix = run('git', ['-C', repoRoot, 'rev-parse', '--show-prefix'], {
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'ignore'],
    }).trim();
    return run('git', ['-C', repoRoot, 'show', `HEAD:${prefix}${rel}`], {
      stdio: ['ignore', 'pipe', 'ignore'],
      maxBuffer: 256 * 1024 * 1024,
    });
  } catch {
    return null;
  }
}

/** この場所で git が使えるか（コミットが 1 つ以上あるか）。 */
export function hasGit(repoRoot, { run = execFileSync } = {}) {
  try {
    run('git', ['-C', repoRoot, 'rev-parse', 'HEAD'], {
      stdio: ['ignore', 'ignore', 'ignore'],
    });
    return true;
  } catch {
    return false;
  }
}

/**
 * files を作り直して、コミットされている中身と突き合わせる。
 * rebuild は「生成物を作り直す」処理（同期でも Promise でもよい）。
 *
 * vendor 以外の生成物（app.html などの JSX / Tailwind ビルド）も同じ穴を持つので、
 * 各リポジトリの check-generated.mjs からもここを呼ぶ。
 */
export async function verifyAgainstCommitted(
  repoRoot,
  files,
  rebuild,
  { log = console.log, err = console.error, run = execFileSync } = {},
) {
  const git = hasGit(repoRoot, { run });

  // 控える相手を先に決める。git があれば HEAD、無ければ作業ツリー。
  const before = {};
  for (const f of files) {
    if (git) {
      const bytes = committedBytes(repoRoot, f, { run });
      if (bytes === null) {
        err(`❌ ${f} がコミットされていません。生成物も一緒にコミットしてください。`);
        return 1;
      }
      before[f] = sha(bytes);
    } else {
      const p = path.join(repoRoot, f);
      if (!fs.existsSync(p)) {
        err(`❌ ${f} がありません。\`npm run build\` を実行してください。`);
        return 1;
      }
      before[f] = hashOf(p);
    }
  }

  if (!git) {
    err(
      '⚠️ git が使えないので、作業ツリーのファイルと比べます。' +
        'この比べ方は前に `npm run build` が走っていると必ず緑になります。',
    );
  }

  await rebuild();

  let bad = 0;
  for (const f of files) {
    const p = path.join(repoRoot, f);
    if (!fs.existsSync(p)) {
      err(`❌ ${f} が作り直せませんでした。`);
      bad++;
      continue;
    }
    const after = hashOf(p);
    if (before[f] !== after) {
      err(`❌ ${f} が原本と食い違っています（${before[f]} → ${after}）。`);
      bad++;
    } else {
      log(`✅ ${f} は最新です（${after}）。`);
    }
  }

  if (bad) {
    err('\n原本を直したら `npm run build` を走らせ、生成物も一緒にコミットしてください。');
    return 1;
  }
  return 0;
}

export async function verifyGenerated(repoRoot, opts = {}) {
  // build-vendor はここで読む。vendor を持たないリポジトリ（app.html だけを
  // 作るところ）が verifyAgainstCommitted だけを使えるようにするため、
  // ファイルの頭では読まない。
  const { buildVendor, loadConfig, outputs } = await import('./build-vendor.mjs');
  const files = outputs(loadConfig(repoRoot));
  return verifyAgainstCommitted(repoRoot, files, () => buildVendor(repoRoot, { log: () => {} }), opts);
}

if (process.argv[1] && import.meta.url === pathToFileURL(process.argv[1]).href) {
  verifyGenerated(process.cwd())
    .then((code) => process.exit(code))
    .catch((e) => {
      console.error(`❌ ${e.message}`);
      process.exit(1);
    });
}
