# GIGAスクール Webアプリ — Claude Code 開発ガイド

@.agents/rules/gigaschool-standards.md

---

## このファイルについて

⚠️ **これは配布物です。正本は
`GIGAyama.github.io/standards/agents/CLAUDE.md`。**
ここを直しても他のリポジトリには届きません。正本を直してから
`node tools/distribute.mjs` で配ってください。

⚠️ **上の 1 行（`@.agents/rules/…`）を消さないこと。**
Zero-CDN・Zero-PII・正本同期といった艦隊共通のルールは、
すべてその 1 本のファイルに書いてあります。Antigravity（Gemini）は
`.agents/rules/` を直接読みますが、Claude Code はこの取りこみを通して読みます。
消すと、Claude Code だけがルールを知らないまま作業を始めます。

同じ内容を 2 か所に書かないための形です。ルール本文を
ここへ写さないでください（写した瞬間、どちらが正本か決められなくなります）。

## このリポジトリ固有のことを足すには

この見出しより下に書いてください。ただし**このファイル自体が
正本と 1 バイトでも違うと `check-drift` が赤くなります。**
固有の内容を持ちたいリポジトリは、`standards-map.json` の
`unmanaged` に理由つきで宣言してから書き替えてください。

```json
{
  "unmanaged": [
    { "local": "CLAUDE.md",
      "reason": "このアプリ固有の手引きを持つ（YYYY-MM-DD）。冒頭の取りこみ 1 行は残す" }
  ]
}
```

宣言したあとも、**冒頭の `@.agents/rules/…` の行が在ることは
`tools/check-distribution.mjs` が見ています。**
宣言しただけで無検査になる、という形にはしていません。
