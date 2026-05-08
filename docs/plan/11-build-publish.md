# 11. ビルド・公開

## 0. 開発環境（Nix flake）

全コントリビューターで Node / pnpm / Python のバージョンを揃えるため、`flake.nix` で再現可能な devShell を提供する。

```nix
# flake.nix（抜粋）
devShells.default = pkgs.mkShell {
  packages = [ pkgs.nodejs_22 pkgs.nodePackages.pnpm pkgs.git pkgs.python3 ];
};
```

使い方：
```bash
nix develop          # 一回だけ
direnv allow         # .envrc に `use flake` を仕込んでおけば自動で nix develop
```

`nix flake check` は `pnpm typecheck` + `pnpm test` を実行する軽量ゲート。重い build / lint は CI で実行する。

## 1. パッケージ構成

`ooxml-js` 単一パッケージ、サブパス export で機能分割（[01-architecture.md](./01-architecture.md) §6）。

### 1.1 package.json（最小骨子）

```jsonc
{
  "name": "ooxml-js",
  "version": "0.0.0",
  "type": "module",
  "license": "MIT",
  "sideEffects": false,
  "engines": { "node": ">=18.18" },
  "files": ["dist/**", "README.md", "THIRD_PARTY_NOTICES.md", "LICENSE"],
  "exports": {
    ".":               { "types": "./dist/index.d.ts", "node": { "import": "./dist/index.node.mjs", "require": "./dist/index.node.cjs" }, "browser": "./dist/index.browser.mjs", "default": "./dist/index.browser.mjs" },
    "./read":          { "types": "./dist/read.d.ts",  "node": { "import": "./dist/read.node.mjs",  "require": "./dist/read.node.cjs" },  "browser": "./dist/read.browser.mjs",  "default": "./dist/read.browser.mjs" },
    "./write":         { "types": "./dist/write.d.ts", "node": { "import": "./dist/write.node.mjs", "require": "./dist/write.node.cjs" }, "browser": "./dist/write.browser.mjs", "default": "./dist/write.browser.mjs" },
    "./streaming":     { "types": "./dist/streaming.d.ts", "node": { "import": "./dist/streaming.node.mjs", "require": "./dist/streaming.node.cjs" }, "browser": "./dist/streaming.browser.mjs" },
    "./styles":        { "types": "./dist/styles.d.ts", "import": "./dist/styles.mjs", "require": "./dist/styles.cjs" },
    "./conditional":   { "types": "./dist/conditional.d.ts", "import": "./dist/conditional.mjs", "require": "./dist/conditional.cjs" },
    "./formula":       { "types": "./dist/formula.d.ts", "import": "./dist/formula.mjs", "require": "./dist/formula.cjs" },
    "./drawing":       { "types": "./dist/drawing.d.ts", "import": "./dist/drawing.mjs", "require": "./dist/drawing.cjs" },
    "./chart":         { "types": "./dist/chart.d.ts", "import": "./dist/chart.mjs", "require": "./dist/chart.cjs" },
    "./chart/extended":{ "types": "./dist/chart-extended.d.ts", "import": "./dist/chart-extended.mjs", "require": "./dist/chart-extended.cjs" },
    "./pivot":         { "types": "./dist/pivot.d.ts", "import": "./dist/pivot.mjs", "require": "./dist/pivot.cjs" },
    "./schema":        { "types": "./dist/schema.d.ts", "import": "./dist/schema.mjs", "require": "./dist/schema.cjs" },
    "./io/node":       { "types": "./dist/io-node.d.ts", "import": "./dist/io-node.mjs", "require": "./dist/io-node.cjs" },
    "./io/browser":    { "types": "./dist/io-browser.d.ts", "import": "./dist/io-browser.mjs" }
  },
  "dependencies": {
    "fflate": "^0.8.2",
    "fast-xml-parser": "^4.4.0",
    "saxes": "^6.0.0"
  },
  "optionalDependencies": {
    "image-size": "^1.1.1"
  },
  "peerDependenciesMeta": {
    "image-size": { "optional": true }
  },
  "devDependencies": {
    "@types/node": "^22",
    "oxlint": "^1",
    "@vitest/browser": "^2",
    "@vitest/coverage-v8": "^2",
    "fast-check": "^3",
    "playwright": "^1",
    "size-limit": "^11",
    "tsdown": "^0.21",
    "typedoc": "^0.26",
    "typescript": "^5.4",
    "vitest": "^2",
    "@changesets/cli": "^2"
  }
}
```

### 1.2 tsconfig

`tsconfig.json`（共通）:
```jsonc
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "ESNext",
    "moduleResolution": "bundler",
    "strict": true,
    "noUncheckedIndexedAccess": true,
    "exactOptionalPropertyTypes": true,
    "noImplicitOverride": true,
    "noPropertyAccessFromIndexSignature": true,
    "isolatedModules": true,
    "verbatimModuleSyntax": true,
    "esModuleInterop": false,
    "allowSyntheticDefaultImports": false,
    "skipLibCheck": true,
    "useDefineForClassFields": true,
    "noEmit": true
  },
  "include": ["src", "tests", "scripts"]
}
```

`tsconfig.build.json`:
```jsonc
{
  "extends": "./tsconfig.json",
  "compilerOptions": { "noEmit": false, "declaration": true, "emitDeclarationOnly": true, "outDir": "dist", "rootDir": "src" },
  "include": ["src"]
}
```

### 1.3 ビルド: tsdown

`tsdown.config.ts`:
```ts
import { defineConfig } from 'tsdown';

const entries = [
  { name: 'index',          file: 'src/index.ts',                      envs: ['node', 'browser'] },
  { name: 'read',           file: 'src/public/read.ts',                envs: ['node', 'browser'] },
  { name: 'write',          file: 'src/public/write.ts',               envs: ['node', 'browser'] },
  { name: 'streaming',      file: 'src/streaming/index.ts',            envs: ['node', 'browser'] },
  { name: 'styles',         file: 'src/styles/index.ts',               envs: [null] },
  { name: 'conditional',    file: 'src/formatting/index.ts',           envs: [null] },
  { name: 'formula',        file: 'src/formula/index.ts',              envs: [null] },
  { name: 'drawing',        file: 'src/drawing/index.ts',              envs: [null] },
  { name: 'chart',          file: 'src/chart/index.ts',                envs: [null] },
  { name: 'chart-extended', file: 'src/chart/cx/index.ts',             envs: [null] },
  { name: 'pivot',          file: 'src/pivot/index.ts',                envs: [null] },
  { name: 'schema',         file: 'src/schema/index.ts',               envs: [null] },
  { name: 'io-node',        file: 'src/io/node.ts',                    envs: ['node'] },
  { name: 'io-browser',     file: 'src/io/browser.ts',                 envs: ['browser'] },
];

export default defineConfig(
  entries.flatMap(({ name, file, envs }) =>
    envs.map((env) => ({
      entry: { [`${name}${env ? `.${env}` : ''}`]: file },
      format: env === 'node' ? ['esm', 'cjs'] : env === 'browser' ? ['esm'] : ['esm', 'cjs'],
      target: env === 'node' ? 'node18' : 'es2022',
      platform: env === 'browser' ? 'browser' : env === 'node' ? 'node' : 'neutral',
      dts: false,                  // d.ts は別途 tsc で吐く
      minify: false,
      sourcemap: true,
      clean: false,
      treeshake: true,
      external: ['node:fs', 'node:fs/promises', 'node:path', 'node:crypto', 'node:stream', 'image-size'],
    }))
  )
);
```

`d.ts` は `tsc -p tsconfig.build.json` で生成し、entrypoint に対応する `.d.ts` を `dist/` 直下に配置する別スクリプトで整える。

### 1.4 size-limit

```jsonc
[
  { "name": "index (browser, gz)",     "path": "dist/index.browser.mjs",      "limit": "200 KB", "gzip": true },
  { "name": "read (browser, gz)",      "path": "dist/read.browser.mjs",       "limit": "120 KB", "gzip": true },
  { "name": "write (browser, gz)",     "path": "dist/write.browser.mjs",      "limit": "150 KB", "gzip": true },
  { "name": "streaming (browser, gz)", "path": "dist/streaming.browser.mjs",  "limit": "80 KB",  "gzip": true },
  { "name": "chart (gz)",              "path": "dist/chart.mjs",              "limit": "120 KB", "gzip": true },
  { "name": "chart-extended (gz)",     "path": "dist/chart-extended.mjs",     "limit": "60 KB",  "gzip": true },
  { "name": "drawing (gz)",            "path": "dist/drawing.mjs",            "limit": "40 KB",  "gzip": true },
  { "name": "styles (gz)",             "path": "dist/styles.mjs",             "limit": "20 KB",  "gzip": true }
]
```

CI で size-limit を必須通過とする。

## 2. リント

`oxlint`（oxc 系）で一本化。Rust 実装で 10× 速いほか、設定ファイルが ESLint の rule 名と互換 (`eslint(...)` / `eslint-plugin-unicorn(...)` / `typescript-eslint(...)` / `oxc(...)`)。

`.oxlintrc.json`（最大強度・厳格設定）:
```jsonc
{
  "categories": {
    "correctness": "error",
    "suspicious": "error",
    "perf": "error",
    "style": "error",
    "pedantic": "error",
    "restriction": "off",
    "nursery": "off"
  },
  "rules": {
    // 既存コードに合わせたゆるめのオーバーライド (代表例):
    "no-underscore-dangle": "off",        // `_xxxByKey` 内部 dedup map
    "max-statements": "off",
    "capitalized-comments": "off",
    "typescript/no-non-null-assertion": "warn",
    "unicorn/no-array-sort": "off",       // ローカル配列の sort は OK
    "unicorn/prefer-set-has": "off",
    "new-cap": "off"
  },
  "ignorePatterns": ["dist", "reference", "node_modules", "coverage"]
}
```

`pnpm lint` は `oxlint`、`pnpm lint:fix` は `oxlint --fix`。フォーマッタは現時点で oxc 系には専用ツールがないので `tsc` の strict 設定 + `oxlint --fix` で許容範囲を保つ。

加えて、**`class` キーワード使用禁止のカスタムルール**は将来 oxlint プラグインまたは codereview チェックで実装。例外：`Error` 派生のみ許可。

## 3. CI

GitHub Actions（[10-testing.md](./10-testing.md) §7 と統合）：

```yaml
# .github/workflows/ci.yml
name: CI
on: [push, pull_request]
jobs:
  build:
    runs-on: ubuntu-latest
    strategy:
      matrix: { node-version: [18, 20, 22] }
    steps:
      - uses: actions/checkout@v4
        with: { submodules: recursive }   # reference/openpyxl を取得
      - uses: pnpm/action-setup@v3
      - uses: actions/setup-node@v4
        with: { node-version: ${{ matrix.node-version }}, cache: 'pnpm' }
      - run: pnpm install --frozen-lockfile
      - run: pnpm typecheck
      - run: pnpm lint
      - run: pnpm build
      - run: pnpm test --coverage
      - run: pnpm size
      - if: matrix.node-version == 22
        run: pnpm exec playwright install --with-deps chromium firefox webkit
      - if: matrix.node-version == 22
        run: pnpm test:browser
```

## 4. リリースフロー

`@changesets/cli` で運用：
1. PR で `pnpm exec changeset` を実行 → changeset md を作成
2. main へ merge → release-please / changesets/action が `release/x.y.z` PR を作成
3. その PR を merge → npm publish + GitHub Release

semver 方針：
- フェーズ完了で minor を上げる
- 公開 API の breaking 変更は major
- フェーズ7 完了をもって 1.0.0 リリース候補

## 5. ドキュメント

- README.md: 簡潔な使用例 + サブパス一覧 + ライセンス
- docs/plan/: 本計画ドキュメント（本リポジトリ管理者向け）
- docs/api/: typedoc 自動生成（`pnpm doc:api`）
- docs/guide/: ユーザ向けガイド（フェーズ3 完了後に着手）
- docs/migrate-from-openpyxl.md: openpyxl 利用者の移行ガイド（フェーズ7 完了後）

公開先：
- API ドキュメント: GitHub Pages (`docs/api/` を `gh-pages` に publish)
- 利用ガイド: 同 GitHub Pages の `/guide/`

## 6. THIRD_PARTY_NOTICES

`THIRD_PARTY_NOTICES.md` に以下を記載：
- openpyxl の MIT ライセンス文
- fflate / fast-xml-parser / saxes / image-size の各ライセンス
- ECMA-376 の参照について

`scripts/regenerate-notices.ts` で自動生成可能にしておく。
