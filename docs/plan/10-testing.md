# 10. テスト戦略

## 1. テストランナー / 環境

| 環境 | コマンド | 目的 |
|------|--------|------|
| Node 18+（コア） | `pnpm test` | 全単体・結合テスト |
| jsdom | `pnpm test:jsdom` | ブラウザ依存 API のスモーク |
| Chromium / Firefox / WebKit | `pnpm test:browser` | 真のブラウザでのコアスモーク |
| Bench | `pnpm bench` | 性能リグレッション（vitest bench） |
| サイズ | `pnpm size` | size-limit の予算守護 |
| Property | `pnpm test --run --reporter dot` のうち `*.property.ts` | fast-check ファジング |

## 2. ディレクトリ構成

```
tests/
├── phase-1/                    # 基盤層
├── phase-2/                    # コアモデル
├── phase-3/                    # read/write
├── phase-4/                    # streaming
├── phase-5/                    # rich features
├── phase-6/                    # charts/drawings
├── phase-7/                    # pivot/vba
├── public/                     # 公開 API レベル e2e
├── fixtures/
│   ├── genuine -> ../reference/openpyxl/openpyxl/tests/data/genuine
│   ├── reader -> ...
│   ├── writer -> ...
│   └── xlsx-craft/             # 自前フィクスチャ（生成スクリプト付き）
├── golden/                     # canonical XML / SHA256 ハッシュ
├── browser/                    # ブラウザ専用テスト
├── perf/                       # ベンチマーク
└── helpers/
    ├── canonicalize.ts         # XML 正規化
    ├── compare-xml.ts          # XML 比較
    ├── load-fixture.ts
    └── make-temp-sink.ts
```

## 3. canonical XML 比較

openpyxl の `tests/helper.py:compare_xml` を移植する。

```ts
// tests/helpers/canonicalize.ts
export function canonicalize(xml: Uint8Array | string): string {
  const node = parseXml(typeof xml === 'string' ? new TextEncoder().encode(xml) : xml);
  return canonicalizeNode(node);
}

function canonicalizeNode(n: XmlNode): string {
  // 1. 名前空間 prefix を統一（e.g. r → rId, xdr → xdr）
  // 2. 属性をキー昇順で sort
  // 3. text の先頭末尾空白を除去（preserve-space を尊重）
  // 4. 子要素を順番に再帰
  // 5. 出力は <name attr1="..." attr2="..."><child/>...</name>
}

export function compareXml(a: string | Uint8Array, b: string | Uint8Array): { equal: boolean; diff?: string };
```

`compareXml` は `expect(compareXml(actual, expected).equal).toBe(true)` で使う。差分は diff として返し、テスト出力に表示。

## 4. テストの種類

### 4.1 schema round-trip（descriptor 相当）

各 schema に対して以下を必ず書く：

```ts
import { fromTree, toTree } from 'xlsx-craft/schema';
import { FontSchema } from 'src/styles/fonts.schema.js';

describe('Font schema', () => {
  const cases: Array<{ name: string; xml: string; expected: Font }> = [
    { name: 'minimal', xml: '<font/>', expected: {} },
    { name: 'bold + size', xml: '<font><sz val="11"/><b/></font>', expected: { size: 11, bold: true } },
    // …
  ];

  it.each(cases)('round-trips $name', ({ xml, expected }) => {
    const node = parseXml(new TextEncoder().encode(xml));
    const value = fromTree(node, FontSchema);
    expect(value).toEqual(expected);

    const back = toTree(value, FontSchema);
    const written = serializeXml(back);
    expect(compareXml(written, xml).equal).toBe(true);
  });
});
```

### 4.2 ファイル単位 round-trip

```ts
const FIXTURES = ['genuine/empty.xlsx', 'genuine/empty-with-styles.xlsx', 'genuine/sample.xlsx'];

describe.each(FIXTURES)('round-trip: %s', (fixture) => {
  it('XML parts canonicalize equivalently', async () => {
    const orig = await openZipFromFixture(fixture);
    const wb = await loadWorkbook(fromBytes(await orig.fullBytes()));
    const out = toBuffer();
    await saveWorkbook(wb, out);
    const written = await openZipFromBytes(out.result());

    for (const path of XML_PARTS_TO_COMPARE) {
      const a = canonicalize(orig.read(path));
      const b = canonicalize(written.read(path));
      expect(b).toBe(a);
    }
  });
});
```

`XML_PARTS_TO_COMPARE` はフェーズごとに増やす（フェーズ3 では `xl/workbook.xml`, `xl/styles.xml`, sheet1 まで）。

### 4.3 golden hash

```ts
import { createHash } from 'node:crypto';
import golden from 'tests/golden/styles.json';

describe('styles hash stability', () => {
  it('Font(bold=true)', () => {
    const xml = serializeXml(toTree(makeFont({ bold: true }), FontSchema));
    const hash = createHash('sha256').update(canonicalize(xml)).digest('hex').slice(0, 16);
    expect(hash).toBe(golden['Font.bold']);
  });
});
```

意図的に変更したら golden を更新する CLI を用意：
```
pnpm test:update-golden
```

### 4.4 property-based（fast-check）

descriptor の round-trip：
```ts
import * as fc from 'fast-check';

it('Border round-trip', () => {
  const sideArbitrary = fc.record({
    style: fc.constantFrom('thin', 'medium', 'thick', 'double', 'hair', 'dotted', 'dashed'),
    color: fc.option(fc.record({ rgb: fc.hexaString({ minLength: 8, maxLength: 8 }) })),
  });
  const borderArbitrary = fc.record({
    left: fc.option(sideArbitrary),
    right: fc.option(sideArbitrary),
    top: fc.option(sideArbitrary),
    bottom: fc.option(sideArbitrary),
  });
  fc.assert(fc.property(borderArbitrary, (border) => {
    const xml = serializeXml(toTree(border, BorderSchema));
    const round = fromTree(parseXml(new TextEncoder().encode(xml)), BorderSchema);
    expect(round).toEqual(border);
  }), { numRuns: 200 });
});
```

ホットパスにもファジング：
- coordinate 文字列パーサ vs エンコーダの対称性
- formula tokenizer の token 列を再連結して元の式と一致するか（whitespace 完全一致は除く）

### 4.5 性能ベンチ（vitest bench）

```ts
import { bench } from 'vitest';

describe('worksheet write', () => {
  bench('100k cells row-major number', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'S');
    for (let r = 1; r <= 1000; r++) {
      for (let c = 1; c <= 100; c++) setCell(ws, r, c, r * c);
    }
    const sink = toBuffer();
    await saveWorkbook(wb, sink);
  }, { time: 1000, iterations: 5 });
});
```

CI で前 commit と比較し、25% 劣化で fail。

## 5. ブラウザテスト

`@vitest/browser` + Playwright で **3 ブラウザ全部** で `tests/browser/` を回す：

```ts
import { describe, it, expect } from 'vitest';
import { loadWorkbook, fromBlob, toBlob } from 'xlsx-craft';

describe('browser: round-trip from Blob', () => {
  it('reads Blob and writes Blob', async () => {
    const fixtureUrl = '/fixtures/empty-with-styles.xlsx';
    const blob = await fetch(fixtureUrl).then((r) => r.blob());
    const wb = await loadWorkbook(fromBlob(blob));
    const sink = toBlob('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    await saveWorkbook(wb, sink);
    const result = sink.result();
    expect(result.size).toBeGreaterThan(0);
    // round trip 再 read
    const wb2 = await loadWorkbook(fromBlob(result));
    expect(wb2.sheets.length).toBe(wb.sheets.length);
  });
});
```

ブラウザ用 fixture は `tests/browser/public/fixtures/` に置き、Vite dev server で配信する。

## 6. 互換性検証 QA

PR の `release/*` ブランチに対しては手動 QA：
1. 自動的に出力された xlsx 群を artifact として CI からダウンロード
2. **Excel 365 / LibreOffice / Google Sheets** で開く
3. 警告ダイアログ・データ欠損があれば issue 化

QA 用 xlsx は最低限以下を含む：
- 全 chart 種類サンプル
- pivot サンプル
- 大規模シート（10 万 cell）
- 条件付き書式 + dxf
- VBA 付き xlsm

## 7. CI ワークフロー

```yaml
# .github/workflows/ci.yml の概形
jobs:
  typecheck:
    steps:
      - run: pnpm typecheck
  lint:
    steps:
      - run: pnpm lint
  test:
    strategy:
      matrix: { node-version: [18, 20, 22] }
    steps:
      - run: pnpm test --coverage
  test-browser:
    steps:
      - run: pnpm exec playwright install --with-deps
      - run: pnpm test:browser
  bench:
    if: github.event_name == 'pull_request'
    steps:
      - run: pnpm bench:smoke
  size:
    steps:
      - run: pnpm size
  qa-fixtures:
    if: github.event.pull_request.head.label.endsWith('/release')
    steps:
      - run: pnpm qa:generate
      - uses: actions/upload-artifact@v4
        with:
          name: qa-xlsx
          path: tests/qa/dist
```

## 8. テスト品質指標

- **行カバレッジ ≥ 90%**（フェーズ3 完了時に到達目標）
- **分岐カバレッジ ≥ 80%**
- **descriptor schema 単位の必須テスト**: 各 schema に最低 3 ケース（minimal / typical / edge）
- **fixture round-trip**: openpyxl `tests/data/genuine/` の全件
- **bench regression**: 25% 以下

## 9. 将来の拡張

- xlsx-validator 相当の **自前 xlsx 妥当性チェッカー**（フェーズ8 以降）
  - manifest / rels / sheet 数 / dimension 整合性 を検査
- xsd validation: **採用しない**（openpyxl precedent と同じ理由）
- mutation testing（stryker）: 採用検討（フェーズ完了後の品質強化として）
