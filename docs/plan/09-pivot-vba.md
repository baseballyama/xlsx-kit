# 09. フェーズ7: ピボット / VBA / 暗号化（passthrough 主体）

**目的**: openpyxl と同等の「壊さずに通す」レベルのサポートを最終化する。ピボットの構造編集 API は提供しない。
**期間目安**: 2〜3週間
**前提**: フェーズ1〜6
**完了条件**: ピボット / VBA / ActiveX / customUI / OLE / カスタム XML パーツを含む xlsx の round-trip。

## 1. 方針

このフェーズの主目的は **「openpyxl が壊さない xlsx は ooxml-js も壊さない」** こと。新しい構造編集 API を提供するのは目的ではない。

passthrough と construction は明確に区別する：

| サブシステム | passthrough（フェーズ7） | construction API（将来） |
|------------|-----------------------|------------------------|
| Pivot Table | ✅ | ❌（将来検討） |
| Pivot Cache | ✅ | ❌（将来検討） |
| VBA Project | ✅（バイナリ温存） | ❌（範囲外） |
| ActiveX | ✅ | ❌ |
| Form Controls | ✅ | ❌ |
| OLE Objects | ✅ | ❌ |
| Custom UI ribbon | ✅ | ❌ |
| Custom XML Parts | ✅ | △（メタのみ） |

## 2. ピボット（`src/pivot/`）

### 2.1 read

参照: openpyxl `pivot/cache.py`, `pivot/record.py`, `pivot/table.py`。

read 時にやること：
- `xl/pivotCache/pivotCacheDefinitionN.xml` を **schema 経由で plain object に**
- `xl/pivotCache/pivotCacheRecordsN.xml` を **バイナリのまま** `wb.passthrough.set('xl/pivotCache/pivotCacheRecordsN.xml', bytes)`
- `xl/pivotTables/pivotTableN.xml` を **schema 経由** で plain object に
- worksheet → pivot rels も保持

### 2.2 write

worksheet 側の `<pivotTables>` 参照、workbook 側の `<pivotCaches>` 参照を **すべて再生成**。pivot definition は plain object → schema → XML へ。pivot records は **バイナリをそのまま** zip に書き戻す。

### 2.3 公開 API

```ts
export interface PivotTablePassthrough {
  id: number;
  name: string;
  cacheId: number;
  /** 構造編集はしない。読み取り情報を提供する程度 */
  rowFields: string[];
  colFields: string[];
  dataFields: string[];
  /** raw XML */
  rawXml: Uint8Array;
}

export function listPivotTables(wb: Workbook): PivotTablePassthrough[];
```

### 2.4 受け入れ条件

- [ ] openpyxl `pivot/tests/data/` の round-trip
- [ ] 大規模 pivot（cache 100k records）でも save 後に Excel で再 open できる
- [ ] pivot が壊れたら `OpenXmlInvalidWorkbookError` で具体的にどこが壊れているか報告

## 3. VBA / ActiveX / OLE

### 3.1 read

`load_workbook(source, { keepVba: true })` でバイナリを **完全保持**：

| ZIP path | 保存先 |
|----------|--------|
| `xl/vbaProject.bin` | `wb.vbaProject` |
| `xl/vbaProjectSignature.bin` | `wb.vbaSignature` |
| `xl/activeX/*` | `wb.passthrough.set(path, bytes)` |
| `xl/embeddings/*` | `wb.passthrough` |
| `xl/ctrlProps/*` | `wb.passthrough` |
| `customUI/*` | `wb.passthrough` |
| `xl/drawings/*.vml` (controls) | `wb.passthrough` |

### 3.2 write

manifest と rels を整合させて再書き出し。VBA を含む xlsx は **`.xlsm`** content type を `app/vnd.ms-excel.sheet.macroEnabled.main+xml` に変える。

### 3.3 受け入れ条件

- [ ] VBA を含む xlsm の round-trip で **vbaProject.bin が byte-identical**
- [ ] ActiveX コントロール付き sheet の round-trip
- [ ] customUI ribbon が消えない
- [ ] keepVba: false で xlsm を読むと VBA が消えて xlsx になる

## 4. 暗号化（明示的に未対応）

EncryptedDocument ストリームベースの暗号化は **明示的に未対応**：

- read 時に検出 → `OpenXmlNotImplementedError('Encrypted xlsx is not supported. Decrypt with msoffcrypto-tool first.')`
- write は元から無理（暗号化 API なし）

検出方法：ZIP の中身が `EncryptedPackage` / `EncryptionInfo` の OLE Compound Document であれば暗号化。`fflate` で開く前に **マジックバイト** を確認：

```ts
function detectFormat(bytes: Uint8Array): 'xlsx' | 'cfb-encrypted' | 'xls' | 'unknown' {
  // ZIP header: 50 4B 03 04
  // CFB header: D0 CF 11 E0 A1 B1 1A E1
  // …
}
```

## 5. カスタム XML パーツ

`customXml/itemN.xml` 等：そのまま passthrough。一覧 API のみ提供：

```ts
export function listCustomXmlParts(wb: Workbook): Array<{ path: string; content: Uint8Array }>;
```

## 6. content type / rels の自動補正

passthrough バイナリがあるとき、manifest の override がきちんと出るように **content type 推論テーブル** を持つ：

```ts
const KNOWN_CONTENT_TYPES: Record<string, string> = {
  'bin':  'application/vnd.ms-office.vbaProject',
  'vml':  'application/vnd.openxmlformats-officedocument.vmlDrawing',
  'emf':  'image/x-emf',
  'wmf':  'image/x-wmf',
  // …
};
```

## 7. 完了条件

- [ ] §2〜§5 の受け入れ条件
- [ ] openpyxl の `tests/data/reader/vba+comments.xlsm` のような重め fixture で round-trip
- [ ] Excel 365 で開いた時にマクロ署名警告が壊れない
