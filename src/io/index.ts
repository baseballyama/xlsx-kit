// Environment-neutral re-exports. The Node-only and browser-only helper
// surfaces live in node.ts and browser.ts respectively and are reached via
// `openxml-js/io/node` / `openxml-js/io/browser` subpaths once the
// package.json exports map gains those conditions (planned in
// docs/plan/11-build-publish.md §1.1).

export type { BufferedSinkWriter, XlsxSink } from './sink';
export type { XlsxSource } from './source';
