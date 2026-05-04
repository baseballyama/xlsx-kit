// Public entry for openxml-js.
//
// Phase 0 (bootstrap): empty surface. Real exports land in phase 1+ per
// docs/plan/01-architecture.md §3 and 11-build-publish.md §1.1.
//
// The eventual public surface will be additive `export *` re-exports from
// internal subpackages — never default exports — to keep tree-shaking sound.

export {};
