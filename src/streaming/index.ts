// xlsx streaming entry point — read-only iter + write-only append.
// Format-agnostic byte I/O lives at `xlsx-kit/io` and `xlsx-kit/node`;
// error types at `xlsx-kit/utils`.

export {
  loadWorkbookStream,
  type IterRowsOptions,
  type LoadWorkbookStreamOptions,
  type ReadOnlyCell,
  type ReadOnlyWorkbook,
  type ReadOnlyWorksheet,
} from './read-only';

export {
  createWriteOnlyWorkbook,
  type WriteOnlyOptions,
  type WriteOnlyRowItem,
  type WriteOnlyStyle,
  type WriteOnlyWorkbook,
  type WriteOnlyWorksheet,
} from './write-only';
