// xlsx streaming entry point — read-only iter + write-only append.
// Format-agnostic byte I/O lives at `xlsxlite/io` and `xlsxlite/node`;
// error types at `xlsxlite/utils`.

export {
  loadWorkbookStream,
  type IterRowsOptions,
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
