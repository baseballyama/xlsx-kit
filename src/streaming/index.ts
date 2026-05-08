// xlsx streaming entry point — read-only iter + write-only append.
// Format-agnostic byte I/O lives at `xlsxify/io` and `xlsxify/node`;
// error types at `xlsxify/utils`.

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
