// Tests for renameDefinedName.

import { describe, expect, it } from 'vitest';
import {
  addDefinedName,
  getDefinedName,
  renameDefinedName,
} from '../../src/xlsx/workbook/defined-names';
import { createWorkbook } from '../../src/xlsx/workbook/workbook';

describe('renameDefinedName', () => {
  it('renames a workbook-scope name', () => {
    const wb = createWorkbook();
    addDefinedName(wb, { name: 'Old', value: '$A$1' });
    expect(renameDefinedName(wb, 'Old', 'New')).toBe(true);
    expect(getDefinedName(wb, 'Old')).toBeUndefined();
    expect(getDefinedName(wb, 'New')?.value).toBe('$A$1');
  });

  it('renames within a specific sheet scope', () => {
    const wb = createWorkbook();
    addDefinedName(wb, { name: 'Foo', value: '$A$1', scope: 0 });
    addDefinedName(wb, { name: 'Foo', value: '$A$1', scope: 1 });
    renameDefinedName(wb, 'Foo', 'Bar', 0);
    expect(getDefinedName(wb, 'Foo', 0)).toBeUndefined();
    expect(getDefinedName(wb, 'Bar', 0)?.value).toBe('$A$1');
    // Other scope still has Foo.
    expect(getDefinedName(wb, 'Foo', 1)?.value).toBe('$A$1');
  });

  it('returns false when nothing matches', () => {
    const wb = createWorkbook();
    expect(renameDefinedName(wb, 'Missing', 'New')).toBe(false);
  });

  it('throws when newName collides with another entry at the same scope', () => {
    const wb = createWorkbook();
    addDefinedName(wb, { name: 'A', value: '$A$1' });
    addDefinedName(wb, { name: 'B', value: '$B$1' });
    expect(() => renameDefinedName(wb, 'A', 'B')).toThrow(/already in use/);
  });

  it('different-scope same-name is OK (no collision)', () => {
    const wb = createWorkbook();
    addDefinedName(wb, { name: 'A', value: '$A$1' });
    addDefinedName(wb, { name: 'A', value: '$A$1', scope: 0 });
    // Rename the workbook-scope entry; sheet-scope 'A' is unaffected.
    expect(renameDefinedName(wb, 'A', 'B')).toBe(true);
    expect(getDefinedName(wb, 'B')).toBeDefined();
    expect(getDefinedName(wb, 'A', 0)).toBeDefined();
  });
});
