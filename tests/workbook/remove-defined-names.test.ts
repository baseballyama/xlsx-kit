// Tests for removeDefinedNames(predicate).

import { describe, expect, it } from 'vitest';
import {
  addDefinedName,
  listDefinedNames,
  removeDefinedNames,
} from '../../src/workbook/defined-names';
import { createWorkbook } from '../../src/workbook/workbook';

describe('removeDefinedNames', () => {
  it('removes only entries matching the predicate and returns the count', () => {
    const wb = createWorkbook();
    addDefinedName(wb, { name: 'Print_Area', value: '$A$1', scope: 0 });
    addDefinedName(wb, { name: 'Print_Area', value: '$A$1', scope: 1 });
    addDefinedName(wb, { name: 'KeepMe', value: '$A$1' });
    expect(removeDefinedNames(wb, (d) => d.name === 'Print_Area')).toBe(2);
    expect(listDefinedNames(wb).map((d) => d.name)).toEqual(['KeepMe']);
  });

  it('returns 0 when nothing matches', () => {
    const wb = createWorkbook();
    addDefinedName(wb, { name: 'X', value: '$A$1' });
    expect(removeDefinedNames(wb, (d) => d.name === 'Y')).toBe(0);
    expect(listDefinedNames(wb).length).toBe(1);
  });

  it('predicate sees scope so callers can filter sheet-scope names', () => {
    const wb = createWorkbook();
    addDefinedName(wb, { name: 'Wb', value: '$A$1' });
    addDefinedName(wb, { name: 'A', value: '$A$1', scope: 0 });
    addDefinedName(wb, { name: 'B', value: '$A$1', scope: 1 });
    expect(removeDefinedNames(wb, (d) => d.scope !== undefined)).toBe(2);
    expect(listDefinedNames(wb).map((d) => d.name)).toEqual(['Wb']);
  });
});
