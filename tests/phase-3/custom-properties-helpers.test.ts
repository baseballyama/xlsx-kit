// Tests for docProps/custom.xml ergonomic property setters.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import {
  getCustomPropertyValue,
  listCustomProperties,
  removeCustomProperty,
  setCustomBoolProperty,
  setCustomDateProperty,
  setCustomNumberProperty,
  setCustomStringProperty,
} from '../../src/packaging/custom';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';

describe('setCustomXxxProperty', () => {
  it('lazily allocates customProperties and appends each kind', () => {
    const wb = createWorkbook();
    expect(wb.customProperties).toBeUndefined();
    setCustomStringProperty(wb, 'project', 'Apollo');
    setCustomNumberProperty(wb, 'priority', 3);
    setCustomNumberProperty(wb, 'budget', 1234.56);
    setCustomBoolProperty(wb, 'archived', false);
    setCustomDateProperty(wb, 'created', new Date('2024-01-15T00:00:00Z'));
    expect(listCustomProperties(wb).length).toBe(5);
    expect(getCustomPropertyValue(wb, 'project')).toBe('Apollo');
    expect(getCustomPropertyValue(wb, 'priority')).toBe(3);
    expect(getCustomPropertyValue(wb, 'budget')).toBeCloseTo(1234.56);
    expect(getCustomPropertyValue(wb, 'archived')).toBe(false);
    expect(getCustomPropertyValue(wb, 'created')).toBe('2024-01-15T00:00:00.000Z');
  });

  it('numeric int picks vt:i4; non-int picks vt:r8', () => {
    const wb = createWorkbook();
    const intProp = setCustomNumberProperty(wb, 'count', 42);
    const fltProp = setCustomNumberProperty(wb, 'rate', 3.14);
    expect(intProp.value.name).toContain('i4');
    expect(fltProp.value.name).toContain('r8');
  });

  it('replacing an existing property by name preserves the pid', () => {
    const wb = createWorkbook();
    const a = setCustomStringProperty(wb, 'project', 'old');
    const b = setCustomStringProperty(wb, 'project', 'new');
    expect(a.pid).toBe(b.pid);
    expect(getCustomPropertyValue(wb, 'project')).toBe('new');
    expect(listCustomProperties(wb).length).toBe(1);
  });

  it('removeCustomProperty drops the entry and returns the right boolean', () => {
    const wb = createWorkbook();
    setCustomStringProperty(wb, 'a', 'x');
    setCustomStringProperty(wb, 'b', 'y');
    expect(removeCustomProperty(wb, 'a')).toBe(true);
    expect(removeCustomProperty(wb, 'a')).toBe(false);
    expect(listCustomProperties(wb).map((p) => p.name)).toEqual(['b']);
  });

  it('rejects NaN / Infinity numeric values', () => {
    const wb = createWorkbook();
    expect(() => setCustomNumberProperty(wb, 'x', Number.NaN)).toThrow(/not finite/);
    expect(() => setCustomNumberProperty(wb, 'x', Number.POSITIVE_INFINITY)).toThrow(/not finite/);
  });
});

describe('custom-property round-trip', () => {
  it('all kinds survive saveWorkbook → loadWorkbook', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    setCustomStringProperty(wb, 'project', 'Apollo');
    setCustomNumberProperty(wb, 'priority', 3);
    setCustomBoolProperty(wb, 'archived', false);
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(getCustomPropertyValue(wb2, 'project')).toBe('Apollo');
    expect(getCustomPropertyValue(wb2, 'priority')).toBe(3);
    expect(getCustomPropertyValue(wb2, 'archived')).toBe(false);
  });
});
