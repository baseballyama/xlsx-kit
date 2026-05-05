// Tests for the typed workbook-level smart-tag and function-group
// models.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/public/load';
import { workbookToBytes } from '../../src/public/save';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import { makeFunctionGroups } from '../../src/workbook/function-groups';
import { makeSmartTagProperties, makeSmartTagType } from '../../src/workbook/smart-tags';

describe('smartTagPr round-trip', () => {
  it('preserves embed + show', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    wb.smartTagPr = makeSmartTagProperties({ embed: true, show: 'noIndicator' });
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.smartTagPr?.embed).toBe(true);
    expect(wb2.smartTagPr?.show).toBe('noIndicator');
  });

  it('drops unknown show enum values', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    wb.smartTagPr = makeSmartTagProperties({ embed: false, show: 'gibberish' as never });
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.smartTagPr?.embed).toBe(false);
    expect(wb2.smartTagPr?.show).toBeUndefined();
  });
});

describe('smartTagTypes round-trip', () => {
  it('preserves a list of three smart-tag type registrations', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    wb.smartTagTypes = [
      makeSmartTagType({ namespaceUri: 'urn:schemas-microsoft-com:office:smarttags', name: 'PersonName', url: 'https://example.com/name' }),
      makeSmartTagType({ namespaceUri: 'urn:schemas-microsoft-com:office:smarttags', name: 'date' }),
      makeSmartTagType({ name: 'phone' }),
    ];
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.smartTagTypes?.length).toBe(3);
    expect(wb2.smartTagTypes?.[0]?.name).toBe('PersonName');
    expect(wb2.smartTagTypes?.[0]?.url).toBe('https://example.com/name');
    expect(wb2.smartTagTypes?.[2]?.namespaceUri).toBeUndefined();
    expect(wb2.smartTagTypes?.[2]?.name).toBe('phone');
  });
});

describe('functionGroups round-trip', () => {
  it('preserves builtInGroupCount + custom user-defined groups', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    wb.functionGroups = makeFunctionGroups({
      builtInGroupCount: 16,
      groups: [{ name: 'MyXLLGroup' }, { name: 'OtherGroup' }],
    });
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.functionGroups?.builtInGroupCount).toBe(16);
    expect(wb2.functionGroups?.groups.length).toBe(2);
    expect(wb2.functionGroups?.groups[0]?.name).toBe('MyXLLGroup');
    expect(wb2.functionGroups?.groups[1]?.name).toBe('OtherGroup');
  });

  it('emits no <functionGroups/> when undefined', async () => {
    const wb = createWorkbook();
    addWorksheet(wb, 'A');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    expect(wb2.functionGroups).toBeUndefined();
  });
});