// Tests for chartsheet-level <webPublishItems>.

import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import { loadWorkbook } from '../../src/io/load';
import { workbookToBytes } from '../../src/io/save';
import { addChartsheet, createWorkbook } from '../../src/workbook/workbook';
import { makeWebPublishItem } from '../../src/worksheet/web-publish';

describe('chartsheet webPublishItems round-trip', () => {
  it('preserves a chartsheet entry', async () => {
    const wb = createWorkbook();
    addChartsheet(wb, 'Chart1');
    const ref = wb.sheets[0];
    if (ref?.kind !== 'chartsheet') throw new Error('expected chartsheet');
    ref.sheet.webPublishItems.push(
      makeWebPublishItem({
        id: 1,
        divId: 'wpChart',
        sourceType: 'chart',
        destinationFile: 'http://example.com/chart.html',
        title: 'My published chart',
        autoRepublish: true,
      }),
    );

    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ref2 = wb2.sheets[0];
    if (ref2?.kind !== 'chartsheet') throw new Error('expected chartsheet on reload');
    expect(ref2.sheet.webPublishItems.length).toBe(1);
    const item = ref2.sheet.webPublishItems[0];
    expect(item?.id).toBe(1);
    expect(item?.divId).toBe('wpChart');
    expect(item?.sourceType).toBe('chart');
    expect(item?.title).toBe('My published chart');
    expect(item?.autoRepublish).toBe(true);
  });

  it('emits no <webPublishItems/> when empty', async () => {
    const wb = createWorkbook();
    addChartsheet(wb, 'Chart1');
    const bytes = await workbookToBytes(wb);
    const wb2 = await loadWorkbook(fromBuffer(bytes));
    const ref = wb2.sheets[0];
    if (ref?.kind !== 'chartsheet') throw new Error('expected chartsheet');
    expect(ref.sheet.webPublishItems.length).toBe(0);
  });
});