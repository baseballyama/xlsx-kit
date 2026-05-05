// Scenario 11: legacy threaded-style comments.
// Output: 11-comments.xlsx
//
// What to verify in Excel:
// - A red triangle in the corner of A1 / B2 / C3 indicates a comment.
// - Hovering each shows the author + body Excel renders in a yellow
//   tooltip box.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook, setCell } from '../../../src/index';
import { writeWorkbook } from '../_helpers';

describe('e2e 11 — comments', () => {
  it('writes 11-comments.xlsx', async () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'Notes');

    setCell(ws, 1, 1, 'A1 has a comment');
    setCell(ws, 2, 2, 'B2 has a comment');
    setCell(ws, 3, 3, 'C3 has a comment');

    ws.legacyComments.push({
      ref: 'A1',
      author: 'Alice',
      text: 'Alice:\nFirst comment in A1',
    });
    ws.legacyComments.push({
      ref: 'B2',
      author: 'Bob',
      text: 'Bob:\nReminder about B2 — needs review',
    });
    ws.legacyComments.push({
      ref: 'C3',
      author: 'Carol',
      text: 'Carol:\nLong-running issue, see ticket #123 for context',
    });

    const result = await writeWorkbook('11-comments.xlsx', wb);
    expect(result.bytes).toBeGreaterThan(0);
  });
});
