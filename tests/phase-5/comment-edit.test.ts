// Tests for editCommentText / editCommentAuthor in-place edit helpers.

import { describe, expect, it } from 'vitest';
import { addWorksheet, createWorkbook } from '../../src/workbook/workbook';
import {
  editCommentAuthor,
  editCommentText,
  getComment,
  setComment,
} from '../../src/worksheet/worksheet';

describe('editCommentText', () => {
  it('rewrites the text and leaves author + ref alone', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setComment(ws, { ref: 'A1', author: 'Alice', text: 'old' });
    expect(editCommentText(ws, 'A1', 'new')).toBe(true);
    const c = getComment(ws, 'A1');
    expect(c?.text).toBe('new');
    expect(c?.author).toBe('Alice');
    expect(c?.ref).toBe('A1');
  });

  it('returns false when the ref is missing', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(editCommentText(ws, 'A1', 'x')).toBe(false);
  });

  it('multi-comment sheet only edits the matching cell', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setComment(ws, { ref: 'A1', author: 'Alice', text: 'a1' });
    setComment(ws, { ref: 'B2', author: 'Bob', text: 'b2' });
    editCommentText(ws, 'B2', 'edited');
    expect(getComment(ws, 'A1')?.text).toBe('a1');
    expect(getComment(ws, 'B2')?.text).toBe('edited');
  });
});

describe('editCommentAuthor', () => {
  it('rewrites the author and leaves text + ref alone', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    setComment(ws, { ref: 'A1', author: 'Alice', text: 'note' });
    expect(editCommentAuthor(ws, 'A1', 'Carol')).toBe(true);
    const c = getComment(ws, 'A1');
    expect(c?.author).toBe('Carol');
    expect(c?.text).toBe('note');
  });

  it('returns false when the ref is missing', () => {
    const wb = createWorkbook();
    const ws = addWorksheet(wb, 'A');
    expect(editCommentAuthor(ws, 'A1', 'x')).toBe(false);
  });
});
