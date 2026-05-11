import { describe, expect, it } from 'vitest';
import { SHEET_MAIN_NS } from '../../src/xml/namespaces';
import { appendChild, el, elNs, findChild, findChildren, type XmlNode } from '../../src/xml/tree';

describe('el()', () => {
  it('creates a node with the given Clark name and empty attrs/children when nothing else is supplied', () => {
    const n = el(`{${SHEET_MAIN_NS}}c`);
    expect(n.name).toBe(`{${SHEET_MAIN_NS}}c`);
    expect(n.attrs).toEqual({});
    expect(n.children).toEqual([]);
    expect(n.text).toBeUndefined();
  });

  it('coerces attribute values to strings', () => {
    const n = el('row', { r: 1, ht: 12.5, customHeight: true });
    expect(n.attrs).toEqual({ r: '1', ht: '12.5', customHeight: 'true' });
  });

  it('drops null and undefined attrs', () => {
    const n = el('row', { r: 1, hidden: null, customHeight: undefined });
    expect(n.attrs).toEqual({ r: '1' });
  });

  it('copies the children array (caller mutation does not leak)', () => {
    const a = el('a');
    const kids: XmlNode[] = [a];
    const n = el('parent', {}, kids);
    kids.push(el('rogue'));
    expect(n.children).toEqual([a]);
  });

  it('preserves text when supplied', () => {
    const n = el('t', {}, [], 'hello');
    expect(n.text).toBe('hello');
  });
});

describe('elNs()', () => {
  it('builds a Clark-notation name from ns + local', () => {
    const n = elNs(SHEET_MAIN_NS, 'workbook');
    expect(n.name).toBe(`{${SHEET_MAIN_NS}}workbook`);
  });

  it('omits the namespace when blank', () => {
    expect(elNs(undefined, 'plain').name).toBe('plain');
    expect(elNs('', 'plain').name).toBe('plain');
  });
});

describe('findChild / findChildren', () => {
  it('returns the first matching child', () => {
    const a1 = el('a', { idx: 1 });
    const a2 = el('a', { idx: 2 });
    const root = el('root', {}, [el('b'), a1, a2]);
    expect(findChild(root, 'a')).toBe(a1);
  });

  it('returns undefined when no child matches', () => {
    const root = el('root', {}, [el('b')]);
    expect(findChild(root, 'a')).toBeUndefined();
  });

  it('findChildren returns every match in document order', () => {
    const a1 = el('a');
    const a2 = el('a');
    const root = el('root', {}, [el('b'), a1, el('c'), a2]);
    expect(findChildren(root, 'a')).toEqual([a1, a2]);
  });
});

describe('appendChild', () => {
  it('mutates and returns the parent', () => {
    const root = el('root');
    const child = el('child');
    const result = appendChild(root, child);
    expect(result).toBe(root);
    expect(root.children).toEqual([child]);
  });
});
