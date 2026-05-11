import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../src/io/node';
import {
  appendRel,
  findAllByType,
  findById,
  findByType,
  makeRelationships,
  relsFromBytes,
  relsToBytes,
} from '../../src/packaging/relationships';
import { openZip } from '../../src/zip/reader';

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURES = resolve(here, '../../reference/openpyxl/openpyxl/tests/data/genuine');

const T_OFFICE_DOC = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument';
const T_CORE_PROPS = 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';

describe('relationships — appendRel', () => {
  it('assigns rId1 first, rId2 second, etc.', () => {
    const r = makeRelationships();
    const a = appendRel(r, T_OFFICE_DOC, 'xl/workbook.xml');
    const b = appendRel(r, T_CORE_PROPS, 'docProps/core.xml');
    expect(a.id).toBe('rId1');
    expect(b.id).toBe('rId2');
  });

  it('skips rIds already manually present in the list', () => {
    const r = makeRelationships();
    r.rels.push({ id: 'rId1', type: T_OFFICE_DOC, target: 'foo' });
    r.rels.push({ id: 'rId3', type: T_CORE_PROPS, target: 'bar' });
    const next = appendRel(r, T_CORE_PROPS, 'baz');
    expect(next.id).toBe('rId2');
  });

  it('honours external targetMode', () => {
    const r = makeRelationships();
    const ext = appendRel(r, T_OFFICE_DOC, 'http://example.com/x.xlsx', 'External');
    expect(ext.targetMode).toBe('External');
  });

  it('findById / findByType / findAllByType', () => {
    const r = makeRelationships();
    appendRel(r, T_OFFICE_DOC, 'xl/workbook.xml');
    appendRel(r, T_CORE_PROPS, 'docProps/core.xml');
    appendRel(r, T_CORE_PROPS, 'docProps/core.xml.bak');

    expect(findById(r, 'rId1')?.target).toBe('xl/workbook.xml');
    expect(findById(r, 'rIdMissing')).toBeUndefined();
    expect(findByType(r, T_CORE_PROPS)?.target).toBe('docProps/core.xml');
    expect(findAllByType(r, T_CORE_PROPS).map((x) => x.target)).toEqual(['docProps/core.xml', 'docProps/core.xml.bak']);
  });
});

describe('relationships — XML round-trip', () => {
  it('hand-built rels round-trip', () => {
    const r = makeRelationships();
    appendRel(r, T_OFFICE_DOC, 'xl/workbook.xml');
    appendRel(r, T_CORE_PROPS, 'docProps/core.xml');
    const back = relsFromBytes(relsToBytes(r));
    expect(back).toEqual(r);
  });

  it('parses openpyxl genuine/empty.xlsx _rels/.rels', async () => {
    const zip = await openZip(fromBuffer(readFileSync(resolve(FIXTURES, 'empty.xlsx'))));
    const r = relsFromBytes(zip.read('_rels/.rels'));
    expect(r.rels.length).toBeGreaterThan(0);
    // Root rels must include the officeDocument relationship to xl/workbook.xml.
    const wb = findByType(r, T_OFFICE_DOC);
    expect(wb).toBeDefined();
    expect(wb?.target.endsWith('workbook.xml')).toBe(true);
  });

  it('parsing → re-serialising → re-parsing yields the same Relationships', async () => {
    const zip = await openZip(fromBuffer(readFileSync(resolve(FIXTURES, 'empty.xlsx'))));
    const original = zip.read('_rels/.rels');
    const r1 = relsFromBytes(original);
    const re = relsToBytes(r1);
    const r2 = relsFromBytes(re);
    expect(r2).toEqual(r1);
  });

  it('round-trips xl/_rels/workbook.xml.rels (multiple sheet relationships)', async () => {
    const zip = await openZip(fromBuffer(readFileSync(resolve(FIXTURES, 'empty.xlsx'))));
    const r = relsFromBytes(zip.read('xl/_rels/workbook.xml.rels'));
    // genuine/empty.xlsx defines 3 sheets + theme + styles relationships.
    expect(r.rels.length).toBeGreaterThanOrEqual(5);
    const re = relsToBytes(r);
    const back = relsFromBytes(re);
    expect(back).toEqual(r);
  });
});
