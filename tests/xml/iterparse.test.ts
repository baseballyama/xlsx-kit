import { readFileSync } from 'node:fs';
import { dirname, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { describe, expect, it } from 'vitest';
import { fromBuffer } from '../../../src/io/node';
import { OpenXmlSchemaError } from '../../../src/utils/exceptions';
import { iterParse, type SaxEvent } from '../../../src/xml/iterparse';
import { REL_NS, SHEET_MAIN_NS } from '../../../src/xml/namespaces';
import { openZip } from '../../../src/zip/reader';

const here = dirname(fileURLToPath(import.meta.url));
const FIXTURES = resolve(here, '../../../reference/openpyxl/openpyxl/tests/data/genuine');

const collect = async (input: Parameters<typeof iterParse>[0]): Promise<SaxEvent[]> => {
  const events: SaxEvent[] = [];
  for await (const e of iterParse(input)) events.push(e);
  return events;
};

describe('iterParse — basics', () => {
  it('emits start / end events for a self-closing element', async () => {
    expect(await collect('<r/>')).toEqual([
      { kind: 'start', name: 'r', attrs: {} },
      { kind: 'end', name: 'r' },
    ]);
  });

  it('emits a text event between start and end for a text-only element', async () => {
    expect(await collect('<t>hello</t>')).toEqual([
      { kind: 'start', name: 't', attrs: {} },
      { kind: 'text', text: 'hello' },
      { kind: 'end', name: 't' },
    ]);
  });

  it('preserves attribute values as Clark-keyed strings', async () => {
    const xml = `<workbook xmlns="${SHEET_MAIN_NS}" xmlns:r="${REL_NS}"><sheet r:id="rId1" name="S1"/></workbook>`;
    const events = await collect(xml);
    const start = events.find((e) => e.kind === 'start' && e.name.endsWith('}sheet'));
    expect(start).toBeDefined();
    expect(start?.kind).toBe('start');
    if (start?.kind !== 'start') return;
    expect(start.name).toBe(`{${SHEET_MAIN_NS}}sheet`);
    expect(start.attrs).toEqual({ name: 'S1', [`{${REL_NS}}id`]: 'rId1' });
  });

  it('drops xmlns / xmlns:* declarations from event attributes', async () => {
    const events = await collect(`<r xmlns="${SHEET_MAIN_NS}" xmlns:r="${REL_NS}" foo="bar"/>`);
    const start = events.find((e) => e.kind === 'start');
    expect(start?.kind).toBe('start');
    if (start?.kind !== 'start') return;
    expect(Object.keys(start.attrs).sort()).toEqual(['foo']);
  });

  it('accepts a Uint8Array input', async () => {
    const bytes = new TextEncoder().encode('<r a="1"/>');
    const events = await collect(bytes);
    expect(events.length).toBe(2);
  });

  it('expands the standard XML entities in text', async () => {
    const events = await collect('<t>&amp; &lt; &gt;</t>');
    const text = events.find((e) => e.kind === 'text');
    expect(text?.kind === 'text' && text.text).toBe('& < >');
  });
});

describe('iterParse — security', () => {
  it('rejects DOCTYPE declarations', async () => {
    await expect(collect('<!DOCTYPE foo SYSTEM "u"><foo/>')).rejects.toBeInstanceOf(OpenXmlSchemaError);
  });

  it('rejects loose <!ENTITY declarations', async () => {
    await expect(collect('<!ENTITY foo "bar"><r/>')).rejects.toBeInstanceOf(OpenXmlSchemaError);
  });
});

describe('iterParse — streaming input', () => {
  it('consumes a ReadableStream of Uint8Array chunks', async () => {
    const xml = `<workbook xmlns="${SHEET_MAIN_NS}"><sheets><sheet name="A"/><sheet name="B"/></sheets></workbook>`;
    const enc = new TextEncoder();
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        // Split deliberately mid-tag to exercise the chunked path.
        controller.enqueue(enc.encode(xml.slice(0, 100)));
        controller.enqueue(enc.encode(xml.slice(100)));
        controller.close();
      },
    });
    const events = await collect(stream);
    const sheetStarts = events.filter((e) => e.kind === 'start' && e.name === `{${SHEET_MAIN_NS}}sheet`);
    expect(sheetStarts.length).toBe(2);
  });

  it('rejects DOCTYPE hidden in the stream prologue', async () => {
    const enc = new TextEncoder();
    const stream = new ReadableStream<Uint8Array>({
      start(controller) {
        controller.enqueue(enc.encode('<!DOCTYPE x><x/>'));
        controller.close();
      },
    });
    await expect(collect(stream)).rejects.toBeInstanceOf(OpenXmlSchemaError);
  });
});

describe('iterParse — openpyxl genuine/sample.xlsx sheet1.xml', () => {
  it('counts every <row> and <c> element matching the file', async () => {
    const zip = await openZip(fromBuffer(readFileSync(resolve(FIXTURES, 'sample.xlsx'))));
    const events = await collect(zip.read('xl/worksheets/sheet1.xml'));

    let rows = 0;
    let cells = 0;
    for (const e of events) {
      if (e.kind === 'start' && e.name === `{${SHEET_MAIN_NS}}row`) rows++;
      if (e.kind === 'start' && e.name === `{${SHEET_MAIN_NS}}c`) cells++;
    }
    // Cross-checked against grep output captured this turn (2 rows, 2 cells).
    expect(rows).toBe(2);
    expect(cells).toBe(2);
  });

  it('start/end events nest correctly (matched count, balanced order)', async () => {
    const zip = await openZip(fromBuffer(readFileSync(resolve(FIXTURES, 'sample.xlsx'))));
    const events = await collect(zip.read('xl/worksheets/sheet1.xml'));
    const stack: string[] = [];
    let starts = 0;
    let ends = 0;
    for (const e of events) {
      if (e.kind === 'start') {
        starts++;
        stack.push(e.name);
      } else if (e.kind === 'end') {
        ends++;
        const top = stack.pop();
        expect(top).toBe(e.name);
      }
    }
    expect(stack).toEqual([]);
    expect(starts).toBe(ends);
    expect(starts).toBeGreaterThan(0);
  });
});
