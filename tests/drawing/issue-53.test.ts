import { describe, expect, it } from 'vitest';
import { serializeFill } from '../../src/drawing/dml/dml-xml';
import type { Fill } from '../../src/drawing/dml/fill';

describe('issue #53 — DmlColorWithMods.mods omitted does not crash the serializer', () => {
  it('emits a valid solidFill when the caller forgets to supply `mods: []`', () => {
    const fill = {
      kind: 'solidFill' as const,
      color: { base: { kind: 'srgb' as const, value: 'FF0000' } } as never,
    } satisfies Fill;
    const xml = serializeFill(fill);
    expect(xml).toBe('<a:solidFill><a:srgbClr val="FF0000"></a:srgbClr></a:solidFill>');
  });

  it('emits a valid gradient stop when mods is missing on a stop colour', () => {
    const fill: Fill = {
      kind: 'gradFill',
      stops: [{ pos: 0, color: { base: { kind: 'srgb', value: '00FF00' } } as never }],
    };
    const xml = serializeFill(fill);
    expect(xml).toContain('<a:srgbClr val="00FF00"></a:srgbClr>');
  });

  it('emits a valid pattern fill when fgClr/bgClr have no mods', () => {
    const fill: Fill = {
      kind: 'pattFill',
      preset: 'pct50',
      fgClr: { base: { kind: 'srgb', value: 'AABBCC' } } as never,
      bgClr: { base: { kind: 'srgb', value: 'DDEEFF' } } as never,
    };
    const xml = serializeFill(fill);
    expect(xml).toContain('<a:srgbClr val="AABBCC"></a:srgbClr>');
    expect(xml).toContain('<a:srgbClr val="DDEEFF"></a:srgbClr>');
  });
});
