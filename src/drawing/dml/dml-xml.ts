// DrawingML primitive parse / serialize. Per docs/plan/08-charts-drawings.md §4.

import { OpenXmlSchemaError } from '../../utils/exceptions';
import { DRAWING_NS, REL_NS } from '../../xml/namespaces';
import { findChild, findChildren, type XmlNode } from '../../xml/tree';
import {
  type ColorMod,
  type DmlColor,
  type DmlColorWithMods,
  SCHEME_COLOR_NAMES,
  type SchemeColorName,
  VALUED_COLOR_MOD_KINDS,
  VALUELESS_COLOR_MOD_KINDS,
} from './colors';
import type {
  Effect,
  EffectContainer,
  EffectList,
  EffectsRef,
  FillBlendMode,
  PresetShadowName,
  ShadowAlign,
} from './effect';
import { PRESET_SHADOW_NAMES } from './effect';
import type { Blip, BlipEffect, Fill, GradientLineDir, GradientStop, RelativeRect, TileFill, TileFlip } from './fill';
import type {
  AdjPoint2D,
  AdjustHandle,
  ConnectionSite,
  CustomGeometry,
  Geometry,
  GeometryPath,
  GuideRect,
  PathCommand,
  PathFill,
  PresetGeometry,
  ShapeGuide,
} from './geometry';
import type { LineCap, LineCompound, LineEnd, LineEndSize, LineEndType, LineProperties, PresetDash } from './line';
import type { BlackWhiteMode, Point2D, PositiveSize2D, ShapeProperties, Transform2D } from './shape-properties';

const A = (local: string): string => `{${DRAWING_NS}}${local}`;

const escapeText = (s: string): string => s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const escapeAttr = (s: string): string => escapeText(s).replace(/"/g, '&quot;');

const valAttr = (n: XmlNode | undefined): string | undefined => n?.attrs['val'];
const intAttr = (n: XmlNode, name: string): number | undefined => {
  const v = n.attrs[name];
  if (v === undefined) return undefined;
  const x = Number.parseInt(v, 10);
  return Number.isInteger(x) ? x : undefined;
};
const boolAttr = (n: XmlNode, name: string): boolean | undefined => {
  const v = n.attrs[name];
  if (v === undefined) return undefined;
  if (v === '1' || v === 'true') return true;
  if (v === '0' || v === 'false') return false;
  return undefined;
};

// ---- Colors ----------------------------------------------------------------

const VALUED_MOD_SET = new Set<string>(VALUED_COLOR_MOD_KINDS);
const VALUELESS_MOD_SET = new Set<string>(VALUELESS_COLOR_MOD_KINDS);
const SCHEME_NAME_SET = new Set<string>(SCHEME_COLOR_NAMES);

const colorBaseFromNode = (el: XmlNode): DmlColor | undefined => {
  const local = el.name.split('}').pop() ?? el.name;
  switch (local) {
    case 'srgbClr': {
      const v = el.attrs['val'];
      if (!v) return undefined;
      return { kind: 'srgb', value: v.toUpperCase() };
    }
    case 'sysClr': {
      const v = el.attrs['val'];
      if (!v) return undefined;
      const lastClr = el.attrs['lastClr'];
      return { kind: 'sysClr', value: v, ...(lastClr ? { lastClr: lastClr.toUpperCase() } : {}) };
    }
    case 'schemeClr': {
      const v = el.attrs['val'];
      if (!v || !SCHEME_NAME_SET.has(v)) return undefined;
      return { kind: 'schemeClr', value: v as SchemeColorName };
    }
    case 'prstClr': {
      const v = el.attrs['val'];
      if (!v) return undefined;
      return { kind: 'prstClr', value: v };
    }
    case 'hslClr': {
      const hue = intAttr(el, 'hue') ?? 0;
      const sat = intAttr(el, 'sat') ?? 0;
      const lum = intAttr(el, 'lum') ?? 0;
      return { kind: 'hslClr', hue, sat, lum };
    }
    case 'scrgbClr': {
      const r = intAttr(el, 'r') ?? 0;
      const g = intAttr(el, 'g') ?? 0;
      const b = intAttr(el, 'b') ?? 0;
      return { kind: 'scrgbClr', r, g, b };
    }
    default:
      return undefined;
  }
};

const colorModsFromNode = (baseEl: XmlNode): ColorMod[] => {
  const mods: ColorMod[] = [];
  for (const c of baseEl.children) {
    if (typeof c === 'string') continue;
    const local = c.name.split('}').pop() ?? c.name;
    if (VALUELESS_MOD_SET.has(local)) {
      mods.push({ kind: local } as ColorMod);
    } else if (VALUED_MOD_SET.has(local)) {
      const val = intAttr(c, 'val');
      if (val !== undefined) mods.push({ kind: local, val } as ColorMod);
    }
  }
  return mods;
};

/** Parse the first known color base element (with its mods) found inside `parent`. */
export const parseDmlColor = (parent: XmlNode): DmlColorWithMods | undefined => {
  for (const c of parent.children) {
    if (typeof c === 'string') continue;
    const base = colorBaseFromNode(c);
    if (!base) continue;
    return { base, mods: colorModsFromNode(c) };
  }
  return undefined;
};

const serializeColorMod = (m: ColorMod): string => {
  if (VALUELESS_MOD_SET.has(m.kind)) return `<a:${m.kind}/>`;
  // VALUED_MOD_SET — every entry has `val`.
  return `<a:${m.kind} val="${(m as { val: number }).val}"/>`;
};

const serializeColorBase = (c: DmlColorWithMods): string => {
  const inner = c.mods.map(serializeColorMod).join('');
  switch (c.base.kind) {
    case 'srgb':
      return `<a:srgbClr val="${c.base.value}">${inner}</a:srgbClr>`;
    case 'sysClr': {
      const last = c.base.lastClr ? ` lastClr="${c.base.lastClr}"` : '';
      return `<a:sysClr val="${c.base.value}"${last}>${inner}</a:sysClr>`;
    }
    case 'schemeClr':
      return `<a:schemeClr val="${c.base.value}">${inner}</a:schemeClr>`;
    case 'prstClr':
      return `<a:prstClr val="${c.base.value}">${inner}</a:prstClr>`;
    case 'hslClr':
      return `<a:hslClr hue="${c.base.hue}" sat="${c.base.sat}" lum="${c.base.lum}">${inner}</a:hslClr>`;
    case 'scrgbClr':
      return `<a:scrgbClr r="${c.base.r}" g="${c.base.g}" b="${c.base.b}">${inner}</a:scrgbClr>`;
  }
};

/** Serialize a color directly (without an enclosing wrapper). */
export const serializeDmlColor = (c: DmlColorWithMods): string => serializeColorBase(c);

// ---- Fills -----------------------------------------------------------------

const FILL_LOCAL_NAMES = new Set(['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill']);

const parseRelativeRect = (el: XmlNode): RelativeRect => {
  const out: RelativeRect = {};
  const l = intAttr(el, 'l');
  const t = intAttr(el, 't');
  const r = intAttr(el, 'r');
  const b = intAttr(el, 'b');
  if (l !== undefined) out.l = l;
  if (t !== undefined) out.t = t;
  if (r !== undefined) out.r = r;
  if (b !== undefined) out.b = b;
  return out;
};

const serializeRelativeRect = (tag: string, r: RelativeRect): string => {
  const a: string[] = [];
  if (r.l !== undefined) a.push(`l="${r.l}"`);
  if (r.t !== undefined) a.push(`t="${r.t}"`);
  if (r.r !== undefined) a.push(`r="${r.r}"`);
  if (r.b !== undefined) a.push(`b="${r.b}"`);
  return `<${tag}${a.length > 0 ? ` ${a.join(' ')}` : ''}/>`;
};

const parseGradientStops = (gsLstEl: XmlNode): GradientStop[] => {
  const stops: GradientStop[] = [];
  for (const gs of findChildren(gsLstEl, A('gs'))) {
    const pos = intAttr(gs, 'pos');
    const color = parseDmlColor(gs);
    if (pos === undefined || !color) continue;
    stops.push({ pos, color });
  }
  return stops;
};

const parseGradientLineDir = (gradEl: XmlNode): GradientLineDir | undefined => {
  const lin = findChild(gradEl, A('lin'));
  if (lin) {
    const ang = intAttr(lin, 'ang') ?? 0;
    const scaled = boolAttr(lin, 'scaled');
    return { kind: 'lin', ang, ...(scaled !== undefined ? { scaled } : {}) };
  }
  const path = findChild(gradEl, A('path'));
  if (path) {
    const pathTypeRaw = path.attrs['path'];
    const pathType =
      pathTypeRaw === 'shape' || pathTypeRaw === 'circle' || pathTypeRaw === 'rect' ? pathTypeRaw : 'shape';
    const tileRectEl = findChild(path, A('fillToRect'));
    return {
      kind: 'path',
      pathType,
      ...(tileRectEl ? { tileRect: parseRelativeRect(tileRectEl) } : {}),
    };
  }
  return undefined;
};

const parseBlipEffectChild = (c: XmlNode): BlipEffect | undefined => {
  const local = c.name.split('}').pop() ?? c.name;
  switch (local) {
    case 'biLevel':
      return { kind: 'biLevel', thresh: intAttr(c, 'thresh') ?? 0 };
    case 'blur': {
      const rad = intAttr(c, 'rad') ?? 0;
      const grow = boolAttr(c, 'grow');
      return { kind: 'blur', rad, ...(grow !== undefined ? { grow } : {}) };
    }
    case 'clrChange': {
      const useA = boolAttr(c, 'useA');
      const fromEl = findChild(c, A('clrFrom'));
      const toEl = findChild(c, A('clrTo'));
      if (!fromEl || !toEl) return undefined;
      const from = parseDmlColor(fromEl)?.base;
      const to = parseDmlColor(toEl)?.base;
      if (!from || !to) return undefined;
      return { kind: 'clrChange', clrFrom: from, clrTo: to, ...(useA !== undefined ? { useA } : {}) };
    }
    case 'clrRepl': {
      const cw = parseDmlColor(c);
      if (!cw) return undefined;
      return { kind: 'clrRepl', color: cw.base };
    }
    case 'duotone': {
      const colors: DmlColor[] = [];
      for (const child of c.children) {
        if (typeof child === 'string') continue;
        const base = colorBaseFromNode(child);
        if (base) colors.push(base);
        if (colors.length === 2) break;
      }
      if (colors.length !== 2) return undefined;
      return { kind: 'duotone', colors: [colors[0] as DmlColor, colors[1] as DmlColor] };
    }
    case 'grayscl':
      return { kind: 'grayscl' };
    case 'lum': {
      const bright = intAttr(c, 'bright');
      const contrast = intAttr(c, 'contrast');
      return {
        kind: 'lum',
        ...(bright !== undefined ? { bright } : {}),
        ...(contrast !== undefined ? { contrast } : {}),
      };
    }
    case 'tint': {
      const hue = intAttr(c, 'hue');
      const amt = intAttr(c, 'amt');
      return {
        kind: 'tint',
        ...(hue !== undefined ? { hue } : {}),
        ...(amt !== undefined ? { amt } : {}),
      };
    }
    case 'alphaModFix':
      return { kind: 'alphaModFix', amt: intAttr(c, 'amt') ?? 0 };
    default:
      return undefined;
  }
};

const parseBlip = (blipEl: XmlNode): Blip => {
  const out: Blip = {};
  const embed = blipEl.attrs[`{${REL_NS}}embed`];
  const link = blipEl.attrs[`{${REL_NS}}link`];
  const cstateRaw = blipEl.attrs['cstate'];
  if (embed) out.embedRId = embed;
  if (link) out.linkRId = link;
  if (cstateRaw === 'email' || cstateRaw === 'screen' || cstateRaw === 'print' || cstateRaw === 'hqprint') {
    out.cstate = cstateRaw;
  }
  const effects: BlipEffect[] = [];
  for (const c of blipEl.children) {
    if (typeof c === 'string') continue;
    const eff = parseBlipEffectChild(c);
    if (eff) effects.push(eff);
  }
  if (effects.length > 0) out.effects = effects;
  return out;
};

const parseTileFill = (tileEl: XmlNode): TileFill => {
  const out: TileFill = {};
  const tx = intAttr(tileEl, 'tx');
  const ty = intAttr(tileEl, 'ty');
  const sx = intAttr(tileEl, 'sx');
  const sy = intAttr(tileEl, 'sy');
  const flipRaw = tileEl.attrs['flip'];
  const algnRaw = tileEl.attrs['algn'];
  if (tx !== undefined) out.tx = tx;
  if (ty !== undefined) out.ty = ty;
  if (sx !== undefined) out.sx = sx;
  if (sy !== undefined) out.sy = sy;
  if (flipRaw === 'x' || flipRaw === 'y' || flipRaw === 'xy' || flipRaw === 'none') out.flip = flipRaw;
  const validAlgn: ReadonlyArray<string> = ['tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br'];
  if (algnRaw && validAlgn.includes(algnRaw)) out.algn = algnRaw as NonNullable<TileFill['algn']>;
  return out;
};

const parseFillNode = (el: XmlNode): Fill | undefined => {
  const local = el.name.split('}').pop() ?? el.name;
  switch (local) {
    case 'noFill':
      return { kind: 'noFill' };
    case 'grpFill':
      return { kind: 'grpFill' };
    case 'solidFill': {
      const color = parseDmlColor(el);
      if (!color) return undefined;
      return { kind: 'solidFill', color };
    }
    case 'gradFill': {
      const flipRaw = el.attrs['flip'];
      const flip: TileFlip | undefined =
        flipRaw === 'x' || flipRaw === 'y' || flipRaw === 'xy' || flipRaw === 'none' ? flipRaw : undefined;
      const rotWithShape = boolAttr(el, 'rotWithShape');
      const gsLstEl = findChild(el, A('gsLst'));
      const stops = gsLstEl ? parseGradientStops(gsLstEl) : [];
      const lineDir = parseGradientLineDir(el);
      return {
        kind: 'gradFill',
        stops,
        ...(flip ? { flip } : {}),
        ...(rotWithShape !== undefined ? { rotWithShape } : {}),
        ...(lineDir ? { lineDir } : {}),
      };
    }
    case 'blipFill': {
      const dpi = intAttr(el, 'dpi');
      const rotWithShape = boolAttr(el, 'rotWithShape');
      const blipEl = findChild(el, A('blip'));
      const blip = blipEl ? parseBlip(blipEl) : {};
      const tileEl = findChild(el, A('tile'));
      const stretchEl = findChild(el, A('stretch'));
      const srcRectEl = findChild(el, A('srcRect'));
      const out: Fill = {
        kind: 'blipFill',
        blip,
        ...(dpi !== undefined ? { dpi } : {}),
        ...(rotWithShape !== undefined ? { rotWithShape } : {}),
        ...(tileEl ? { tile: parseTileFill(tileEl) } : {}),
        ...(stretchEl
          ? {
              stretch: (() => {
                const fr = findChild(stretchEl, A('fillRect'));
                return fr ? { fillRect: parseRelativeRect(fr) } : {};
              })(),
            }
          : {}),
        ...(srcRectEl ? { srcRect: parseRelativeRect(srcRectEl) } : {}),
      };
      return out;
    }
    case 'pattFill': {
      const preset = el.attrs['prst'] ?? '';
      const fgEl = findChild(el, A('fgClr'));
      const bgEl = findChild(el, A('bgClr'));
      const out: Fill = {
        kind: 'pattFill',
        preset,
        ...(fgEl ? { fgClr: parseDmlColor(fgEl) ?? { base: { kind: 'srgb', value: '000000' }, mods: [] } } : {}),
        ...(bgEl ? { bgClr: parseDmlColor(bgEl) ?? { base: { kind: 'srgb', value: 'FFFFFF' }, mods: [] } } : {}),
      };
      return out;
    }
    default:
      return undefined;
  }
};

/** Parse the first fill element under `parent` (any of the 6 kinds). */
export const parseFill = (parent: XmlNode): Fill | undefined => {
  for (const c of parent.children) {
    if (typeof c === 'string') continue;
    const local = c.name.split('}').pop() ?? c.name;
    if (FILL_LOCAL_NAMES.has(local)) return parseFillNode(c);
  }
  return undefined;
};

const serializeBlipEffect = (e: BlipEffect): string => {
  switch (e.kind) {
    case 'biLevel':
      return `<a:biLevel thresh="${e.thresh}"/>`;
    case 'blur':
      return `<a:blur rad="${e.rad}"${e.grow !== undefined ? ` grow="${e.grow ? '1' : '0'}"` : ''}/>`;
    case 'clrChange': {
      const useA = e.useA !== undefined ? ` useA="${e.useA ? '1' : '0'}"` : '';
      const from = serializeColorBase({ base: e.clrFrom, mods: [] });
      const to = serializeColorBase({ base: e.clrTo, mods: [] });
      return `<a:clrChange${useA}><a:clrFrom>${from}</a:clrFrom><a:clrTo>${to}</a:clrTo></a:clrChange>`;
    }
    case 'clrRepl':
      return `<a:clrRepl>${serializeColorBase({ base: e.color, mods: [] })}</a:clrRepl>`;
    case 'duotone':
      return `<a:duotone>${serializeColorBase({ base: e.colors[0], mods: [] })}${serializeColorBase({ base: e.colors[1], mods: [] })}</a:duotone>`;
    case 'grayscl':
      return '<a:grayscl/>';
    case 'lum': {
      const a: string[] = [];
      if (e.bright !== undefined) a.push(`bright="${e.bright}"`);
      if (e.contrast !== undefined) a.push(`contrast="${e.contrast}"`);
      return `<a:lum${a.length > 0 ? ` ${a.join(' ')}` : ''}/>`;
    }
    case 'tint': {
      const a: string[] = [];
      if (e.hue !== undefined) a.push(`hue="${e.hue}"`);
      if (e.amt !== undefined) a.push(`amt="${e.amt}"`);
      return `<a:tint${a.length > 0 ? ` ${a.join(' ')}` : ''}/>`;
    }
    case 'alphaModFix':
      return `<a:alphaModFix amt="${e.amt}"/>`;
  }
};

const serializeBlip = (b: Blip): string => {
  const a: string[] = [];
  if (b.embedRId) a.push(`r:embed="${escapeAttr(b.embedRId)}"`);
  if (b.linkRId) a.push(`r:link="${escapeAttr(b.linkRId)}"`);
  if (b.cstate) a.push(`cstate="${b.cstate}"`);
  const inner = (b.effects ?? []).map(serializeBlipEffect).join('');
  return inner === ''
    ? `<a:blip${a.length > 0 ? ` ${a.join(' ')}` : ''}/>`
    : `<a:blip${a.length > 0 ? ` ${a.join(' ')}` : ''}>${inner}</a:blip>`;
};

const serializeTileFill = (t: TileFill): string => {
  const a: string[] = [];
  if (t.tx !== undefined) a.push(`tx="${t.tx}"`);
  if (t.ty !== undefined) a.push(`ty="${t.ty}"`);
  if (t.sx !== undefined) a.push(`sx="${t.sx}"`);
  if (t.sy !== undefined) a.push(`sy="${t.sy}"`);
  if (t.flip) a.push(`flip="${t.flip}"`);
  if (t.algn) a.push(`algn="${t.algn}"`);
  return `<a:tile${a.length > 0 ? ` ${a.join(' ')}` : ''}/>`;
};

const serializeGradientLineDir = (d: GradientLineDir): string => {
  if (d.kind === 'lin') {
    const scaled = d.scaled !== undefined ? ` scaled="${d.scaled ? '1' : '0'}"` : '';
    return `<a:lin ang="${d.ang}"${scaled}/>`;
  }
  const inner = d.tileRect ? serializeRelativeRect('a:fillToRect', d.tileRect) : '';
  return `<a:path path="${d.pathType}">${inner}</a:path>`;
};

export const serializeFill = (f: Fill): string => {
  switch (f.kind) {
    case 'noFill':
      return '<a:noFill/>';
    case 'grpFill':
      return '<a:grpFill/>';
    case 'solidFill':
      return `<a:solidFill>${serializeColorBase(f.color)}</a:solidFill>`;
    case 'gradFill': {
      const a: string[] = [];
      if (f.flip) a.push(`flip="${f.flip}"`);
      if (f.rotWithShape !== undefined) a.push(`rotWithShape="${f.rotWithShape ? '1' : '0'}"`);
      const stops = `<a:gsLst>${f.stops
        .map((s) => `<a:gs pos="${s.pos}">${serializeColorBase(s.color)}</a:gs>`)
        .join('')}</a:gsLst>`;
      const dir = f.lineDir ? serializeGradientLineDir(f.lineDir) : '';
      return `<a:gradFill${a.length > 0 ? ` ${a.join(' ')}` : ''}>${stops}${dir}</a:gradFill>`;
    }
    case 'blipFill': {
      const a: string[] = [];
      if (f.dpi !== undefined) a.push(`dpi="${f.dpi}"`);
      if (f.rotWithShape !== undefined) a.push(`rotWithShape="${f.rotWithShape ? '1' : '0'}"`);
      const blip = serializeBlip(f.blip);
      const src = f.srcRect ? serializeRelativeRect('a:srcRect', f.srcRect) : '';
      const tile = f.tile ? serializeTileFill(f.tile) : '';
      let stretch = '';
      if (f.stretch) {
        const inner = f.stretch.fillRect ? serializeRelativeRect('a:fillRect', f.stretch.fillRect) : '';
        stretch = `<a:stretch>${inner}</a:stretch>`;
      }
      return `<a:blipFill${a.length > 0 ? ` ${a.join(' ')}` : ''}>${blip}${src}${tile}${stretch}</a:blipFill>`;
    }
    case 'pattFill': {
      const fg = f.fgClr ? `<a:fgClr>${serializeColorBase(f.fgClr)}</a:fgClr>` : '';
      const bg = f.bgClr ? `<a:bgClr>${serializeColorBase(f.bgClr)}</a:bgClr>` : '';
      return `<a:pattFill prst="${escapeAttr(f.preset)}">${fg}${bg}</a:pattFill>`;
    }
  }
};

// ---- Lines -----------------------------------------------------------------

const VALID_LINE_END_TYPES: ReadonlyArray<string> = ['none', 'triangle', 'stealth', 'diamond', 'oval', 'arrow'];
const VALID_LINE_END_SIZES: ReadonlyArray<string> = ['sm', 'med', 'lg'];

const parseLineEnd = (el: XmlNode): LineEnd => {
  const out: LineEnd = {};
  const t = el.attrs['type'];
  const w = el.attrs['w'];
  const len = el.attrs['len'];
  if (t && VALID_LINE_END_TYPES.includes(t)) out.type = t as LineEndType;
  if (w && VALID_LINE_END_SIZES.includes(w)) out.w = w as LineEndSize;
  if (len && VALID_LINE_END_SIZES.includes(len)) out.len = len as LineEndSize;
  return out;
};

const PRESET_DASHES: ReadonlyArray<string> = [
  'solid',
  'dot',
  'dash',
  'lgDash',
  'dashDot',
  'lgDashDot',
  'lgDashDotDot',
  'sysDash',
  'sysDot',
  'sysDashDot',
  'sysDashDotDot',
];

export const parseLine = (el: XmlNode): LineProperties => {
  const out: LineProperties = {};
  const w = intAttr(el, 'w');
  if (w !== undefined) out.w = w;
  const cap = el.attrs['cap'];
  if (cap === 'rnd' || cap === 'sq' || cap === 'flat') out.cap = cap as LineCap;
  const cmpd = el.attrs['cmpd'];
  if (cmpd === 'sng' || cmpd === 'dbl' || cmpd === 'thickThin' || cmpd === 'thinThick' || cmpd === 'tri')
    out.cmpd = cmpd as LineCompound;
  const algn = el.attrs['algn'];
  if (algn === 'ctr' || algn === 'in') out.algn = algn;
  const fill = parseFill(el);
  if (fill) out.fill = fill;
  const prstDashEl = findChild(el, A('prstDash'));
  const custDashEl = findChild(el, A('custDash'));
  if (prstDashEl) {
    const v = valAttr(prstDashEl);
    if (v && PRESET_DASHES.includes(v)) out.dash = { kind: 'preset', val: v as PresetDash };
  } else if (custDashEl) {
    const pattern: number[] = [];
    for (const ds of findChildren(custDashEl, A('ds'))) {
      const d = intAttr(ds, 'd');
      const sp = intAttr(ds, 'sp');
      if (d !== undefined) pattern.push(d);
      if (sp !== undefined) pattern.push(sp);
    }
    out.dash = { kind: 'custDash', pattern };
  }
  const roundEl = findChild(el, A('round'));
  const bevelEl = findChild(el, A('bevel'));
  const miterEl = findChild(el, A('miter'));
  if (roundEl) out.join = { kind: 'round' };
  else if (bevelEl) out.join = { kind: 'bevel' };
  else if (miterEl) {
    const lim = intAttr(miterEl, 'lim');
    out.join = { kind: 'miter', ...(lim !== undefined ? { lim } : {}) };
  }
  const headEnd = findChild(el, A('headEnd'));
  const tailEnd = findChild(el, A('tailEnd'));
  if (headEnd) out.headEnd = parseLineEnd(headEnd);
  if (tailEnd) out.tailEnd = parseLineEnd(tailEnd);
  return out;
};

const serializeLineEnd = (tag: string, e: LineEnd): string => {
  const a: string[] = [];
  if (e.type) a.push(`type="${e.type}"`);
  if (e.w) a.push(`w="${e.w}"`);
  if (e.len) a.push(`len="${e.len}"`);
  return `<${tag}${a.length > 0 ? ` ${a.join(' ')}` : ''}/>`;
};

export const serializeLine = (ln: LineProperties): string => {
  const a: string[] = [];
  if (ln.w !== undefined) a.push(`w="${ln.w}"`);
  if (ln.cap) a.push(`cap="${ln.cap}"`);
  if (ln.cmpd) a.push(`cmpd="${ln.cmpd}"`);
  if (ln.algn) a.push(`algn="${ln.algn}"`);
  const parts: string[] = [`<a:ln${a.length > 0 ? ` ${a.join(' ')}` : ''}>`];
  if (ln.fill) parts.push(serializeFill(ln.fill));
  if (ln.dash) {
    if (ln.dash.kind === 'preset') {
      parts.push(`<a:prstDash val="${ln.dash.val}"/>`);
    } else {
      const dsParts: string[] = [];
      for (let i = 0; i + 1 < ln.dash.pattern.length; i += 2) {
        dsParts.push(`<a:ds d="${ln.dash.pattern[i]}" sp="${ln.dash.pattern[i + 1]}"/>`);
      }
      parts.push(`<a:custDash>${dsParts.join('')}</a:custDash>`);
    }
  }
  if (ln.join) {
    if (ln.join.kind === 'round') parts.push('<a:round/>');
    else if (ln.join.kind === 'bevel') parts.push('<a:bevel/>');
    else parts.push(`<a:miter${ln.join.lim !== undefined ? ` lim="${ln.join.lim}"` : ''}/>`);
  }
  if (ln.headEnd) parts.push(serializeLineEnd('a:headEnd', ln.headEnd));
  if (ln.tailEnd) parts.push(serializeLineEnd('a:tailEnd', ln.tailEnd));
  parts.push('</a:ln>');
  return parts.join('');
};

// ---- Geometry --------------------------------------------------------------

const parseShapeGuides = (lstEl: XmlNode): ShapeGuide[] => {
  const out: ShapeGuide[] = [];
  for (const gd of findChildren(lstEl, A('gd'))) {
    const name = gd.attrs['name'];
    const fmla = gd.attrs['fmla'];
    if (name === undefined || fmla === undefined) continue;
    out.push({ name, fmla });
  }
  return out;
};

const serializeShapeGuides = (tag: string, guides: ShapeGuide[]): string => {
  if (guides.length === 0) return `<${tag}/>`;
  const inner = guides.map((g) => `<a:gd name="${escapeAttr(g.name)}" fmla="${escapeAttr(g.fmla)}"/>`).join('');
  return `<${tag}>${inner}</${tag}>`;
};

const parseAdjPoint = (el: XmlNode): AdjPoint2D => ({
  x: el.attrs['x'] ?? '0',
  y: el.attrs['y'] ?? '0',
});

const parseConnectionSites = (lstEl: XmlNode): ConnectionSite[] => {
  const out: ConnectionSite[] = [];
  for (const cxn of findChildren(lstEl, A('cxn'))) {
    const ang = cxn.attrs['ang'] ?? '0';
    const posEl = findChild(cxn, A('pos'));
    if (!posEl) continue;
    out.push({ ang, pos: parseAdjPoint(posEl) });
  }
  return out;
};

const serializeConnectionSites = (sites: ConnectionSite[]): string => {
  if (sites.length === 0) return '<a:cxnLst/>';
  const inner = sites
    .map(
      (s) => `<a:cxn ang="${escapeAttr(s.ang)}"><a:pos x="${escapeAttr(s.pos.x)}" y="${escapeAttr(s.pos.y)}"/></a:cxn>`,
    )
    .join('');
  return `<a:cxnLst>${inner}</a:cxnLst>`;
};

const parseAdjustHandles = (lstEl: XmlNode): AdjustHandle[] => {
  const out: AdjustHandle[] = [];
  for (const child of lstEl.children) {
    if (typeof child === 'string') continue;
    const local = child.name.split('}').pop() ?? child.name;
    if (local !== 'ahXY' && local !== 'ahPolar') continue;
    const posEl = findChild(child, A('pos'));
    if (!posEl) continue;
    const handle: AdjustHandle = {
      kind: local === 'ahXY' ? 'xy' : 'polar',
      pos: parseAdjPoint(posEl),
    };
    const a = child.attrs;
    if (a['gdRefX'] !== undefined) handle.gdRefX = a['gdRefX'];
    if (a['minX'] !== undefined) handle.minX = a['minX'];
    if (a['maxX'] !== undefined) handle.maxX = a['maxX'];
    if (a['gdRefY'] !== undefined) handle.gdRefY = a['gdRefY'];
    if (a['minY'] !== undefined) handle.minY = a['minY'];
    if (a['maxY'] !== undefined) handle.maxY = a['maxY'];
    if (a['gdRefR'] !== undefined) handle.gdRefR = a['gdRefR'];
    if (a['minR'] !== undefined) handle.minR = a['minR'];
    if (a['maxR'] !== undefined) handle.maxR = a['maxR'];
    if (a['gdRefAng'] !== undefined) handle.gdRefAng = a['gdRefAng'];
    if (a['minAng'] !== undefined) handle.minAng = a['minAng'];
    if (a['maxAng'] !== undefined) handle.maxAng = a['maxAng'];
    out.push(handle);
  }
  return out;
};

const serializeAdjustHandles = (handles: AdjustHandle[]): string => {
  if (handles.length === 0) return '<a:ahLst/>';
  const inner = handles
    .map((h) => {
      const tag = h.kind === 'xy' ? 'a:ahXY' : 'a:ahPolar';
      const attrs: string[] = [];
      const push = (k: string, v: string | undefined): void => {
        if (v !== undefined) attrs.push(`${k}="${escapeAttr(v)}"`);
      };
      push('gdRefX', h.gdRefX);
      push('minX', h.minX);
      push('maxX', h.maxX);
      push('gdRefY', h.gdRefY);
      push('minY', h.minY);
      push('maxY', h.maxY);
      push('gdRefR', h.gdRefR);
      push('minR', h.minR);
      push('maxR', h.maxR);
      push('gdRefAng', h.gdRefAng);
      push('minAng', h.minAng);
      push('maxAng', h.maxAng);
      return `<${tag}${attrs.length > 0 ? ` ${attrs.join(' ')}` : ''}><a:pos x="${escapeAttr(h.pos.x)}" y="${escapeAttr(h.pos.y)}"/></${tag}>`;
    })
    .join('');
  return `<a:ahLst>${inner}</a:ahLst>`;
};

const parseRectGuide = (el: XmlNode): GuideRect => ({
  l: el.attrs['l'] ?? '0',
  t: el.attrs['t'] ?? '0',
  r: el.attrs['r'] ?? '0',
  b: el.attrs['b'] ?? '0',
});

const VALID_PATH_FILLS: ReadonlyArray<string> = ['none', 'norm', 'lighten', 'lightenLess', 'darken', 'darkenLess'];

const parsePathPt = (el: XmlNode): Point2D => ({
  x: intAttr(el, 'x') ?? 0,
  y: intAttr(el, 'y') ?? 0,
});

const parsePathCommand = (el: XmlNode): PathCommand | undefined => {
  const local = el.name.split('}').pop() ?? el.name;
  switch (local) {
    case 'moveTo':
    case 'lnTo': {
      const ptEl = findChild(el, A('pt'));
      if (!ptEl) return undefined;
      return { kind: local, pt: parsePathPt(ptEl) };
    }
    case 'arcTo':
      return {
        kind: 'arcTo',
        wR: el.attrs['wR'] ?? '0',
        hR: el.attrs['hR'] ?? '0',
        stAng: el.attrs['stAng'] ?? '0',
        swAng: el.attrs['swAng'] ?? '0',
      };
    case 'quadBezTo': {
      const pts: Point2D[] = [];
      for (const pt of findChildren(el, A('pt'))) pts.push(parsePathPt(pt));
      if (pts.length < 2) return undefined;
      return { kind: 'quadBezTo', pts: [pts[0] as Point2D, pts[1] as Point2D] };
    }
    case 'cubicBezTo': {
      const pts: Point2D[] = [];
      for (const pt of findChildren(el, A('pt'))) pts.push(parsePathPt(pt));
      if (pts.length < 3) return undefined;
      return {
        kind: 'cubicBezTo',
        pts: [pts[0] as Point2D, pts[1] as Point2D, pts[2] as Point2D],
      };
    }
    case 'close':
      return { kind: 'close' };
    default:
      return undefined;
  }
};

const parsePath = (el: XmlNode): GeometryPath => {
  const out: GeometryPath = { commands: [] };
  const w = intAttr(el, 'w');
  const h = intAttr(el, 'h');
  if (w !== undefined) out.w = w;
  if (h !== undefined) out.h = h;
  const fillRaw = el.attrs['fill'];
  if (fillRaw && VALID_PATH_FILLS.includes(fillRaw)) out.fill = fillRaw as PathFill;
  const stroke = boolAttr(el, 'stroke');
  if (stroke !== undefined) out.stroke = stroke;
  const extrusionOk = boolAttr(el, 'extrusionOk');
  if (extrusionOk !== undefined) out.extrusionOk = extrusionOk;
  for (const c of el.children) {
    if (typeof c === 'string') continue;
    const cmd = parsePathCommand(c);
    if (cmd) out.commands.push(cmd);
  }
  return out;
};

const serializePathCommand = (cmd: PathCommand): string => {
  switch (cmd.kind) {
    case 'moveTo':
      return `<a:moveTo><a:pt x="${cmd.pt.x}" y="${cmd.pt.y}"/></a:moveTo>`;
    case 'lnTo':
      return `<a:lnTo><a:pt x="${cmd.pt.x}" y="${cmd.pt.y}"/></a:lnTo>`;
    case 'arcTo':
      return `<a:arcTo wR="${escapeAttr(cmd.wR)}" hR="${escapeAttr(cmd.hR)}" stAng="${escapeAttr(cmd.stAng)}" swAng="${escapeAttr(cmd.swAng)}"/>`;
    case 'quadBezTo':
      return `<a:quadBezTo><a:pt x="${cmd.pts[0].x}" y="${cmd.pts[0].y}"/><a:pt x="${cmd.pts[1].x}" y="${cmd.pts[1].y}"/></a:quadBezTo>`;
    case 'cubicBezTo':
      return `<a:cubicBezTo><a:pt x="${cmd.pts[0].x}" y="${cmd.pts[0].y}"/><a:pt x="${cmd.pts[1].x}" y="${cmd.pts[1].y}"/><a:pt x="${cmd.pts[2].x}" y="${cmd.pts[2].y}"/></a:cubicBezTo>`;
    case 'close':
      return '<a:close/>';
  }
};

const serializePath = (p: GeometryPath): string => {
  const a: string[] = [];
  if (p.w !== undefined) a.push(`w="${p.w}"`);
  if (p.h !== undefined) a.push(`h="${p.h}"`);
  if (p.fill) a.push(`fill="${p.fill}"`);
  if (p.stroke !== undefined) a.push(`stroke="${p.stroke ? '1' : '0'}"`);
  if (p.extrusionOk !== undefined) a.push(`extrusionOk="${p.extrusionOk ? '1' : '0'}"`);
  const inner = p.commands.map(serializePathCommand).join('');
  return `<a:path${a.length > 0 ? ` ${a.join(' ')}` : ''}>${inner}</a:path>`;
};

export const parseGeometry = (parent: XmlNode): Geometry | undefined => {
  const prstEl = findChild(parent, A('prstGeom'));
  if (prstEl) {
    const prst = prstEl.attrs['prst'] ?? '';
    const avLstEl = findChild(prstEl, A('avLst'));
    const out: PresetGeometry = { kind: 'preset', prst };
    if (avLstEl) {
      const guides = parseShapeGuides(avLstEl);
      if (guides.length > 0) out.avLst = guides;
    }
    return out;
  }
  const custEl = findChild(parent, A('custGeom'));
  if (custEl) {
    const out: CustomGeometry = { kind: 'custom', pathLst: [] };
    const avLstEl = findChild(custEl, A('avLst'));
    if (avLstEl) {
      const g = parseShapeGuides(avLstEl);
      if (g.length > 0) out.avLst = g;
    }
    const gdLstEl = findChild(custEl, A('gdLst'));
    if (gdLstEl) {
      const g = parseShapeGuides(gdLstEl);
      if (g.length > 0) out.gdLst = g;
    }
    const ahLstEl = findChild(custEl, A('ahLst'));
    if (ahLstEl) {
      const ah = parseAdjustHandles(ahLstEl);
      if (ah.length > 0) out.ahLst = ah;
    }
    const cxnLstEl = findChild(custEl, A('cxnLst'));
    if (cxnLstEl) {
      const cxn = parseConnectionSites(cxnLstEl);
      if (cxn.length > 0) out.cxnLst = cxn;
    }
    const rectEl = findChild(custEl, A('rect'));
    if (rectEl) out.rect = parseRectGuide(rectEl);
    const pathLstEl = findChild(custEl, A('pathLst'));
    if (pathLstEl) {
      for (const p of findChildren(pathLstEl, A('path'))) out.pathLst.push(parsePath(p));
    }
    return out;
  }
  return undefined;
};

export const serializeGeometry = (g: Geometry): string => {
  if (g.kind === 'preset') {
    const av = g.avLst ? serializeShapeGuides('a:avLst', g.avLst) : '';
    return av === ''
      ? `<a:prstGeom prst="${escapeAttr(g.prst)}"/>`
      : `<a:prstGeom prst="${escapeAttr(g.prst)}">${av}</a:prstGeom>`;
  }
  const parts: string[] = ['<a:custGeom>'];
  parts.push(g.avLst ? serializeShapeGuides('a:avLst', g.avLst) : '<a:avLst/>');
  parts.push(g.gdLst ? serializeShapeGuides('a:gdLst', g.gdLst) : '<a:gdLst/>');
  parts.push(g.ahLst ? serializeAdjustHandles(g.ahLst) : '<a:ahLst/>');
  parts.push(g.cxnLst ? serializeConnectionSites(g.cxnLst) : '<a:cxnLst/>');
  if (g.rect) {
    parts.push(
      `<a:rect l="${escapeAttr(g.rect.l)}" t="${escapeAttr(g.rect.t)}" r="${escapeAttr(g.rect.r)}" b="${escapeAttr(g.rect.b)}"/>`,
    );
  }
  parts.push(`<a:pathLst>${g.pathLst.map(serializePath).join('')}</a:pathLst>`);
  parts.push('</a:custGeom>');
  return parts.join('');
};

// ---- Effects ---------------------------------------------------------------

const VALID_BLEND: ReadonlyArray<string> = ['over', 'mult', 'screen', 'darken', 'lighten'];
const VALID_SHADOW_ALIGN: ReadonlyArray<string> = ['tl', 't', 'tr', 'l', 'ctr', 'r', 'bl', 'b', 'br'];
const PRESET_SHADOW_SET: ReadonlySet<string> = new Set(PRESET_SHADOW_NAMES);

const EFFECT_LEAF_NAMES = new Set([
  'blur',
  'fillOverlay',
  'glow',
  'innerShdw',
  'outerShdw',
  'prstShdw',
  'reflection',
  'softEdge',
]);

const parseEffectLeaf = (el: XmlNode): Effect | undefined => {
  const local = el.name.split('}').pop() ?? el.name;
  switch (local) {
    case 'blur': {
      const rad = intAttr(el, 'rad') ?? 0;
      const grow = boolAttr(el, 'grow');
      return { kind: 'blur', rad, ...(grow !== undefined ? { grow } : {}) };
    }
    case 'fillOverlay': {
      const blendRaw = el.attrs['blend'];
      const blend: FillBlendMode = VALID_BLEND.includes(blendRaw ?? '') ? (blendRaw as FillBlendMode) : 'over';
      const fill = parseFill(el);
      if (!fill) return undefined;
      return { kind: 'fillOverlay', blend, fill };
    }
    case 'glow': {
      const rad = intAttr(el, 'rad') ?? 0;
      const color = parseDmlColor(el);
      if (!color) return undefined;
      return { kind: 'glow', rad, color };
    }
    case 'innerShdw': {
      const color = parseDmlColor(el);
      if (!color) return undefined;
      const blurRad = intAttr(el, 'blurRad');
      const dist = intAttr(el, 'dist');
      const dir = intAttr(el, 'dir');
      return {
        kind: 'innerShdw',
        color,
        ...(blurRad !== undefined ? { blurRad } : {}),
        ...(dist !== undefined ? { dist } : {}),
        ...(dir !== undefined ? { dir } : {}),
      };
    }
    case 'outerShdw': {
      const color = parseDmlColor(el);
      if (!color) return undefined;
      const blurRad = intAttr(el, 'blurRad');
      const dist = intAttr(el, 'dist');
      const dir = intAttr(el, 'dir');
      const sx = intAttr(el, 'sx');
      const sy = intAttr(el, 'sy');
      const kx = intAttr(el, 'kx');
      const ky = intAttr(el, 'ky');
      const algnRaw = el.attrs['algn'];
      const algn: ShadowAlign | undefined =
        algnRaw && VALID_SHADOW_ALIGN.includes(algnRaw) ? (algnRaw as ShadowAlign) : undefined;
      const rotWithShape = boolAttr(el, 'rotWithShape');
      return {
        kind: 'outerShdw',
        color,
        ...(blurRad !== undefined ? { blurRad } : {}),
        ...(dist !== undefined ? { dist } : {}),
        ...(dir !== undefined ? { dir } : {}),
        ...(sx !== undefined ? { sx } : {}),
        ...(sy !== undefined ? { sy } : {}),
        ...(kx !== undefined ? { kx } : {}),
        ...(ky !== undefined ? { ky } : {}),
        ...(algn ? { algn } : {}),
        ...(rotWithShape !== undefined ? { rotWithShape } : {}),
      };
    }
    case 'prstShdw': {
      const color = parseDmlColor(el);
      if (!color) return undefined;
      const prstRaw = el.attrs['prst'];
      if (!prstRaw || !PRESET_SHADOW_SET.has(prstRaw)) return undefined;
      const dist = intAttr(el, 'dist');
      const dir = intAttr(el, 'dir');
      return {
        kind: 'prstShdw',
        prst: prstRaw as PresetShadowName,
        color,
        ...(dist !== undefined ? { dist } : {}),
        ...(dir !== undefined ? { dir } : {}),
      };
    }
    case 'reflection': {
      const blurRad = intAttr(el, 'blurRad');
      const stA = intAttr(el, 'stA');
      const stPos = intAttr(el, 'stPos');
      const endA = intAttr(el, 'endA');
      const endPos = intAttr(el, 'endPos');
      const dist = intAttr(el, 'dist');
      const dir = intAttr(el, 'dir');
      const fadeDir = intAttr(el, 'fadeDir');
      const sx = intAttr(el, 'sx');
      const sy = intAttr(el, 'sy');
      const kx = intAttr(el, 'kx');
      const ky = intAttr(el, 'ky');
      const algnRaw = el.attrs['algn'];
      const algn: ShadowAlign | undefined =
        algnRaw && VALID_SHADOW_ALIGN.includes(algnRaw) ? (algnRaw as ShadowAlign) : undefined;
      const rotWithShape = boolAttr(el, 'rotWithShape');
      return {
        kind: 'reflection',
        ...(blurRad !== undefined ? { blurRad } : {}),
        ...(stA !== undefined ? { stA } : {}),
        ...(stPos !== undefined ? { stPos } : {}),
        ...(endA !== undefined ? { endA } : {}),
        ...(endPos !== undefined ? { endPos } : {}),
        ...(dist !== undefined ? { dist } : {}),
        ...(dir !== undefined ? { dir } : {}),
        ...(fadeDir !== undefined ? { fadeDir } : {}),
        ...(sx !== undefined ? { sx } : {}),
        ...(sy !== undefined ? { sy } : {}),
        ...(kx !== undefined ? { kx } : {}),
        ...(ky !== undefined ? { ky } : {}),
        ...(algn ? { algn } : {}),
        ...(rotWithShape !== undefined ? { rotWithShape } : {}),
      };
    }
    case 'softEdge':
      return { kind: 'softEdge', rad: intAttr(el, 'rad') ?? 0 };
    default:
      return undefined;
  }
};

const parseEffectContainerChildren = (parent: XmlNode): Array<Effect | EffectContainer> => {
  const out: Array<Effect | EffectContainer> = [];
  for (const c of parent.children) {
    if (typeof c === 'string') continue;
    const local = c.name.split('}').pop() ?? c.name;
    if (local === 'cont') {
      out.push(parseEffectContainer(c));
    } else if (EFFECT_LEAF_NAMES.has(local)) {
      const leaf = parseEffectLeaf(c);
      if (leaf) out.push(leaf);
    }
  }
  return out;
};

const parseEffectContainer = (el: XmlNode): EffectContainer => {
  const typeRaw = el.attrs['type'];
  const type: 'tree' | 'sib' = typeRaw === 'tree' ? 'tree' : 'sib';
  const name = el.attrs['name'];
  return {
    type,
    ...(name !== undefined ? { name } : {}),
    children: parseEffectContainerChildren(el),
  };
};

const parseEffectList = (el: XmlNode): EffectList => {
  const list: Effect[] = [];
  for (const c of el.children) {
    if (typeof c === 'string') continue;
    const leaf = parseEffectLeaf(c);
    if (leaf) list.push(leaf);
  }
  return { list };
};

export const parseEffects = (parent: XmlNode): EffectsRef | undefined => {
  const lstEl = findChild(parent, A('effectLst'));
  if (lstEl) return { kind: 'lst', list: parseEffectList(lstEl) };
  const dagEl = findChild(parent, A('effectDag'));
  if (dagEl) return { kind: 'dag', children: parseEffectContainerChildren(dagEl) };
  return undefined;
};

const serializeEffectLeaf = (e: Effect): string => {
  switch (e.kind) {
    case 'blur': {
      const grow = e.grow !== undefined ? ` grow="${e.grow ? '1' : '0'}"` : '';
      return `<a:blur rad="${e.rad}"${grow}/>`;
    }
    case 'fillOverlay':
      return `<a:fillOverlay blend="${e.blend}">${serializeFill(e.fill)}</a:fillOverlay>`;
    case 'glow':
      return `<a:glow rad="${e.rad}">${serializeDmlColor(e.color)}</a:glow>`;
    case 'innerShdw': {
      const a: string[] = [];
      if (e.blurRad !== undefined) a.push(`blurRad="${e.blurRad}"`);
      if (e.dist !== undefined) a.push(`dist="${e.dist}"`);
      if (e.dir !== undefined) a.push(`dir="${e.dir}"`);
      return `<a:innerShdw${a.length > 0 ? ` ${a.join(' ')}` : ''}>${serializeDmlColor(e.color)}</a:innerShdw>`;
    }
    case 'outerShdw': {
      const a: string[] = [];
      if (e.blurRad !== undefined) a.push(`blurRad="${e.blurRad}"`);
      if (e.dist !== undefined) a.push(`dist="${e.dist}"`);
      if (e.dir !== undefined) a.push(`dir="${e.dir}"`);
      if (e.sx !== undefined) a.push(`sx="${e.sx}"`);
      if (e.sy !== undefined) a.push(`sy="${e.sy}"`);
      if (e.kx !== undefined) a.push(`kx="${e.kx}"`);
      if (e.ky !== undefined) a.push(`ky="${e.ky}"`);
      if (e.algn) a.push(`algn="${e.algn}"`);
      if (e.rotWithShape !== undefined) a.push(`rotWithShape="${e.rotWithShape ? '1' : '0'}"`);
      return `<a:outerShdw${a.length > 0 ? ` ${a.join(' ')}` : ''}>${serializeDmlColor(e.color)}</a:outerShdw>`;
    }
    case 'prstShdw': {
      const a: string[] = [`prst="${e.prst}"`];
      if (e.dist !== undefined) a.push(`dist="${e.dist}"`);
      if (e.dir !== undefined) a.push(`dir="${e.dir}"`);
      return `<a:prstShdw ${a.join(' ')}>${serializeDmlColor(e.color)}</a:prstShdw>`;
    }
    case 'reflection': {
      const a: string[] = [];
      if (e.blurRad !== undefined) a.push(`blurRad="${e.blurRad}"`);
      if (e.stA !== undefined) a.push(`stA="${e.stA}"`);
      if (e.stPos !== undefined) a.push(`stPos="${e.stPos}"`);
      if (e.endA !== undefined) a.push(`endA="${e.endA}"`);
      if (e.endPos !== undefined) a.push(`endPos="${e.endPos}"`);
      if (e.dist !== undefined) a.push(`dist="${e.dist}"`);
      if (e.dir !== undefined) a.push(`dir="${e.dir}"`);
      if (e.fadeDir !== undefined) a.push(`fadeDir="${e.fadeDir}"`);
      if (e.sx !== undefined) a.push(`sx="${e.sx}"`);
      if (e.sy !== undefined) a.push(`sy="${e.sy}"`);
      if (e.kx !== undefined) a.push(`kx="${e.kx}"`);
      if (e.ky !== undefined) a.push(`ky="${e.ky}"`);
      if (e.algn) a.push(`algn="${e.algn}"`);
      if (e.rotWithShape !== undefined) a.push(`rotWithShape="${e.rotWithShape ? '1' : '0'}"`);
      return `<a:reflection${a.length > 0 ? ` ${a.join(' ')}` : ''}/>`;
    }
    case 'softEdge':
      return `<a:softEdge rad="${e.rad}"/>`;
  }
};

const serializeEffectContainer = (c: EffectContainer): string => {
  const a: string[] = [`type="${c.type}"`];
  if (c.name !== undefined) a.push(`name="${escapeAttr(c.name)}"`);
  const inner = c.children
    .map((child) =>
      'kind' in child ? serializeEffectLeaf(child as Effect) : serializeEffectContainer(child as EffectContainer),
    )
    .join('');
  return `<a:cont ${a.join(' ')}>${inner}</a:cont>`;
};

export const serializeEffects = (e: EffectsRef): string => {
  if (e.kind === 'lst') {
    if (e.list.list.length === 0) return '<a:effectLst/>';
    return `<a:effectLst>${e.list.list.map(serializeEffectLeaf).join('')}</a:effectLst>`;
  }
  if (e.children.length === 0) return '<a:effectDag/>';
  const inner = e.children
    .map((child) =>
      'kind' in child ? serializeEffectLeaf(child as Effect) : serializeEffectContainer(child as EffectContainer),
    )
    .join('');
  return `<a:effectDag>${inner}</a:effectDag>`;
};

// ---- ShapeProperties -------------------------------------------------------

const VALID_BWMODE: ReadonlyArray<string> = [
  'clr',
  'auto',
  'gray',
  'ltGray',
  'invGray',
  'grayWhite',
  'blackGray',
  'blackWhite',
  'black',
  'white',
  'hidden',
];

const parsePoint = (el: XmlNode): Point2D => ({
  x: intAttr(el, 'x') ?? 0,
  y: intAttr(el, 'y') ?? 0,
});

const parseExt = (el: XmlNode): PositiveSize2D => ({
  cx: intAttr(el, 'cx') ?? 0,
  cy: intAttr(el, 'cy') ?? 0,
});

export const parseTransform2D = (el: XmlNode): Transform2D => {
  const out: Transform2D = {};
  const off = findChild(el, A('off'));
  const ext = findChild(el, A('ext'));
  const chOff = findChild(el, A('chOff'));
  const chExt = findChild(el, A('chExt'));
  if (off) out.off = parsePoint(off);
  if (ext) out.ext = parseExt(ext);
  if (chOff) out.chOff = parsePoint(chOff);
  if (chExt) out.chExt = parseExt(chExt);
  const rot = intAttr(el, 'rot');
  const flipH = boolAttr(el, 'flipH');
  const flipV = boolAttr(el, 'flipV');
  if (rot !== undefined) out.rot = rot;
  if (flipH !== undefined) out.flipH = flipH;
  if (flipV !== undefined) out.flipV = flipV;
  return out;
};

const serializeTransform2D = (x: Transform2D): string => {
  const a: string[] = [];
  if (x.rot !== undefined) a.push(`rot="${x.rot}"`);
  if (x.flipH !== undefined) a.push(`flipH="${x.flipH ? '1' : '0'}"`);
  if (x.flipV !== undefined) a.push(`flipV="${x.flipV ? '1' : '0'}"`);
  const parts: string[] = [`<a:xfrm${a.length > 0 ? ` ${a.join(' ')}` : ''}>`];
  if (x.off) parts.push(`<a:off x="${x.off.x}" y="${x.off.y}"/>`);
  if (x.ext) parts.push(`<a:ext cx="${x.ext.cx}" cy="${x.ext.cy}"/>`);
  if (x.chOff) parts.push(`<a:chOff x="${x.chOff.x}" y="${x.chOff.y}"/>`);
  if (x.chExt) parts.push(`<a:chExt cx="${x.chExt.cx}" cy="${x.chExt.cy}"/>`);
  parts.push('</a:xfrm>');
  return parts.join('');
};

/**
 * Parse a `<spPr>` (or `<c:spPr>` / `<cdr:spPr>`) element. Caller provides
 * the element directly — we don't filter by tag name so the helper works
 * for any namespace prefix.
 */
export const parseShapeProperties = (el: XmlNode): ShapeProperties => {
  const out: ShapeProperties = {};
  const bwModeRaw = el.attrs['bwMode'];
  if (bwModeRaw && VALID_BWMODE.includes(bwModeRaw)) out.bwMode = bwModeRaw as BlackWhiteMode;
  const xfrm = findChild(el, A('xfrm'));
  if (xfrm) out.xfrm = parseTransform2D(xfrm);
  const geom = parseGeometry(el);
  if (geom) out.geometry = geom;
  const fill = parseFill(el);
  if (fill) out.fill = fill;
  const ln = findChild(el, A('ln'));
  if (ln) out.ln = parseLine(ln);
  const effects = parseEffects(el);
  if (effects) out.effects = effects;
  return out;
};

/** Serialize a ShapeProperties as `<wrapper>...</wrapper>`. Wrapper defaults to `<c:spPr>`. */
export const serializeShapeProperties = (sp: ShapeProperties, wrapperTag = 'c:spPr'): string => {
  const a: string[] = [];
  if (sp.bwMode) a.push(`bwMode="${sp.bwMode}"`);
  const parts: string[] = [`<${wrapperTag}${a.length > 0 ? ` ${a.join(' ')}` : ''}>`];
  if (sp.xfrm) parts.push(serializeTransform2D(sp.xfrm));
  if (sp.geometry) parts.push(serializeGeometry(sp.geometry));
  if (sp.fill) parts.push(serializeFill(sp.fill));
  if (sp.ln) parts.push(serializeLine(sp.ln));
  if (sp.effects) parts.push(serializeEffects(sp.effects));
  parts.push(`</${wrapperTag}>`);
  return parts.join('');
};

/** Sanity helper: parse + serialize a ShapeProperties from a raw `<a:spPr>` style XML string. */
export function _ensureShapeProperties(node: XmlNode): ShapeProperties {
  if (!node) throw new OpenXmlSchemaError('parseShapeProperties: node is undefined');
  return parseShapeProperties(node);
}
