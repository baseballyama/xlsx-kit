// DrawingML colors. Per docs/plan/08-charts-drawings.md §4.1.
//
// ECMA-376 spreads "color" across six element kinds (`<a:srgbClr>`,
// `<a:sysClr>`, `<a:schemeClr>`, `<a:prstClr>`, `<a:hslClr>`,
// `<a:scrgbClr>`) plus a long modifier list (`<a:lumMod>`, `<a:tint>`,
// etc.) that may follow the base color. The model splits the two: a
// `DmlColor` discriminated union for the base, and a `ColorMod[]` carrying
// the ordered modifier list. Wrapping them in `DmlColorWithMods` keeps
// the modifier order — Excel re-applies them in the same sequence.

/** ECMA-376 scheme-color names (`<a:schemeClr val="...">`). */
export type SchemeColorName =
  | 'bg1'
  | 'tx1'
  | 'bg2'
  | 'tx2'
  | 'accent1'
  | 'accent2'
  | 'accent3'
  | 'accent4'
  | 'accent5'
  | 'accent6'
  | 'hlink'
  | 'folHlink'
  | 'phClr'
  | 'dk1'
  | 'lt1'
  | 'dk2'
  | 'lt2';

export const SCHEME_COLOR_NAMES: ReadonlyArray<SchemeColorName> = [
  'bg1',
  'tx1',
  'bg2',
  'tx2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
  'phClr',
  'dk1',
  'lt1',
  'dk2',
  'lt2',
];

/** ECMA-376 preset-color names (`<a:prstClr val="...">`). 140 entries (CSS 1 + X11 extensions). */
export const PRESET_COLOR_NAMES: ReadonlyArray<string> = [
  'aliceBlue',
  'antiqueWhite',
  'aqua',
  'aquamarine',
  'azure',
  'beige',
  'bisque',
  'black',
  'blanchedAlmond',
  'blue',
  'blueViolet',
  'brown',
  'burlyWood',
  'cadetBlue',
  'chartreuse',
  'chocolate',
  'coral',
  'cornflowerBlue',
  'cornsilk',
  'crimson',
  'cyan',
  'darkBlue',
  'darkCyan',
  'darkGoldenrod',
  'darkGray',
  'darkGrey',
  'darkGreen',
  'darkKhaki',
  'darkMagenta',
  'darkOliveGreen',
  'darkOrange',
  'darkOrchid',
  'darkRed',
  'darkSalmon',
  'darkSeaGreen',
  'darkSlateBlue',
  'darkSlateGray',
  'darkSlateGrey',
  'darkTurquoise',
  'darkViolet',
  'deepPink',
  'deepSkyBlue',
  'dimGray',
  'dimGrey',
  'dodgerBlue',
  'firebrick',
  'floralWhite',
  'forestGreen',
  'fuchsia',
  'gainsboro',
  'ghostWhite',
  'gold',
  'goldenrod',
  'gray',
  'grey',
  'green',
  'greenYellow',
  'honeydew',
  'hotPink',
  'indianRed',
  'indigo',
  'ivory',
  'khaki',
  'lavender',
  'lavenderBlush',
  'lawnGreen',
  'lemonChiffon',
  'lightBlue',
  'lightCoral',
  'lightCyan',
  'lightGoldenrodYellow',
  'lightGray',
  'lightGrey',
  'lightGreen',
  'lightPink',
  'lightSalmon',
  'lightSeaGreen',
  'lightSkyBlue',
  'lightSlateGray',
  'lightSlateGrey',
  'lightSteelBlue',
  'lightYellow',
  'lime',
  'limeGreen',
  'linen',
  'magenta',
  'maroon',
  'medAquamarine',
  'medBlue',
  'medOrchid',
  'medPurple',
  'medSeaGreen',
  'medSlateBlue',
  'medSpringGreen',
  'medTurquoise',
  'medVioletRed',
  'midnightBlue',
  'mintCream',
  'mistyRose',
  'moccasin',
  'navajoWhite',
  'navy',
  'oldLace',
  'olive',
  'oliveDrab',
  'orange',
  'orangeRed',
  'orchid',
  'paleGoldenrod',
  'paleGreen',
  'paleTurquoise',
  'paleVioletRed',
  'papayaWhip',
  'peachPuff',
  'peru',
  'pink',
  'plum',
  'powderBlue',
  'purple',
  'red',
  'rosyBrown',
  'royalBlue',
  'saddleBrown',
  'salmon',
  'sandyBrown',
  'seaGreen',
  'seaShell',
  'sienna',
  'silver',
  'skyBlue',
  'slateBlue',
  'slateGray',
  'slateGrey',
  'snow',
  'springGreen',
  'steelBlue',
  'tan',
  'teal',
  'thistle',
  'tomato',
  'turquoise',
  'violet',
  'wheat',
  'white',
  'whiteSmoke',
  'yellow',
  'yellowGreen',
];

export type DmlColor =
  | { kind: 'srgb'; value: string /* RRGGBB */ }
  | { kind: 'sysClr'; value: string; lastClr?: string }
  | { kind: 'schemeClr'; value: SchemeColorName }
  | { kind: 'prstClr'; value: string }
  | { kind: 'hslClr'; hue: number; sat: number; lum: number }
  | { kind: 'scrgbClr'; r: number; g: number; b: number };

export type ColorMod =
  | { kind: 'lumMod'; val: number }
  | { kind: 'lumOff'; val: number }
  | { kind: 'satMod'; val: number }
  | { kind: 'satOff'; val: number }
  | { kind: 'hueMod'; val: number }
  | { kind: 'hueOff'; val: number }
  | { kind: 'tint'; val: number }
  | { kind: 'shade'; val: number }
  | { kind: 'alpha'; val: number }
  | { kind: 'alphaMod'; val: number }
  | { kind: 'alphaOff'; val: number }
  | { kind: 'red'; val: number }
  | { kind: 'green'; val: number }
  | { kind: 'blue'; val: number }
  | { kind: 'redMod'; val: number }
  | { kind: 'greenMod'; val: number }
  | { kind: 'blueMod'; val: number }
  | { kind: 'redOff'; val: number }
  | { kind: 'greenOff'; val: number }
  | { kind: 'blueOff'; val: number }
  | { kind: 'gray' }
  | { kind: 'comp' }
  | { kind: 'inv' }
  | { kind: 'invGamma' }
  | { kind: 'gamma' };

/** Color mod kinds that take no `val` attribute (their elements are empty). */
export const VALUELESS_COLOR_MOD_KINDS: ReadonlyArray<ColorMod['kind']> = ['gray', 'comp', 'inv', 'invGamma', 'gamma'];

/** Color mod kinds that carry a numeric `val` attribute. */
export const VALUED_COLOR_MOD_KINDS: ReadonlyArray<ColorMod['kind']> = [
  'lumMod',
  'lumOff',
  'satMod',
  'satOff',
  'hueMod',
  'hueOff',
  'tint',
  'shade',
  'alpha',
  'alphaMod',
  'alphaOff',
  'red',
  'green',
  'blue',
  'redMod',
  'greenMod',
  'blueMod',
  'redOff',
  'greenOff',
  'blueOff',
];

export interface DmlColorWithMods {
  base: DmlColor;
  mods: ColorMod[];
}

export const makeSrgbColor = (rrggbb: string): DmlColor => ({ kind: 'srgb', value: rrggbb.toUpperCase() });

export const makeSchemeColor = (name: SchemeColorName): DmlColor => ({ kind: 'schemeClr', value: name });

export const makeColor = (base: DmlColor, mods: ColorMod[] = []): DmlColorWithMods => ({ base, mods });
