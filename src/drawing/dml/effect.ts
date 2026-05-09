// DrawingML effects.
//
// Effects come in two containers: `<a:effectLst>` (a sequential list of effect
// elements) and `<a:effectDag>` (a tree of effect-containers, each composing
// children with type=`tree` or `sib`). The Effect union covers the eight
// ECMA-376 §20.1.8 effect leaf kinds.

import type { DmlColorWithMods } from './colors';
import type { Fill } from './fill';

export type FillBlendMode = 'over' | 'mult' | 'screen' | 'darken' | 'lighten';

export type ShadowAlign = 'tl' | 't' | 'tr' | 'l' | 'ctr' | 'r' | 'bl' | 'b' | 'br';

/** Preset-shadow names (`<a:prstShdw prst="shdwN"/>`, 20 entries). */
export type PresetShadowName =
  | 'shdw1'
  | 'shdw2'
  | 'shdw3'
  | 'shdw4'
  | 'shdw5'
  | 'shdw6'
  | 'shdw7'
  | 'shdw8'
  | 'shdw9'
  | 'shdw10'
  | 'shdw11'
  | 'shdw12'
  | 'shdw13'
  | 'shdw14'
  | 'shdw15'
  | 'shdw16'
  | 'shdw17'
  | 'shdw18'
  | 'shdw19'
  | 'shdw20';

export const PRESET_SHADOW_NAMES: ReadonlyArray<PresetShadowName> = [
  'shdw1',
  'shdw2',
  'shdw3',
  'shdw4',
  'shdw5',
  'shdw6',
  'shdw7',
  'shdw8',
  'shdw9',
  'shdw10',
  'shdw11',
  'shdw12',
  'shdw13',
  'shdw14',
  'shdw15',
  'shdw16',
  'shdw17',
  'shdw18',
  'shdw19',
  'shdw20',
];

export type Effect =
  | { kind: 'blur'; rad: number; grow?: boolean }
  | { kind: 'fillOverlay'; blend: FillBlendMode; fill: Fill }
  | { kind: 'glow'; rad: number; color: DmlColorWithMods }
  | { kind: 'innerShdw'; blurRad?: number; dist?: number; dir?: number; color: DmlColorWithMods }
  | {
      kind: 'outerShdw';
      blurRad?: number;
      dist?: number;
      dir?: number;
      sx?: number;
      sy?: number;
      kx?: number;
      ky?: number;
      algn?: ShadowAlign;
      rotWithShape?: boolean;
      color: DmlColorWithMods;
    }
  | { kind: 'prstShdw'; prst: PresetShadowName; dist?: number; dir?: number; color: DmlColorWithMods }
  | {
      kind: 'reflection';
      blurRad?: number;
      stA?: number;
      stPos?: number;
      endA?: number;
      endPos?: number;
      dist?: number;
      dir?: number;
      fadeDir?: number;
      sx?: number;
      sy?: number;
      kx?: number;
      ky?: number;
      algn?: ShadowAlign;
      rotWithShape?: boolean;
    }
  | { kind: 'softEdge'; rad: number };

/** `<a:effectLst>` — ordered list of effect leaves. */
export interface EffectList {
  list: Effect[];
}

/**
 * Inner `<a:cont type="tree|sib">` node. Lives inside `<a:effectDag>` or inside
 * another `<a:cont>`. Children may be effect leaves or further containers.
 */
export interface EffectContainer {
  /** Compose children sequentially (`sib`) or as a tree (`tree`). */
  type: 'tree' | 'sib';
  /** Optional name for cross-referencing other DAG nodes. */
  name?: string;
  children: Array<Effect | EffectContainer>;
}

/**
 * Either `<a:effectLst>` or `<a:effectDag>`. The dag root has no `type`
 * attribute in ECMA-376; only its `<a:cont>` children do, hence we model the
 * dag as a flat list of children rather than a single EffectContainer.
 */
export type EffectsRef = { kind: 'lst'; list: EffectList } | { kind: 'dag'; children: Array<Effect | EffectContainer> };

export const makeEffectList = (list: Effect[]): EffectList => ({ list });

export const makeEffectContainer = (
  type: 'tree' | 'sib',
  children: Array<Effect | EffectContainer>,
  name?: string,
): EffectContainer => ({
  type,
  ...(name !== undefined ? { name } : {}),
  children,
});
