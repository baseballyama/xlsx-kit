// Render typedoc Type and Reflection nodes as TS-ish source. Not a full
// TypeScript printer — best-effort coverage of the shapes that show up
// in this codebase's public surface.

/* eslint-disable @typescript-eslint/no-explicit-any */
type Type = any;
type Reflection = any;

const KEYWORDS_NEED_PARENS = new Set(['union', 'intersection', 'conditional']);

export function renderType(t: Type): string {
  if (!t) return 'unknown';
  switch (t.type) {
    case 'intrinsic':
      return t.name ?? 'unknown';
    case 'literal':
      if (typeof t.value === 'string') return JSON.stringify(t.value);
      if (t.value === null) return 'null';
      return String(t.value);
    case 'reference': {
      const args = t.typeArguments?.length
        ? `<${t.typeArguments.map(renderType).join(', ')}>`
        : '';
      return (t.name ?? 'unknown') + args;
    }
    case 'array': {
      const inner = renderType(t.elementType);
      const needsParens = KEYWORDS_NEED_PARENS.has(t.elementType?.type ?? '');
      return needsParens ? `(${inner})[]` : `${inner}[]`;
    }
    case 'union':
      return (t.types ?? []).map(renderType).join(' | ');
    case 'intersection':
      return (t.types ?? []).map(renderType).join(' & ');
    case 'tuple':
      return `[${(t.elements ?? []).map(renderType).join(', ')}]`;
    case 'reflection':
      return renderReflection(t.declaration);
    case 'mapped':
      return `{ [${t.parameter ?? 'K'} in ${renderType(t.parameterType)}]${
        t.optionalModifier === '+' ? '?' : t.optionalModifier === '-' ? '-?' : ''
      }: ${renderType(t.templateType)} }`;
    case 'predicate':
      return t.targetType
        ? `${t.name} is ${renderType(t.targetType)}`
        : 'boolean';
    case 'query':
      return `typeof ${renderType(t.queryType)}`;
    case 'conditional':
      return `${renderType(t.checkType)} extends ${renderType(t.extendsType)} ? ${renderType(
        t.trueType,
      )} : ${renderType(t.falseType)}`;
    case 'indexedAccess':
      return `${renderType(t.objectType)}[${renderType(t.indexType)}]`;
    case 'typeOperator':
      return `${t.operator} ${renderType(t.target)}`;
    case 'rest':
      return `...${renderType(t.elementType)}`;
    case 'optional':
      return `${renderType(t.elementType)}?`;
    case 'named-tuple-member': {
      const opt = t.isOptional ? '?' : '';
      return `${t.name}${opt}: ${renderType(t.element)}`;
    }
    case 'unknown':
      return t.name ?? 'unknown';
    default:
      return '...';
  }
}

export function renderReflection(decl: Reflection | undefined): string {
  if (!decl) return '{}';
  // Function shape (callable, no own children): (a: T) => U
  if (decl.signatures?.length && !(decl.children?.length ?? 0)) {
    const sig = decl.signatures[0];
    return renderSignature(sig, { arrow: true });
  }
  // Object literal
  const props = (decl.children ?? []).map((c: Reflection) => {
    const opt = c.flags?.isOptional ? '?' : '';
    return `${c.name}${opt}: ${renderType(c.type)}`;
  });
  if (props.length === 0) return '{}';
  return `{ ${props.join('; ')} }`;
}

export function renderSignature(
  sig: Reflection,
  opts: { arrow?: boolean; name?: string } = {},
): string {
  const params = (sig.parameters ?? []).map(renderParam).join(', ');
  const ret = renderType(sig.type);
  if (opts.arrow) return `(${params}) => ${ret}`;
  const typeParams = (sig.typeParameter ?? []).map((tp: Reflection) => {
    const constraint = tp.type ? ` extends ${renderType(tp.type)}` : '';
    const def = tp.default ? ` = ${renderType(tp.default)}` : '';
    return `${tp.name}${constraint}${def}`;
  });
  const tp = typeParams.length ? `<${typeParams.join(', ')}>` : '';
  return `${opts.name ?? sig.name}${tp}(${params}): ${ret}`;
}

export function renderParam(p: Reflection): string {
  const opt = p.flags?.isOptional || p.defaultValue !== undefined ? '?' : '';
  const rest = p.flags?.isRest ? '...' : '';
  return `${rest}${p.name}${opt}: ${renderType(p.type)}`;
}

/** Concatenate a typedoc comment.summary into plain text. Inline @link
 *  tags collapse to their display text — links are resolved separately
 *  by the parser when it has the cross-id index. */
export function summaryToText(summary: Reflection[] | undefined): string {
  if (!summary?.length) return '';
  return summary
    .map((s: Reflection) => {
      if (s.kind === 'text' || s.kind === 'code') return s.text ?? '';
      if (s.kind === 'inline-tag') return s.text ?? '';
      return '';
    })
    .join('');
}
