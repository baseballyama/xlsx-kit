import { describe, expect, it } from 'vitest';
import {
  OpenXmlError,
  OpenXmlInvalidWorkbookError,
  OpenXmlIoError,
  OpenXmlNotImplementedError,
  OpenXmlSchemaError,
} from '../../../src/utils/exceptions';

describe('OpenXmlError hierarchy', () => {
  it('OpenXmlError extends Error and carries its name', () => {
    const e = new OpenXmlError('boom');
    expect(e).toBeInstanceOf(Error);
    expect(e).toBeInstanceOf(OpenXmlError);
    expect(e.name).toBe('OpenXmlError');
    expect(e.message).toBe('boom');
  });

  it('subclasses set their own name', () => {
    expect(new OpenXmlIoError('x').name).toBe('OpenXmlIoError');
    expect(new OpenXmlSchemaError('x').name).toBe('OpenXmlSchemaError');
    expect(new OpenXmlInvalidWorkbookError('x').name).toBe('OpenXmlInvalidWorkbookError');
    expect(new OpenXmlNotImplementedError('x').name).toBe('OpenXmlNotImplementedError');
  });

  it('subclasses are instanceof OpenXmlError', () => {
    expect(new OpenXmlIoError('x')).toBeInstanceOf(OpenXmlError);
    expect(new OpenXmlSchemaError('x')).toBeInstanceOf(OpenXmlError);
  });

  it('preserves a cause via the standard Error.cause property', () => {
    const root = new TypeError('underlying');
    const e = new OpenXmlIoError('wrap', { cause: root });
    expect(e.cause).toBe(root);
  });
});
