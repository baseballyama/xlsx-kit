// Fast UTF-8 byte-length scan. Used by streaming writers to decide when their
// pending string buffer should be encoded + flushed.
//
// `s.length` returns UTF-16 code units, which undercounts BMP characters above
// U+007F (1 code unit, 2 UTF-8 bytes for U+0080–U+07FF, 3 bytes for the rest
// of the BMP). For CJK-heavy payloads the discrepancy is ~3× — large enough
// to push a "64 KB" flush threshold to 192 KB of resident text. We scan the
// string once and account for each codepoint instead of running a full
// TextEncoder, which would also have to materialise the byte buffer we don't
// need yet.
export function utf8ByteLength(s: string): number {
  let n = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c < 0x80) {
      n += 1;
    } else if (c < 0x800) {
      n += 2;
    } else if (c >= 0xd800 && c <= 0xdbff) {
      // High surrogate. Only consume the next code unit when it actually
      // is a low surrogate — an unpaired high surrogate (followed by a
      // BMP char or by EOS) would otherwise swallow the next code unit
      // and undercount the string. TextEncoder replaces a lone high
      // surrogate with U+FFFD (3 bytes), so use 3 here too. A paired
      // surrogate encodes one 4-byte codepoint.
      const next = i + 1 < s.length ? s.charCodeAt(i + 1) : 0;
      if (next >= 0xdc00 && next <= 0xdfff) {
        n += 4;
        i++;
      } else {
        n += 3;
      }
    } else {
      // Includes unpaired low surrogates (0xDC00-0xDFFF), which encode as
      // U+FFFD = 3 UTF-8 bytes through TextEncoder.
      n += 3;
    }
  }
  return n;
}
