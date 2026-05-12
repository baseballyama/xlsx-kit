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
      // High surrogate: paired with the next code unit encodes one 4-byte
      // codepoint. A lone high surrogate at the end of the string still
      // produces 3 replacement bytes through TextEncoder; treat it as 4 here
      // so the threshold check stays conservative.
      n += 4;
      i++;
    } else {
      n += 3;
    }
  }
  return n;
}
