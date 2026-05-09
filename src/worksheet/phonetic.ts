// Worksheet-level <phoneticPr> for East-Asian (mostly Japanese) furigana
// rendering.
//
// Excel uses `<phoneticPr fontId="..." type="..." alignment="..."/>` to drive
// how it renders the small phonetic annotation strip above CJK characters in
// cells. The per-cell `<rPh>` annotations live on shared-string entries and are
// a separate concern (richer model).

export type PhoneticType = 'halfwidthKatakana' | 'fullwidthKatakana' | 'Hiragana' | 'noConversion';
export type PhoneticAlignment = 'noControl' | 'left' | 'center' | 'distributed';

export interface WorksheetPhoneticProperties {
  /** Font index in the workbook's stylesheet for the furigana glyphs. */
  fontId?: number;
  /** Conversion mode the IME should default to when adding furigana. */
  type?: PhoneticType;
  /** Horizontal alignment of the furigana strip relative to the base text. */
  alignment?: PhoneticAlignment;
}

export const makeWorksheetPhoneticProperties = (
  opts: WorksheetPhoneticProperties = {},
): WorksheetPhoneticProperties => ({ ...opts });
