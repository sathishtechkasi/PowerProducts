export const systemFields = ['Attachments','ContentType','Created','Modified','Author','Editor'];
export const stripBraces = (g: string) => g.replace(/[{}]/g, '');
export const escapeOdataValue = (v: string) => String(v).replace(/'/g, "''");
export const norm = (s: string) => (s || '').toLowerCase().replace(/\s+/g, '').replace(/[_-]+/g, '');
// Clean display values: trim, normalize spaces (NBSP→space)
export const clean = (s: any): string => {
  if (s === null || s === undefined) return '';
  return String(s).replace(/\u00A0/g, ' ').replace(/\s+/g, ' ').trim();
};
// Choice tokenization with SharePoint-friendly separators
export const splitChoiceCell = (
  raw: any,
  isMulti: boolean,
  allowedList?: string[]
): string[] => {
  const val = clean(raw);
  if (!val) return [];
  if (!isMulti) return [val];
  const allowedHasComma = (allowedList || []).some(x => x.indexOf(',') >= 0);
  if (val.indexOf(';#') >= 0) return val.split(/;#/).map(clean).filter(Boolean);
  if (val.indexOf(';')  >= 0) return val.split(';').map(clean).filter(Boolean);
  if (!allowedHasComma && val.indexOf(',') >= 0) return val.split(',').map(clean).filter(Boolean);
  return [val];
};
// Unwrap SharePoint “Choices” shapes
export const toChoicesArray = (choices: any): string[] => {
  if (!choices) return [];
  if (Array.isArray(choices)) return choices;
  if (choices && Array.isArray(choices.results)) return choices.results;
  return [];
};
