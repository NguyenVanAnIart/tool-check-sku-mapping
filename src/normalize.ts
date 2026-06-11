export function excelCellValueToString(v: unknown): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number' || typeof v === 'boolean') return String(v);
  if (v instanceof Date) return v.toISOString();

  if (typeof v === 'object') {
    const obj = v as Record<string, unknown>;

    if (typeof obj.text === 'string') return obj.text;
    if ('result' in obj) return excelCellValueToString(obj.result);
    if (Array.isArray(obj.richText)) {
      return (obj.richText as Array<{ text?: string }>)
        .map((part) => part.text ?? '')
        .join('');
    }
    if (typeof obj.error === 'string') return obj.error;
    if (typeof obj.hyperlink === 'string') return obj.hyperlink;
  }

  return String(v);
}

export function normalizeHeaderCell(v: unknown): string {
  return excelCellValueToString(v)
    .trim()
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

export function canonicalizeHeader(h: string): string {
  const s = normalizeHeaderCell(h);
  if (s === 'product name') return 'product_name';
  if (s === 'printtech' || s === 'print tech') return 'print_tech';
  if (s === 'flash sale 1') return 'flash_sale_1';
  if (s === 'flash sale 2') return 'flash_sale_2';
  if (s === 'additional price') return 'additional_price';
  return s.replace(/\s+/g, '_');
}

export function toNumberOrNull(v: unknown): number | null {
  const raw = excelCellValueToString(v).trim();
  if (!raw) return null;
  const n = Number(raw.replace(/,/g, ''));
  return Number.isFinite(n) ? n : null;
}

export function toBooleanStock(v: unknown): boolean {
  const s = excelCellValueToString(v).trim().toLowerCase();
  return s === 'y' || s === 'yes' || s === 'true' || s === '1';
}

export function normString(v: unknown, emptyAsNull = false): string | null {
  const s = excelCellValueToString(v).trim();
  if (!s) return emptyAsNull ? null : '';
  return s;
}

export function formatDisplayValue(v: unknown): string {
  if (v === null || v === undefined) return '(empty)';
  if (typeof v === 'boolean') return v ? 'true' : 'false';
  return excelCellValueToString(v);
}

export function numbersEqual(
  a: number | null | undefined,
  b: number | null | undefined,
): boolean {
  const na = a === null || a === undefined ? null : Number(a);
  const nb = b === null || b === undefined ? null : Number(b);
  if (na === null && nb === null) return true;
  if (na === null || nb === null) return false;
  if (!Number.isFinite(na) || !Number.isFinite(nb)) return false;
  return Math.abs(na - nb) < 0.005;
}

export function buildMatchKey(
  factory: string,
  skuFactory: string,
  type: string,
  printTech: string,
  include: string,
): string {
  return [
    factory.trim().toLowerCase(),
    skuFactory.trim(),
    type.trim(),
    printTech.trim(),
    include.trim(),
  ].join('|');
}

export function isFactorySheetHeaderRow(cells: unknown[]): boolean {
  const c0 = normalizeHeaderCell(cells[0]);
  const c1 = normalizeHeaderCell(cells[1]);
  const c2 = normalizeHeaderCell(cells[2]);
  return (
    c0 === 'id' &&
    c1 === 'sku' &&
    (c2 === 'product name' || c2 === 'product_name')
  );
}

export function isLikelyFactorySheetName(name: string): boolean {
  const n = name.trim();
  if (!n) return false;
  const upper = n.toUpperCase();
  const skip = new Set([
    'SKU MAPPING',
    'SKU_MAPPING',
    'README',
    'INSTRUCTIONS',
    'TEMPLATE',
    'CONFIG',
  ]);
  return !skip.has(upper);
}
