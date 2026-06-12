import ExcelJS from 'exceljs';
import { canonicalizeHeader, excelCellValueToString } from './normalize.js';
import type { SkuMappingIndex } from './types.js';

const STREAM_READER_OPTIONS: Partial<ExcelJS.stream.xlsx.WorkbookStreamReaderOptions> = {
  sharedStrings: 'cache',
  hyperlinks: 'ignore',
  styles: 'ignore',
};

const SKU_MAPPING_HEADER_ROW = 1;

const META_FIELDS = new Set([
  'id',
  'dtg_print_sku',
  'product_name',
  'color',
  'size',
  'type',
  'print_tech',
]);

function worksheetName(reader: ExcelJS.stream.xlsx.WorksheetReader): string {
  return String((reader as { name?: string }).name ?? '').trim();
}

export function isSkuMappingSheetName(name: string): boolean {
  const upper = name.trim().toUpperCase();
  return upper === 'SKU MAPPING' || upper === 'SKU_MAPPING';
}

function rowToCells(row: ExcelJS.Row): unknown[] {
  const cells: unknown[] = [];
  row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
    while (cells.length < colNumber - 1) cells.push('');
    cells[colNumber - 1] = excelCellValueToString(cell.value);
  });
  return cells;
}

function buildColToField(headerCells: unknown[]): Map<number, string> {
  const colToField = new Map<number, string>();
  for (let i = 0; i < headerCells.length; i++) {
    const field = canonicalizeHeader(excelCellValueToString(headerCells[i]));
    if (field) colToField.set(i, field);
  }
  return colToField;
}

function cellText(cells: unknown[], colIndex: number): string {
  return excelCellValueToString(cells[colIndex] ?? '').trim();
}

async function drainWorksheetRows(
  worksheetReader: ExcelJS.stream.xlsx.WorksheetReader,
): Promise<void> {
  for await (const _row of worksheetReader) {
    // Bỏ qua sheet không cần đọc.
  }
}

function parseSkuMappingRows(
  rows: Array<{ rowNumber: number; cells: unknown[] }>,
): SkuMappingIndex {
  const headerRow = rows.find((r) => r.rowNumber === SKU_MAPPING_HEADER_ROW);
  if (!headerRow) {
    throw new Error(
      `Sheet SKU MAPPING thiếu dòng header (dòng ${SKU_MAPPING_HEADER_ROW}).`,
    );
  }

  const colToField = buildColToField(headerRow.cells);
  let dtgPrintSkuCol: number | null = null;
  const factoryCols = new Map<number, string>();
  const byFactoryId = new Map<string, Map<string, string[]>>();
  const factories = new Set<string>();

  for (const [colIndex, field] of colToField.entries()) {
    if (field === 'dtg_print_sku') {
      dtgPrintSkuCol = colIndex;
      continue;
    }
    if (META_FIELDS.has(field)) continue;

    factoryCols.set(colIndex, field);
    factories.add(field);
    if (!byFactoryId.has(field)) byFactoryId.set(field, new Map());
  }

  if (dtgPrintSkuCol === null) {
    throw new Error('Sheet SKU MAPPING thiếu cột "DTG Print SKU".');
  }
  if (factoryCols.size === 0) {
    throw new Error('Sheet SKU MAPPING không có cột xưởng (NB1, LK401, ...).');
  }

  for (const row of rows) {
    if (row.rowNumber <= SKU_MAPPING_HEADER_ROW) continue;

    const dtgPrintSku = cellText(row.cells, dtgPrintSkuCol);
    if (!dtgPrintSku) continue;

    for (const [colIndex, factory] of factoryCols.entries()) {
      const factoryId = cellText(row.cells, colIndex);
      if (!factoryId) continue;

      const idMap = byFactoryId.get(factory)!;
      const existing = idMap.get(factoryId);
      if (existing) {
        if (!existing.includes(dtgPrintSku)) existing.push(dtgPrintSku);
      } else {
        idMap.set(factoryId, [dtgPrintSku]);
      }
    }
  }

  return { byFactoryId, factories };
}

export function resolveDtgPrintSkus(
  index: SkuMappingIndex,
  factory: string,
  factoryRowId: string,
): string[] {
  const idMap = index.byFactoryId.get(factory.trim().toLowerCase());
  if (!idMap) return [];

  const id = factoryRowId.trim();
  if (!id) return [];

  const results: string[] = [];
  const seen = new Set<string>();

  const addAll = (skus: string[]) => {
    for (const sku of skus) {
      if (!seen.has(sku)) {
        seen.add(sku);
        results.push(sku);
      }
    }
  };

  const direct = idMap.get(id);
  if (direct) addAll(direct);

  const lowerId = id.toLowerCase();
  for (const [key, skus] of idMap.entries()) {
    if (key !== id && key.toLowerCase() === lowerId) addAll(skus);
  }

  return results;
}

export async function parseSkuMapping(filePath: string): Promise<SkuMappingIndex> {
  const reader = new ExcelJS.stream.xlsx.WorkbookReader(filePath, STREAM_READER_OPTIONS);
  let mappingRows: Array<{ rowNumber: number; cells: unknown[] }> | null = null;

  for await (const worksheetReader of reader) {
    const sheetName = worksheetName(worksheetReader);

    if (!isSkuMappingSheetName(sheetName)) {
      await drainWorksheetRows(worksheetReader);
      continue;
    }

    mappingRows = [];
    for await (const row of worksheetReader) {
      mappingRows.push({ rowNumber: row.number, cells: rowToCells(row) });
    }
    break;
  }

  if (!mappingRows) {
    throw new Error('Không tìm thấy sheet "SKU MAPPING" trong file Excel.');
  }

  return parseSkuMappingRows(mappingRows);
}
